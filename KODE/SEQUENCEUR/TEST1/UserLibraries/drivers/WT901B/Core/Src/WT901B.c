/*
 * WT901B.c — Driver implementation for the WITMotion WT901B IMU.
 *
 * This source file provides a lightweight parser for the serial data frames
 * produced by a WT901B inertial sensor and helper functions to configure
 * the device via its UART interface. Refer to WT901B.h for a detailed
 * description of the available registers, frame formats and conversion
 * constants. The formulas used here are derived directly from the
 * manufacturer’s standard communication protocol.
 */

#include "WT901B.h"
#include "usart.h"

#include <string.h>

    
/* -------------------------------------------------------------------------- */
/*                        Internal Frame Sub-Parsers                          */
/* -------------------------------------------------------------------------- */

void WT901B_Parse_One_Frame(WT901B_t *wt, const uint8_t frame_ptr[11]);

static void WT901B_parse_time_frame(		const uint8_t data[8], WT901B_TimeFrame_t *frame_obj);
static void WT901B_parse_accel_frame(		const uint8_t data[8], WT901B_AccelFrame_t *frame_obj);
static void WT901B_parse_gyro_frame(		const uint8_t data[8], WT901B_GyroFrame_t *frame_obj);
static void WT901B_parse_angle_frame(		const uint8_t data[8], WT901B_AngleFrame_t *frame_obj);
static void WT901B_parse_mag_frame(			const uint8_t data[8], WT901B_MagFrame_t *frame_obj);
static void WT901B_parse_port_frame(		const uint8_t data[8], WT901B_PortFrame_t *frame_obj);
static void WT901B_parse_pressure_frame(	const uint8_t data[8], WT901B_PressureFrame_t *frame_obj);
static void WT901B_parse_gps_frame(			const uint8_t data[8], WT901B_GPSFrame_t *frame_obj);
static void WT901B_parse_velocity_frame(	const uint8_t data[8], WT901B_VelocityFrame_t *frame_obj);
static void WT901B_parse_quaternion_frame(	const uint8_t data[8], WT901B_QuaternionFrame_t *frame_obj);
static void WT901B_parse_gsa_frame(			const uint8_t data[8], WT901B_GSAFrame_t *frame_obj);



WT901B_t wt901b = { 0 };

/* -------------------------------------------------------------------------- */
/*                            Internal Helper Macros                          */
/* -------------------------------------------------------------------------- */
/**
 * @brief Compute checksum of the first 10 bytes of an 11‑byte frame.
 *
 * The WT901B frame checksum is defined as the low 8 bits of the sum of the
 * header (0x55), the type byte and the eight data bytes【522002904479991†L820-L864】.
 *
 * @param buf Pointer to the 11‑byte frame buffer.
 * @return uint8_t Lower 8 bits of the computed checksum.
 */
static inline uint8_t wt901b_compute_checksum(const uint8_t buf[11]) {
    uint16_t sum = 0U;
    for (uint8_t i = 0U; i < 10U; i++) {
        sum += buf[i];
    }
    return (uint8_t)sum;
}

/* -------------------------------------------------------------------------- */
/*                         Public API Implementation                         */
/* -------------------------------------------------------------------------- */


WT901B_status_t WT901B_Init(WT901B_t *wt, UART_HandleTypeDef *huart) {
    wt->huart = huart;
    memset(wt->uart_buffer, 0, WT901B_FRAME_LENGTH * WT901B_FRAME_TYPE_NBR);
    wt->new_data_available = false;
    wt->data_in_progress = false;
    wt->last_status = WT901B_OK;

    UART_buffer_t *buffer_obj;
    HAL_StatusTypeDef res;
    UART_get_buffer(huart, &buffer_obj);
    if (buffer_obj != NULL) {
        buffer_obj->rx_buffer = wt->uart_buffer;
        buffer_obj->rx_length = sizeof(wt->uart_buffer);
        res = HAL_UARTEx_ReceiveToIdle_IT(huart, buffer_obj->rx_buffer, buffer_obj->rx_length);
        if (res != HAL_OK) {
            wt->last_status = WT901B_UART_ERROR;
            return WT901B_UART_ERROR; // UART receive error
        }
    } else {
        wt->last_status = WT901B_UART_ERROR;
        return WT901B_UART_ERROR; // Unknown UART instance
    }

	data_topic_init(&wt->data_topic, wt->dt_storage, sizeof(WT901B_Frame_t), WT901B_DATA_TOPIC_LENGTH, CB_OVERWRITE_OLDEST);

    return WT901B_OK;
}

void WT901B_UART_Callback_RX_IRQHandler(WT901B_t *wt, uint16_t Size) {
    // Frame received
    if (Size >= WT901B_FRAME_LENGTH) {
        if (!wt->data_in_progress) {
            memcpy(wt->parse_buffer, wt->uart_buffer, Size);
            wt->last_received_size = Size;
            wt->data_in_progress = true;
			wt->last_timestamp_ms = HAL_GetTick();
        } else {
            // Previous frame was not processed yet, data overrun
            wt->last_status = WT901B_OVERRUN_ERROR;
        }
    }
}


void WT901B_Parse_Frames(WT901B_t *wt) {
    if (!wt->data_in_progress || wt->last_received_size < WT901B_FRAME_LENGTH) {
        wt->last_status = WT901B_NO_DATA;
		// Make sure uart interrupts are re-enabled
		HAL_UARTEx_ReceiveToIdle_IT(wt->huart, wt->uart_buffer, sizeof(wt->uart_buffer));
        return; // No new data to parse (or not enought)
    }

    // Get the head of the frames
    for (size_t i = 0; i <= (wt->last_received_size - WT901B_FRAME_LENGTH); i++) {
        if (wt->parse_buffer[i] == WT901B_FRAME_HEADER) {
            // Found potential frame header
            WT901B_Parse_One_Frame(wt, &wt->parse_buffer[i]);
            i += WT901B_FRAME_LENGTH; // Move to next potential frame
            if (i + WT901B_FRAME_LENGTH > wt->last_received_size) {
                break; // No more complete frames
            }
            i--; // Adjust for the loop increment
        }
    }
    wt->data_in_progress = false; // All data processed
    wt->last_status = WT901B_OK;
}

void WT901B_Parse_One_Frame(WT901B_t *wt, const uint8_t frame_ptr[11]) {
    // Verify header
    if (frame_ptr[0] != WT901B_FRAME_HEADER) {
        return; // Invalid header, skip frame
    }

    // Verify checksum
    uint8_t computed_checksum = wt901b_compute_checksum(frame_ptr);
    if (computed_checksum != frame_ptr[10]) {
        wt->last_status = WT901B_CHECKSUM_ERROR;
        return; // Checksum mismatch, skip frame
    }
    wt->last_status = WT901B_OK;

    // Valid frame received, process based on type
    uint8_t frame_type = frame_ptr[1];
    const uint8_t *data = &frame_ptr[2];
    WT901B_Frame_t frame_data = {
		.type = frame_type,
		.timestamp_ms = wt->last_timestamp_ms,
		.data = { 0 }
	};
    switch (frame_type) {
        case WT901B_FRAME_TIME:         { WT901B_parse_time_frame(		data, &frame_data.data.time			); break; }
        case WT901B_FRAME_ACCEL:        { WT901B_parse_accel_frame(		data, &frame_data.data.accel		); break; }
        case WT901B_FRAME_GYRO:         { WT901B_parse_gyro_frame(		data, &frame_data.data.gyro			); break; }
        case WT901B_FRAME_ANGLE:        { WT901B_parse_angle_frame(		data, &frame_data.data.angle		); break; }
        case WT901B_FRAME_MAG:          { WT901B_parse_mag_frame(		data, &frame_data.data.mag			); break; }
        case WT901B_FRAME_PORT:         { WT901B_parse_port_frame(		data, &frame_data.data.port			); break; }
        case WT901B_FRAME_PRESSURE:     { WT901B_parse_pressure_frame(	data, &frame_data.data.pressure		); break; }
        case WT901B_FRAME_GPS:          { WT901B_parse_gps_frame(		data, &frame_data.data.gps			); break; }
        case WT901B_FRAME_VELOCITY:     { WT901B_parse_velocity_frame(	data, &frame_data.data.velocity		); break; }
        case WT901B_FRAME_QUATERNION:   { WT901B_parse_quaternion_frame(data, &frame_data.data.quaternion	); break; }
        case WT901B_FRAME_GSA:          { WT901B_parse_gsa_frame(		data, &frame_data.data.gsa			); break; }
        default: { wt->last_status = WT901B_INVALID_DATA; return; } // Unknown frame type
    }
	data_topic_publish(&wt->data_topic, &frame_data);
}

static void WT901B_parse_time_frame(const uint8_t data[8], WT901B_TimeFrame_t *frame_obj) {
    frame_obj->year		= data[0];  // YY
    frame_obj->month	= data[1];  // MM
    frame_obj->day		= data[2];  // DD
    frame_obj->hour		= data[3];  // HH
    frame_obj->minute	= data[4];  // MN
    frame_obj->second	= data[5];  // SS
    frame_obj->ms		= (uint16_t)((data[7] << 8) | data[6]); // MSL, MSH
}

static void WT901B_parse_accel_frame(const uint8_t data[8], WT901B_AccelFrame_t *frame_obj) {
    frame_obj->ax_g				= (float)((int16_t)(data[1] << 8 | data[0])) * WT901B_ACCEL_SCALE_G;
    frame_obj->ay_g				= (float)((int16_t)(data[3] << 8 | data[2])) * WT901B_ACCEL_SCALE_G;
    frame_obj->az_g				= (float)((int16_t)(data[5] << 8 | data[4])) * WT901B_ACCEL_SCALE_G;
    frame_obj->temperature_c	= (float)((int16_t)(data[7] << 8 | data[6])) * WT901B_TEMPERATURE_SCALE_C;
}

static void WT901B_parse_gyro_frame(const uint8_t data[8], WT901B_GyroFrame_t *frame_obj) {
    frame_obj->gx_dps = (float)((int16_t)(data[1] << 8 | data[0])) * WT901B_GYRO_SCALE_DPS;
    frame_obj->gy_dps = (float)((int16_t)(data[3] << 8 | data[2])) * WT901B_GYRO_SCALE_DPS;
    frame_obj->gz_dps = (float)((int16_t)(data[5] << 8 | data[4])) * WT901B_GYRO_SCALE_DPS;
    /* data[6..7] : tension d’alim (non utilisée ici) */
}

static void WT901B_parse_angle_frame(const uint8_t data[8], WT901B_AngleFrame_t *frame_obj) {
    frame_obj->roll_deg  = (float)((int16_t)(data[1] << 8 | data[0])) * WT901B_ANGLE_SCALE_DEG;
    frame_obj->pitch_deg = (float)((int16_t)(data[3] << 8 | data[2])) * WT901B_ANGLE_SCALE_DEG;
    frame_obj->yaw_deg   = (float)((int16_t)(data[5] << 8 | data[4])) * WT901B_ANGLE_SCALE_DEG;
    /* data[6..7] : version / statut, ignorés ici */
}

static void WT901B_parse_mag_frame(const uint8_t data[8], WT901B_MagFrame_t *frame_obj) {
    frame_obj->hx_uT			= (float)((int16_t)(data[1] << 8 | data[0])) * WT901B_MAG_SCALE_UT;
    frame_obj->hy_uT			= (float)((int16_t)(data[3] << 8 | data[2])) * WT901B_MAG_SCALE_UT;
    frame_obj->hz_uT			= (float)((int16_t)(data[5] << 8 | data[4])) * WT901B_MAG_SCALE_UT;
    frame_obj->temperature_c	= (float)((int16_t)(data[7] << 8 | data[6])) * WT901B_TEMPERATURE_SCALE_C;
}

static void WT901B_parse_port_frame(const uint8_t data[8], WT901B_PortFrame_t *frame_obj) {
    frame_obj->d0 = (uint16_t)((uint16_t)data[1] << 8 | (uint16_t)data[0]);
    frame_obj->d1 = (uint16_t)((uint16_t)data[3] << 8 | (uint16_t)data[2]);
    frame_obj->d2 = (uint16_t)((uint16_t)data[5] << 8 | (uint16_t)data[4]);
    frame_obj->d3 = (uint16_t)((uint16_t)data[7] << 8 | (uint16_t)data[6]);
}

static void WT901B_parse_pressure_frame(const uint8_t data[8], WT901B_PressureFrame_t *frame_obj) {
    frame_obj->pressure_pa = ((uint32_t)data[3] << 24) |
                			 ((uint32_t)data[2] << 16) |
                			 ((uint32_t)data[1] <<  8) |
                			 ((uint32_t)data[0] <<  0);

    frame_obj->altitude_cm = ((uint32_t)data[7] << 24) |
                    		 ((uint32_t)data[6] << 16) |
                    		 ((uint32_t)data[5] << 8)  |
                    		 ((uint32_t)data[4] << 0);    
}

static void WT901B_parse_gps_frame(const uint8_t data[8], WT901B_GPSFrame_t *frame_obj) {
    frame_obj->longitude = (float)(
        ((uint32_t)data[3] << 24) |
        ((uint32_t)data[2] << 16) |
        ((uint32_t)data[1] << 8)  |
        (uint32_t)data[0]) * 1.0f; // TODO: change converstion

    frame_obj->latitude = (float)(
        ((uint32_t)data[7] << 24) |
        ((uint32_t)data[6] << 16) |
        ((uint32_t)data[5] << 8)  |
        (uint32_t)data[4]) * 1.0f; // TODO: change converstion
}

static void WT901B_parse_velocity_frame(const uint8_t data[8], WT901B_VelocityFrame_t *frame_obj) {
    frame_obj->gps_altitude_m	= (float)((int16_t)(data[1] << 8 | data[0])) * WT901B_GPS_ALTITUDE_SCALE_M;
    frame_obj->gps_heading_deg	= (float)((int16_t)(data[3] << 8 | data[2])) * WT901B_GPS_HEADING_SCALE_DEG;

    frame_obj->gps_speed_kmh = (float)(((uint32_t)data[7] << 24) |
        							   ((uint32_t)data[6] << 16) |
									   ((uint32_t)data[5] <<  8) |
									   ((uint32_t)data[4] <<  0)) * WT901B_GPS_SPEED_SCALE_KMH;
}

static void WT901B_parse_quaternion_frame(const uint8_t data[8], WT901B_QuaternionFrame_t *frame_obj) {
    frame_obj->q0 = (float)((int16_t)(data[1] << 8 | data[0])) * WT901B_QUATERNION_SCALE;
    frame_obj->q1 = (float)((int16_t)(data[3] << 8 | data[2])) * WT901B_QUATERNION_SCALE;
    frame_obj->q2 = (float)((int16_t)(data[5] << 8 | data[4])) * WT901B_QUATERNION_SCALE;
    frame_obj->q3 = (float)((int16_t)(data[7] << 8 | data[6])) * WT901B_QUATERNION_SCALE;
}

static void WT901B_parse_gsa_frame(const uint8_t data[8], WT901B_GSAFrame_t *frame_obj) {
    frame_obj->svnum = (uint16_t)((uint16_t)data[1] << 8 | (uint16_t)data[0]);
    frame_obj->pdop = (float)((int16_t)(data[3] << 8 | data[2])) * 1.0f; // TODO: change converstion
    frame_obj->hdop = (float)((int16_t)(data[5] << 8 | data[4])) * 1.0f; // TODO: change converstion
    frame_obj->vdop = (float)((int16_t)(data[7] << 8 | data[6])) * 1.0f; // TODO: change converstion
}
