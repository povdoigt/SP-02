/*
 * WT901B.c — Driver implementation for the WITMotion WT901B IMU.
 *
 * This source file provides a lightweight parser for the serial data frames
 * produced by a WT901B inertial sensor and helper functions to configure
 * the device via its UART interface. Refer to WT901B.h for a detailed
 * description of the available registers, frame formats and conversion
 * constants. The formulas used here are derived directly from the
 * manufacturer’s standard communication protocol【522002904479991†L866-L899】【522002904479991†L987-L1014】.
 */

#include "WT901B.h"
#include "usart.h"

#include <string.h>


/* -------------------------------------------------------------------------- */
/*                        Internal Frame Sub-Parsers                          */
/* -------------------------------------------------------------------------- */

static void WT901B_parse_time_frame(const uint8_t *data);
static void WT901B_parse_accel_frame(const uint8_t *data);
static void WT901B_parse_gyro_frame(const uint8_t *data);
static void WT901B_parse_angle_frame(const uint8_t *data);
static void WT901B_parse_mag_frame(const uint8_t *data);
static void WT901B_parse_port_frame(const uint8_t *data);
static void WT901B_parse_pressure_frame(const uint8_t *data);
static void WT901B_parse_gps_frame(const uint8_t *data);
static void WT901B_parse_velocity_frame(const uint8_t *data);
static void WT901B_parse_quaternion_frame(const uint8_t *data);
static void WT901B_parse_gsa_frame(const uint8_t *data);
static void WT901B_parse_readreg_frame(const uint8_t *data);



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
    memset(wt->frame_buf, 0, WT901B_FRAME_LENGTH);
    wt->new_data_nbr = 0;

    UART_buffer_t *buffer_obj;
    HAL_StatusTypeDef res;
    UART_get_buffer(huart, &buffer_obj);
    if (buffer_obj != NULL) {
        buffer_obj->rx_buffer = wt->frame_buf;
        buffer_obj->rx_length = WT901B_FRAME_LENGTH * WT901B_FRAME_TYPE_NBR;
        res = HAL_UARTEx_ReceiveToIdle_IT(huart, buffer_obj->rx_buffer, buffer_obj->rx_length);
        if (res != HAL_OK) {
            return WT901B_UART_ERROR; // UART receive error
        }
    } else {
        return WT901B_UART_ERROR; // Unknown UART instance
    }

    return WT901B_OK;
}

void WT901B_UART_Callback_RX_IRQHandler(WT901B_t *wt, uint16_t Size) {
    // Frame received
    wt->new_data_nbr = Size / WT901B_FRAME_LENGTH; // Number of complete frames received
    // Further processing should be done in the main application loop
    // to avoid lengthy operations in the interrupt context
}


void WT901B_Parse_Frames(WT901B_t *wt) {
    for (uint8_t frame_idx = 0; frame_idx < wt->new_data_nbr; frame_idx++) {
        uint8_t *frame_ptr = &wt->frame_buf[frame_idx * WT901B_FRAME_LENGTH];

        // Verify header
        if (frame_ptr[0] != WT901B_FRAME_HEADER) {
            continue; // Invalid header, skip frame
        }

        // Verify checksum
        uint8_t computed_checksum = wt901b_compute_checksum(frame_ptr);
        if (computed_checksum != frame_ptr[10]) {
            continue; // Checksum mismatch, skip frame
        }

        // Valid frame received, process based on type
        uint8_t frame_type = frame_ptr[1];
        uint8_t *data = &frame_ptr[2];
        switch (frame_type) {
            case WT901B_FRAME_TIME:         { WT901B_parse_time_frame(data);       break; }
            case WT901B_FRAME_ACCEL:        { WT901B_parse_accel_frame(data);      break; }
            case WT901B_FRAME_GYRO:         { WT901B_parse_gyro_frame(data);       break; }
            case WT901B_FRAME_ANGLE:        { WT901B_parse_angle_frame(data);      break; }
            case WT901B_FRAME_MAG:          { WT901B_parse_mag_frame(data);        break; }
            case WT901B_FRAME_PORT:         { WT901B_parse_port_frame(data);       break; }
            case WT901B_FRAME_PRESSURE:     { WT901B_parse_pressure_frame(data);   break; }
            case WT901B_FRAME_GPS:          { WT901B_parse_gps_frame(data);        break; }
            case WT901B_FRAME_VELOCITY:     { WT901B_parse_velocity_frame(data);   break; }
            case WT901B_FRAME_QUATERNION:   { WT901B_parse_quaternion_frame(data); break; }
            case WT901B_FRAME_GSA:          { WT901B_parse_gsa_frame(data);        break; }
            case WT901B_FRAME_READREG:      { WT901B_parse_readreg_frame(data);    break; }
            default: { /* Unknown frame type, skip */ break; }
        }
    }
    wt->new_data_nbr = 0; // Reset new data counter after processing
}

static void WT901B_parse_time_frame(const uint8_t *data) {
    uint8_t year   = data[0];  // YY
    uint8_t month  = data[1];  // MM
    uint8_t day    = data[2];  // DD
    uint8_t hour   = data[3];  // HH
    uint8_t minute = data[4];  // MN
    uint8_t second = data[5];  // SS
    uint16_t ms    = (uint16_t)((data[7] << 8) | data[6]); // MSL, MSH

    (void)year; (void)month; (void)day;
    (void)hour; (void)minute; (void)second; (void)ms;

    // TODO: publier / stocker la date/heure si nécessaire
}

static void WT901B_parse_accel_frame(const uint8_t *data) {
    float ax = (float)((int16_t)data[1] << 8 | (int16_t)data[0]) * WT901B_ACCEL_SCALE_G;
    float ay = (float)((int16_t)data[3] << 8 | (int16_t)data[2]) * WT901B_ACCEL_SCALE_G;
    float az = (float)((int16_t)data[5] << 8 | (int16_t)data[4]) * WT901B_ACCEL_SCALE_G;
    float t  = (float)((int16_t)data[7] << 8 | (int16_t)data[6]) * WT901B_TEMPERATURE_SCALE_C;

    (void)ax; (void)ay; (void)az; (void)t;

    // TODO: publier / stocker l'accélération si nécessaire
}

static void WT901B_parse_gyro_frame(const uint8_t *data) {
    float wx = (float)((int16_t)data[1] << 8 | (int16_t)data[0]) * WT901B_GYRO_SCALE_DPS;
    float wy = (float)((int16_t)data[3] << 8 | (int16_t)data[2]) * WT901B_GYRO_SCALE_DPS;
    float wz = (float)((int16_t)data[5] << 8 | (int16_t)data[4]) * WT901B_GYRO_SCALE_DPS;
    /* data[6..7] : tension d’alim (non utilisée ici) */

    (void)wx; (void)wy; (void)wz;

    // TODO: publier / stocker la vitesse angulaire si nécessaire
}

static void WT901B_parse_angle_frame(const uint8_t *data) {
    float roll  = (float)((int16_t)data[1] << 8 | (int16_t)data[0]) * WT901B_ANGLE_SCALE_DEG;
    float pitch = (float)((int16_t)data[3] << 8 | (int16_t)data[2]) * WT901B_ANGLE_SCALE_DEG;
    float yaw   = (float)((int16_t)data[5] << 8 | (int16_t)data[4]) * WT901B_ANGLE_SCALE_DEG;
    /* data[6..7] : version / statut, ignorés ici */

    (void)roll; (void)pitch; (void)yaw;

    // TODO: publier / stocker les angles d’Euler si nécessaire
}

static void WT901B_parse_mag_frame(const uint8_t *data) {
    float hx = (float)((int16_t)data[1] << 8 | (int16_t)data[0]) * 1.0f; // TODO: change converstion
    float hy = (float)((int16_t)data[3] << 8 | (int16_t)data[2]) * 1.0f; // TODO: change converstion
    float hz = (float)((int16_t)data[5] << 8 | (int16_t)data[4]) * 1.0f; // TODO: change converstion
    float t  = (float)((int16_t)data[7] << 8 | (int16_t)data[6]) * WT901B_TEMPERATURE_SCALE_C;

    (void)hx; (void)hy; (void)hz; (void)t;

    // TODO: publier / stocker le champ magnétique si nécessaire 
}

static void WT901B_parse_port_frame(const uint8_t *data) {
    uint16_t d0 = (uint16_t)((uint16_t)data[1] << 8 | (uint16_t)data[0]);
    uint16_t d1 = (uint16_t)((uint16_t)data[3] << 8 | (uint16_t)data[2]);
    uint16_t d2 = (uint16_t)((uint16_t)data[5] << 8 | (uint16_t)data[4]);
    uint16_t d3 = (uint16_t)((uint16_t)data[7] << 8 | (uint16_t)data[6]);

    (void)d0; (void)d1; (void)d2; (void)d3;

    // TODO: publier / stocker les données du port si nécessaire
}

static void WT901B_parse_pressure_frame(const uint8_t *data) {
    float p = (float)(
            ((uint32_t)data[3] << 24) |
            ((uint32_t)data[2] << 16) |
            ((uint32_t)data[1] << 8)  |
            (uint32_t)data[0]) * 1.0f; // TODO: change converstion

    uint32_t h_cm =
        ((uint32_t)data[7] << 24) |
        ((uint32_t)data[6] << 16) |
        ((uint32_t)data[5] << 8)  |
        (uint32_t)data[4];

    (void)p; (void)h_cm;

    // TODO: publier / stocker la pression et l’altitude si nécessaire    
}

static void WT901B_parse_gps_frame(const uint8_t *data) {
    float lon = (float)(
        ((uint32_t)data[3] << 24) |
        ((uint32_t)data[2] << 16) |
        ((uint32_t)data[1] << 8)  |
        (uint32_t)data[0]) * 1.0f; // TODO: change converstion

    float lat = (float)(
        ((uint32_t)data[7] << 24) |
        ((uint32_t)data[6] << 16) |
        ((uint32_t)data[5] << 8)  |
        (uint32_t)data[4]) * 1.0f; // TODO: change converstion

    (void)lon; (void)lat;

    // TODO: publier / stocker la position GPS si nécessaire
}

static void WT901B_parse_velocity_frame(const uint8_t *data) {
    float alt = (float)((int16_t)data[1] << 8 | (int16_t)data[0]) * WT901B_GPS_ALTITUDE_SCALE_M;
    float yaw = (float)((int16_t)data[3] << 8 | (int16_t)data[2]) * 1.0f; // TODO: change converstion

    float speed = (float)(
        ((uint32_t)data[7] << 24) |
        ((uint32_t)data[6] << 16) |
        ((uint32_t)data[5] << 8)  |
        (uint32_t)data[4]) * WT901B_GPS_SPEED_SCALE_KMH;

    (void)alt; (void)yaw; (void)speed;

    // TODO: publier / stocker la vitesse et l’altitude si nécessaire
}

static void WT901B_parse_quaternion_frame(const uint8_t *data) {
    float q0 = (float)((int16_t)data[1] << 8 | (int16_t)data[0]) * WT901B_QUATERNION_SCALE;
    float q1 = (float)((int16_t)data[3] << 8 | (int16_t)data[2]) * WT901B_QUATERNION_SCALE;
    float q2 = (float)((int16_t)data[5] << 8 | (int16_t)data[4]) * WT901B_QUATERNION_SCALE;
    float q3 = (float)((int16_t)data[7] << 8 | (int16_t)data[6]) * WT901B_QUATERNION_SCALE;

    (void)q0; (void)q1; (void)q2; (void)q3;

    // TODO: publier / stocker le quaternion si nécessaire
}

static void WT901B_parse_gsa_frame(const uint8_t *data) {
    uint16_t svnum = (uint16_t)((uint16_t)data[1] << 8 | (uint16_t)data[0]);
    float pdop = (float)((int16_t)data[3] << 8 | (int16_t)data[2]) * 1.0f; // TODO: change converstion
    float hdop = (float)((int16_t)data[5] << 8 | (int16_t)data[4]) * 1.0f; // TODO: change converstion
    float vdop = (float)((int16_t)data[7] << 8 | (int16_t)data[6]) * 1.0f; // TODO: change converstion

    (void)svnum; (void)pdop; (void)hdop; (void)vdop;

    // TODO: publier / stocker les données GSA si nécessaire
}

static void WT901B_parse_readreg_frame(const uint8_t *data) {
    uint16_t r0 = (uint16_t)((uint16_t)data[1] << 8 | (uint16_t)data[0]);
    uint16_t r1 = (uint16_t)((uint16_t)data[3] << 8 | (uint16_t)data[2]);
    uint16_t r2 = (uint16_t)((uint16_t)data[5] << 8 | (uint16_t)data[4]);
    uint16_t r3 = (uint16_t)((uint16_t)data[7] << 8 | (uint16_t)data[6]);

    (void)r0; (void)r1; (void)r2; (void)r3;

    // TODO: publier / stocker les valeurs des registres si nécessaire
}
