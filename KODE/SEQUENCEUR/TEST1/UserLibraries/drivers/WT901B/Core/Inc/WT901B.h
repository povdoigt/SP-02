/**
 * @file    WT901B.h
 * @brief   Driver for WITMotion WT901B IMU sensor (UART interface)
 *
 * This header defines constants, data structures and function prototypes for
 * communicating with a WT901B inertial measurement unit. The WT901B outputs
 * various sensor frames over a serial link; each frame begins with 0x55 and
 * contains a type byte identifying the payload. The formulas used to
 * convert raw values into physical units are taken directly from the
 * manufacturer’s standard communication protocol.
 *
 * To keep the source readable and to avoid magic numbers, all register
 * addresses, frame IDs and bit masks are defined as macros. Enumerations
 * are provided for output rate and baud rate codes, matching the tables in
 * the datasheet.
 */

#ifndef DRIVERS_WT901B_H
#define DRIVERS_WT901B_H

#ifdef __cplusplus
extern "C" {
#endif

#include <stdbool.h>
#include <stdint.h>

#include "usart.h"

#include "data_topic.h"

/* -------------------------------------------------------------------------- */
/*                         Frame Identifiers (TYPE)                           */
/* -------------------------------------------------------------------------- */

/* Trame standard WIT : 0x55, TYPE, 8 data bytes, checksum                    */
#define WT901B_FRAME_HEADER      0x55U
#define WT901B_FRAME_LENGTH      11U
#define WT901B_FRAME_TYPE_NBR    12U    /**< Number of different frame types */

#define WT901B_FRAME_TIME        0x50U  /**< Time output frame */
#define WT901B_FRAME_ACCEL       0x51U  /**< Acceleration and temperature */
#define WT901B_FRAME_GYRO        0x52U  /**< Angular velocity and voltage */
#define WT901B_FRAME_ANGLE       0x53U  /**< Euler angles and firmware version */
#define WT901B_FRAME_MAG         0x54U  /**< Magnetic field and temperature */
#define WT901B_FRAME_PORT        0x55U  /**< D0–D3 port status */
#define WT901B_FRAME_PRESSURE    0x56U  /**< Air pressure and altitude */
#define WT901B_FRAME_GPS         0x57U  /**< Longitude and latitude */
#define WT901B_FRAME_VELOCITY    0x58U  /**< GPS altitude, heading and ground speed */
#define WT901B_FRAME_QUATERNION  0x59U  /**< Quaternion */
#define WT901B_FRAME_GSA         0x5AU  /**< GPS positioning accuracy (PDOP/HDOP/VDOP) */
#define WT901B_FRAME_READREG     0x5FU  /**< Return value for read register commands */

/* -------------------------------------------------------------------------- */
/*                          Register Map (address)                            */
/* -------------------------------------------------------------------------- */

#define WT901B_REG_SAVECONF     0x00U   /**< Save/reboot/reset register */
#define WT901B_REG_CALSW        0x01U   /**< Calibration mode */
#define WT901B_REG_RSW          0x02U   /**< Output content selection */
#define WT901B_REG_RRATE        0x03U   /**< Output rate selection */
#define WT901B_REG_BAUD         0x04U   /**< Serial baud rate */
#define WT901B_REG_AXOFFSET     0x05U   /**< Acceleration X zero offset */
#define WT901B_REG_AYOFFSET     0x06U   /**< Acceleration Y zero offset */
#define WT901B_REG_AZOFFSET     0x07U   /**< Acceleration Z zero offset */
#define WT901B_REG_GXOFFSET     0x08U   /**< Gyro X zero offset */
#define WT901B_REG_GYOFFSET     0x09U   /**< Gyro Y zero offset */
#define WT901B_REG_GZOFFSET     0x0AU   /**< Gyro Z zero offset */
#define WT901B_REG_HXOFFSET     0x0BU   /**< Magnetic field X zero offset */
#define WT901B_REG_HYOFFSET     0x0CU   /**< Magnetic field Y zero offset */
#define WT901B_REG_HZOFFSET     0x0DU   /**< Magnetic field Z zero offset */
#define WT901B_REG_D0MODE       0x0EU   /**< D0 pin mode */
#define WT901B_REG_D1MODE       0x0FU   /**< D1 pin mode */
#define WT901B_REG_D2MODE       0x10U   /**< D2 pin mode */
#define WT901B_REG_D3MODE       0x11U   /**< D3 pin mode */
// REG 0x12 to 0x19 are reserved
#define WT901B_REG_IICADDR      0x1AU   /**< I2C device address */
#define WT901B_REG_LEDOFF       0x1BU   /**< LED off control */
#define WT901B_REG_MAGRANGEX    0x1CU   /**< Magnetic field measurement range X */
#define WT901B_REG_MAGRANGEY    0x1DU   /**< Magnetic field measurement range Y */
#define WT901B_REG_MAGRANGEZ    0x1EU   /**< Magnetic field measurement range Z */
#define WT901B_REG_BANDWIDTH    0x1FU   /**< Gyro bandwidth setting */
#define WT901B_REG_GYRORANGE    0x20U   /**< Gyro measurement range */
#define WT901B_REG_ACCRANGE     0x21U   /**< Acceleration measurement range */
#define WT901B_REG_SLEEP        0x22U   /**< Sleep mode control */
#define WT901B_REG_ORIENT       0x23U   /**< Orientation setting */
#define WT901B_REG_AXIS         0x24U   /**< Axis direction setting */
#define WT901B_REG_FILTK        0x25U   /**< Acceleration filter setting */
#define WT901B_REG_GPSBAND      0x26U   /**< GPS working band */
#define WT901B_REG_READADDR     0x27U   /**< READADDR register for reading other registers */
// REG 0x28 and 0x29 are reserved
#define WT901B_REG_ACCFILT      0x2AU   /**< Acceleration filter coefficient */
// REG 0x2B and 0x2C
#define WT901B_REG_ACCFILT      0x2AU   /**< Acceleration filter coefficient */
// REG 0x2B and 0x2C are reserved
#define WT901B_REG_POWONSEND    0x2DU   /**< Power-on send / command start control */
#define WT901B_REG_VERSION      0x2EU   /**< Firmware version */
// REG 0x2F is reserved
#define WT901B_REG_YYMM         0x30U   /**< Date: year/month (YYMM) */
#define WT901B_REG_DDHH         0x31U   /**< Date: day/hour (DDHH) */
#define WT901B_REG_MMSS         0x32U   /**< Time: minute/second (MMSS) */
#define WT901B_REG_MS           0x33U   /**< Time: millisecond */
#define WT901B_REG_AX           0x34U   /**< Raw acceleration X */
#define WT901B_REG_AY           0x35U   /**< Raw acceleration Y */
#define WT901B_REG_AZ           0x36U   /**< Raw acceleration Z */
#define WT901B_REG_GX           0x37U   /**< Raw angular velocity X */
#define WT901B_REG_GY           0x38U   /**< Raw angular velocity Y */
#define WT901B_REG_GZ           0x39U   /**< Raw angular velocity Z */
#define WT901B_REG_HX           0x3AU   /**< Raw magnetic field X */
#define WT901B_REG_HY           0x3BU   /**< Raw magnetic field Y */
#define WT901B_REG_HZ           0x3CU   /**< Raw magnetic field Z */
#define WT901B_REG_ROLL         0x3DU   /**< Roll angle */
#define WT901B_REG_PITCH        0x3EU   /**< Pitch angle */
#define WT901B_REG_YAW          0x3FU   /**< Yaw / heading angle */
#define WT901B_REG_TEMP         0x40U   /**< Temperature */
#define WT901B_REG_D0STATUS     0x41U   /**< D0 pin status */
#define WT901B_REG_D1STATUS     0x42U   /**< D1 pin status */
#define WT901B_REG_D2STATUS     0x43U   /**< D2 pin status */
#define WT901B_REG_D3STATUS     0x44U   /**< D3 pin status */
#define WT901B_REG_PRESSUREL    0x45U   /**< Air pressure low 16 bits */
#define WT901B_REG_PRESSUREH    0x46U   /**< Air pressure high 16 bits */
#define WT901B_REG_HEIGHTL      0x47U   /**< Height low 16 bits */
#define WT901B_REG_HEIGHTH      0x48U   /**< Height high 16 bits */
#define WT901B_REG_LONL         0x49U   /**< Longitude low 16 bits */
#define WT901B_REG_LONH         0x4AU   /**< Longitude high 16 bits */
#define WT901B_REG_LATL         0x4BU   /**< Latitude low 16 bits */
#define WT901B_REG_LATH         0x4CU   /**< Latitude high 16 bits */
#define WT901B_REG_GPSHEIGHT    0x4DU   /**< GPS altitude */
#define WT901B_REG_GPSYAW       0x4EU   /**< GPS heading */
#define WT901B_REG_GPSVL        0x4FU   /**< GPS ground speed low 16 bits */
#define WT901B_REG_GPSVH        0x50U   /**< GPS ground speed high 16 bits */
#define WT901B_REG_Q0           0x51U   /**< Quaternion component q0 */
#define WT901B_REG_Q1           0x52U   /**< Quaternion component q1 */
#define WT901B_REG_Q2           0x53U   /**< Quaternion component q2 */
#define WT901B_REG_Q3           0x54U   /**< Quaternion component q3 */
#define WT901B_REG_SVNUM        0x55U   /**< Number of satellites */
#define WT901B_REG_PDOP         0x56U   /**< Position DOP */
#define WT901B_REG_HDOP         0x57U   /**< Horizontal DOP */
#define WT901B_REG_VDOP         0x58U   /**< Vertical DOP */
#define WT901B_REG_DELAYT       0x59U   /**< Alarm delay time */
#define WT901B_REG_XMIN         0x5AU   /**< X-axis angle alarm minimum */
#define WT901B_REG_XMAX         0x5BU   /**< X-axis angle alarm maximum */
#define WT901B_REG_BATVAL       0x5CU   /**< Supply voltage measurement */
#define WT901B_REG_ALARMPIN     0x5DU   /**< Alarm pin mapping */
#define WT901B_REG_YMIN         0x5EU   /**< Y-axis angle alarm minimum */
#define WT901B_REG_YMAX         0x5FU   /**< Y-axis angle alarm maximum */
// REG 0x60 is reserved
#define WT901B_REG_GYROCALITHR  0x61U   /**< Gyro still threshold */
#define WT901B_REG_ALARMLEVEL   0x62U   /**< Angle alarm level */
#define WT901B_REG_GYROCALTIME  0x63U   /**< Gyro auto-calibration time */
// REG 0x64 to 0x67 are reserved
#define WT901B_REG_TRIGTIME     0x68U   /**< Alarm continuous trigger time */
#define WT901B_REG_KEY          0x69U   /**< Unlock key */
#define WT901B_REG_WERROR       0x6AU   /**< Gyro change indicator */
#define WT901B_REG_TIMEZONE     0x6BU   /**< GPS time zone */
// REG 0x6C and 0x6D are reserved
#define WT901B_REG_WZTIME       0x6EU   /**< Angular velocity continuous rest time */
#define WT901B_REG_WZSTATIC     0x6FU   /**< Angular velocity integral threshold */
// REG 0x70 to 0x73 are reserved
#define WT901B_REG_MODDELAY     0x74U   /**< RS-485 data response delay */
// REG 0x75 to 0x78 are reserved
#define WT901B_REG_XREFROLL     0x79U   /**< Roll angle zero reference value */
#define WT901B_REG_YREFPITCH    0x7AU   /**< Pitch angle zero reference value */
// REG 0x7B to 0x7E are reserved
#define WT901B_REG_NUMBERID1    0x7FU   /**< Device ID bytes 1–2 */
#define WT901B_REG_NUMBERID2    0x80U   /**< Device ID bytes 3–4 */
#define WT901B_REG_NUMBERID3    0x81U   /**< Device ID bytes 5–6 */
#define WT901B_REG_NUMBERID4    0x82U   /**< Device ID bytes 7–8 */
#define WT901B_REG_NUMBERID5    0x83U   /**< Device ID bytes 9–10 */
#define WT901B_REG_NUMBERID6    0x84U   /**< Device ID bytes 11–12 */

/* -------------------------------------------------------------------------- */
/*                        RSW Output Content Bit Masks                        */
/* -------------------------------------------------------------------------- */
/**
 * Bits in the RSW register control which frames are output by the sensor. A
 * value of 1 enables the corresponding frame, while 0 disables it. Bit 0
 * corresponds to the time frame and bit 10 corresponds to the GPS
 * positioning accuracy frame.
 */
#define WT901B_RSW_TIME_BIT        (1U << 0)    /**< Enable time frame (TYPE=0x50) */
#define WT901B_RSW_ACC_BIT         (1U << 1)    /**< Enable acceleration frame (TYPE=0x51) */
#define WT901B_RSW_GYRO_BIT        (1U << 2)    /**< Enable gyro frame (TYPE=0x52) */
#define WT901B_RSW_ANGLE_BIT       (1U << 3)    /**< Enable Euler angle frame (TYPE=0x53) */
#define WT901B_RSW_MAG_BIT         (1U << 4)    /**< Enable magnetic field frame (TYPE=0x54) */
#define WT901B_RSW_PORT_BIT        (1U << 5)    /**< Enable port status frame (TYPE=0x55) */
#define WT901B_RSW_PRESSURE_BIT    (1U << 6)    /**< Enable air pressure/altitude frame (TYPE=0x56) */
#define WT901B_RSW_GPS_BIT         (1U << 7)    /**< Enable longitude/latitude frame (TYPE=0x57) */
#define WT901B_RSW_VELOCITY_BIT    (1U << 8)    /**< Enable GPS velocity frame (TYPE=0x58) */
#define WT901B_RSW_QUATERNION_BIT  (1U << 9)    /**< Enable quaternion frame (TYPE=0x59) */
#define WT901B_RSW_GSA_BIT         (1U << 10)   /**< Enable GPS accuracy frame (TYPE=0x5A) */

/* -------------------------------------------------------------------------- */
/*                        RRATE Output Rate Enumerations                      */
/* -------------------------------------------------------------------------- */
/**
 * Output rates supported by the WT901B. The codes correspond to the
 * lower 4 bits of the RRATE register. For example, a value of 0x06 sets a
 * 10 Hz update rate. Note that some product variants support extended rates such
 * as 500 Hz/1000 Hz (codes 0x0C and
 * 0x0D). These are included here for completeness but may not be supported
 * by the WT901B itself.
 */
typedef enum WT901B_rrate_t {
    WT901B_RRATE_0P2HZ    = 0x01U,  /**< 0.2 Hz output rate */
    WT901B_RRATE_0P5HZ    = 0x02U,  /**< 0.5 Hz output rate */
    WT901B_RRATE_1HZ      = 0x03U,  /**< 1 Hz output rate */
    WT901B_RRATE_2HZ      = 0x04U,  /**< 2 Hz output rate */
    WT901B_RRATE_5HZ      = 0x05U,  /**< 5 Hz output rate */
    WT901B_RRATE_10HZ     = 0x06U,  /**< 10 Hz output rate */
    WT901B_RRATE_20HZ     = 0x07U,  /**< 20 Hz output rate */
    WT901B_RRATE_50HZ     = 0x08U,  /**< 50 Hz output rate */
    WT901B_RRATE_100HZ    = 0x09U,  /**< 100 Hz output rate */
    WT901B_RRATE_200HZ    = 0x0BU,  /**< 200 Hz output rate */
    WT901B_RRATE_SINGLE   = 0x0CU,  /**< Single measurement / 500 Hz on some models */
    WT901B_RRATE_NORETURN = 0x0DU   /**< Disable continuous output (no return) */
} WT901B_rrate_t;

/* -------------------------------------------------------------------------- */
/*                          BAUD Serial Rate Enumerations                    */
/* -------------------------------------------------------------------------- */
/**
 * Serial baud rates supported by the WT901B. Codes are taken from the
 * BAUD register definition【522002904479991†L1689-L1702】. Selecting a baud
 * rate outside the supported range will cause the device to stop responding.
 */
typedef enum WT901B_baud_t {
    WT901B_BAUD_4800    = 0x01U,    /**< 4800 bps */
    WT901B_BAUD_9600    = 0x02U,    /**< 9600 bps */
    WT901B_BAUD_19200   = 0x03U,    /**< 19 200 bps */
    WT901B_BAUD_38400   = 0x04U,    /**< 38 400 bps */
    WT901B_BAUD_57600   = 0x05U,    /**< 57 600 bps */
    WT901B_BAUD_115200  = 0x06U,    /**< 115 200 bps */
    WT901B_BAUD_230400  = 0x07U,    /**< 230 400 bps */
    WT901B_BAUD_460800  = 0x08U,    /**< 460 800 bps (not supported on WT901B) */
    WT901B_BAUD_921600  = 0x09U     /**< 921 600 bps (not supported on WT901B) */
} WT901B_baud_t;

/* -------------------------------------------------------------------------- */
/*                         Conversion Factors (float)                        */
/* -------------------------------------------------------------------------- */
/**
 * The following constants implement the scaling formulas provided in the
 * protocol document for converting raw sensor readings into physical units.
 * These values are defined as macros to avoid magic numbers in the source
 * code. Refer to the comments for the origin of each constant.
 */
/** Acceleration scale (g per LSB). Formula: raw/32768*16 g */
#define WT901B_ACCEL_SCALE_G        (16.0f / 32768.0f)
/** Gyroscope scale (°/s per LSB). Formula: raw/32768*2000 °/s */
#define WT901B_GYRO_SCALE_DPS       (2000.0f / 32768.0f)
/** Angle scale (° per LSB). Formula: raw/32768*180 ° */
#define WT901B_ANGLE_SCALE_DEG      (180.0f / 32768.0f)
/** Magnetic field scale (uT per LSB). Formula: raw/0.15 uT */
#define WT901B_MAG_SCALE_UT         (1.0f / 0.15f)
/** Temperature scale (°C per LSB). Formula: raw/100 °C */
#define WT901B_TEMPERATURE_SCALE_C  (1.0f / 100.0f)
/** Quaternion scale (dimensionless per LSB). Formula: raw/32768 */
#define WT901B_QUATERNION_SCALE     (1.0f / 32768.0f)
/** GPS altitude scale (m per LSB). Formula: raw/10 m */
#define WT901B_GPS_ALTITUDE_SCALE_M (1.0f / 10.0f)
/** GPS heading scale (° per LSB). Formula: raw/10 ° */
#define WT901B_GPS_HEADING_SCALE_DEG (1.0f / 10.0f)
/** GPS speed scale (km/h per LSB). Formula: raw/1000 km/h */
#define WT901B_GPS_SPEED_SCALE_KMH  (1.0f / 1000.0f)
/** PDOP/HDOP/VDOP scale (dimensionless per LSB). Formula: raw/100 */
#define WT901B_DOP_SCALE            (1.0f / 100.0f)

/* -------------------------------------------------------------------------- */
/*                           Status Code Definition                          */
/* -------------------------------------------------------------------------- */
/**
 * Return codes for driver functions. Negative values indicate errors.
 */
typedef enum WT901B_status_t{
    WT901B_OK               =  0,	/**< Operation succeeded */
    WT901B_ERROR            = -1,	/**< Generic error */
    WT901B_UART_ERROR       = -2,	/**< HAL UART transmit/receive error */
    WT901B_INVALID_ARG      = -3,	/**< Null pointer or invalid argument */
    WT901B_CHECKSUM_ERROR   = -4,	/**< Frame checksum mismatch */
    WT901B_NO_DATA          = -5,	/**< No new data available */
    WT901B_OVERRUN_ERROR    = -6,	/**< Data overrun: incoming data lost */
    WT901B_TIMEOUT          = -7,	/**< Operation timed out */
    WT901B_INVALID_DATA     = -8,	/**< Received invalid data */
} WT901B_status_t;



typedef struct {
    uint8_t year;
    uint8_t month;
    uint8_t day;
    uint8_t hour;
    uint8_t minute;
    uint8_t second;
    uint16_t ms;
} WT901B_TimeFrame_t;

typedef struct WT901B_AccelFrame_t {
    float ax_g;
    float ay_g;
    float az_g;
    float temperature_c;
} WT901B_AccelFrame_t;

typedef struct WT901B_GyroFrame_t {
    float gx_dps;
    float gy_dps;
    float gz_dps;
    // float voltage_v;       /* optionnel si tu l’exploites */
} WT901B_GyroFrame_t;

typedef struct WT901B_AngleFrame_t {
    float roll_deg;
    float pitch_deg;
    float yaw_deg;
    // uint16_t version;      /* ou status/version */
} WT901B_AngleFrame_t;

typedef struct WT901B_MagFrame_t {
    int16_t hx_uT;
    int16_t hy_uT;
    int16_t hz_uT;
    float temperature_c;
} WT901B_MagFrame_t;

typedef struct WT901B_PortFrame_t {
    uint16_t d0;
    uint16_t d1;
    uint16_t d2;
    uint16_t d3;
} WT901B_PortFrame_t;

typedef struct WT901B_PressureFrame_t {
    uint32_t pressure_pa;
    uint32_t altitude_cm;
} WT901B_PressureFrame_t;

typedef struct WT901B_GPSFrame_t {
    float longitude;
    float latitude;
} WT901B_GPSFrame_t;

typedef struct WT901B_VelocityFrame_t {
    float gps_altitude_m;
    float gps_heading_deg;
    float gps_speed_kmh;
} WT901B_VelocityFrame_t;

typedef struct WT901B_QuaternionFrame_t {
    float q0;
    float q1;
    float q2;
    float q3;
} WT901B_QuaternionFrame_t;

typedef struct WT901B_GSAFrame_t {
    uint16_t svnum;
    float pdop;
    float hdop;
    float vdop;
} WT901B_GSAFrame_t;

typedef union WT901B_FrameData_u {
    WT901B_TimeFrame_t time;
    WT901B_AccelFrame_t accel;
    WT901B_GyroFrame_t gyro;
    WT901B_AngleFrame_t angle;
    WT901B_MagFrame_t mag;
    WT901B_PortFrame_t port;
    WT901B_PressureFrame_t pressure;
    WT901B_GPSFrame_t gps;
    WT901B_VelocityFrame_t velocity;
    WT901B_QuaternionFrame_t quaternion;
    WT901B_GSAFrame_t gsa;
} WT901B_FrameData_u;

typedef struct WT901B_Frame_t {
    uint8_t type;
	uint32_t timestamp_ms;
    WT901B_FrameData_u data;
} WT901B_Frame_t;


#define WT901B_DATA_TOPIC_LENGTH 16

/* -------------------------------------------------------------------------- */
/*                          Driver Data Structures                           */
/* -------------------------------------------------------------------------- */
/**
 * @brief Device state structure for the WT901B driver.
 *
 * This structure holds both the hardware handle (UART) and the most recent
 * parsed sensor values. It also contains a small internal buffer used to
 * assemble frames from the serial stream. Flags indicate when new data has
 * been parsed since the last read.
 */
typedef struct WT901B_t {
    UART_HandleTypeDef *huart;      /**< UART handle used for communication */

    /* Internal parser state */
    uint8_t uart_buffer[WT901B_FRAME_LENGTH * WT901B_FRAME_TYPE_NBR];  /**< Buffer for incoming UART data */
    uint8_t parse_buffer[WT901B_FRAME_LENGTH * WT901B_FRAME_TYPE_NBR];  /**< Buffer for assembling incoming frames */
    bool new_data_available;      /**< Flag indicating new data is available to parse */
    bool data_in_progress;      /**< Flag indicating a frame is currently being assembled */
    size_t last_received_size;   /**< Size of the last received data chunk */
	uint32_t last_timestamp_ms;  /**< Timestamp of the last received data chunk */

    WT901B_status_t last_status;   /**< Last operation status */

    /* Data topic */
    WT901B_Frame_t dt_storage[WT901B_DATA_TOPIC_LENGTH];
	data_topic_t data_topic;

} WT901B_t;

extern WT901B_t wt901b;

/* -------------------------------------------------------------------------- */
/*                         API Function Declarations                         */
/* -------------------------------------------------------------------------- */
/**
 * @brief Initialise a WT901B device structure.
 *
 * This function resets the parser state and stores the UART handle passed
 * by the user. It does not perform any hardware initialisation; the UART
 * should already be configured at the correct baud rate before calling this
 * function.
 *
 * @param dev  Pointer to a wt901b_t structure.
 * @param huart Pointer to the HAL UART handle used for communication.
 * @return WT901B_OK if the parameters are valid, WT901B_INVALID_ARG otherwise.
 */
WT901B_status_t WT901B_Init(WT901B_t *wt, UART_HandleTypeDef *huart);

void WT901B_UART_Callback_RX_IRQHandler(WT901B_t *wt, uint16_t Size);

void WT901B_Parse_Frames(WT901B_t *wt);












#ifdef __cplusplus
}
#endif

#endif /* DRIVERS_WT901B_H */