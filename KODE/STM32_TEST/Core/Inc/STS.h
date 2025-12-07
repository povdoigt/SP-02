/**
    ******************************************************************************
    * @file    STS.h
    * @author  Alexis Paillard
    * @brief   This file contains all the functions prototypes for the STS3215
    *          servomotore from Feetech.
    ******************************************************************************
*/


#ifndef STS_H
#define STS_H

#include "stm32l4xx_hal.h"
#include <stdbool.h>
#include <stdint.h>

#ifdef __cplusplus
extern "C" {
#endif

// Register definitions for STS3215 sensor

// Version information registers (read-only)
#define STS_REG_FIRMWARE_MAJ_VERSION    0x00    // Major version register
#define STS_REG_FIRMWARE_MIN_VERSION    0x01    // Minor version register
#define STS_REG_END                     0x02    // 0: little endian, 1: big endian
#define STS_REG_MAJ_VERSION             0x03    // Major version of the servomotor
#define STS_REG_MIN_VERSION             0x04    // Minor version of the servomotor

// EPROM registers
#define STS_REG_ID                      0x05    // ID of the servomotor
#define STS_REG_BAUDRATE                0x06    // Baudrate setting
#define STS_REG_RETURN_DELAY_TIME       0x07    // Return delay time setting
#define STS_REG_RETURN_STATUS           0x08    // Return status setting
#define STS_REG_CW_ANGLE_LIMIT_L        0x09    // Clockwise angle limit low byte
#define STS_REG_CW_ANGLE_LIMIT_H        0x0A    // Clockwise angle limit high byte
#define STS_REG_CCW_ANGLE_LIMIT_L       0x0B    // Counter-clockwise angle limit low byte
#define STS_REG_CCW_ANGLE_LIMIT_H       0x0C    // Counter-clockwise angle limit high byte
#define STS_REG_MAX_TEMP                0x0D    // Maximum temperature limit
#define STS_REG_MAX_VOLTAGE             0x0E    // Maximum voltage limit
#define STS_REG_MIN_VOLTAGE             0x0F    // Minimum voltage limit
#define STS_REG_MAX_TORQUE_L            0x10    // Maximum torque low byte
#define STS_REG_MAX_TORQUE_H            0x11    // Maximum torque high byte
#define STS_REG_PHASE                   0x12    // Phase setting
#define STS_REG_UNINSTALL_CONDITION     0x13    // Uninstall condition setting
#define STS_REG_ALARM_LED               0x14    // Alarm LED setting
#define STS_REG_P_COEF                  0x15    // Proportional coefficient setting of PID
#define STS_REG_D_COEF                  0x16    // Differential coefficient setting of PID
#define STS_REG_I_COEF                  0x17    // Integral coefficient setting of PID
#define STS_REG_MIN_START_TORQUE        0x18    // Minimum start torque setting
#define STS_REG_POINTS_LIMIT            0x19    // ???
#define STS_REG_POS_INSENSITIVITY       0x1A    // ???
#define STS_REG_NEG_INSENSITIVITY       0x1B    // ???
#define STS_REG_PROTECTIVE_CURRENT_L    0x1C    // Protective current low byte
#define STS_REG_PROTECTIVE_CURRENT_H    0x1D    // Protective current high byte
#define STS_REG_ANGULAR_RESOLUTION      0x1E    // Angular resolution setting
#define STS_REG_POSITION_OFFSET_L       0x1F    // Position offset low byte
#define STS_REG_POSITION_OFFSET_H       0x20    // Position offset high byte
#define STS_REG_OPERATING_MODE          0x21    // Operating mode setting
#define STS_REG_MAINTAIN_TORQUE         0x22    // Maintain torque setting
#define STS_REG_PROTECTION_TIME         0x23    // Protection time setting
#define STS_REG_OVERLOAD_TORQUE         0x24    // Overload torque
#define STS_REG_SPEED_CL_P_COEF         0x25    // Speed closed-loop proportional coefficient
#define STS_REG_OVERCURRENT_PROTECT     0x26    // Overcurrent protection setting
#define STS_REG_SPEED_CL_I_COEF         0x27    // Speed closed-loop integral coefficient

// SRAM Control registers
#define STS_REG_TORQUE_SWITCH           0x28    // Torque switch
#define STS_REG_GOAL_ACCELERATION       0x29    // Acceleration setting
#define STS_REG_GOAL_POSITION_L         0x2A    // Goal position low byte
#define STS_REG_GOAL_POSITION_H         0x2B    // Goal position high byte
#define STS_REG_GOAL_PWM_L              0x2C    // Goal PWM low byte
#define STS_REG_GOAL_PWM_H              0x2D    // Goal PWM high byte
#define STS_REG_GOAL_SPEED_L            0x2E    // Goal speed low byte
#define STS_REG_GOAL_SPEED_H            0x2F    // Goal speed high byte
#define STS_REG_TORQUE_LIMIT_L          0x30    // Torque limit low byte
#define STS_REG_TORQUE_LIMIT_H          0x31    // Torque limit high byte
// REG 0x32 to 0x36 are reserved
#define STS_REG_LOCK_SYMBOL             0x37    // Lock symbol

// SRAM Feedback registers (read-only)
#define STS_REG_CURRENT_POSITION_L      0x38    // Current position low byte
#define STS_REG_CURRENT_POSITION_H      0x39    // Current position high
#define STS_REG_CURRENT_SPEED_L         0x3A    // Current speed low byte
#define STS_REG_CURRENT_SPEED_H         0x3B    // Current speed
#define STS_REG_CURRENT_LOAD_L          0x3C    // Current load low byte
#define STS_REG_CURRENT_LOAD_H          0x3D    // Current load
#define STS_REG_CURRENT_VOLTAGE         0x3E    // Current voltage
#define STS_REG_CURRENT_TEMPERATURE     0x3F    // Current temperature
#define STS_REG_ASYNC_FLAG              0x40    // Async flag
#define STS_REG_SERVO_STATUS            0x41    // Servomotor status
#define STS_REG_MOVING                  0x42    // Moving status
#define STS_REG_TARGET_LOCATION_L       0x43    // Target location low byte
#define STS_REG_TARGET_LOCATION_H       0x44    // Target location high
#define STS_REG_CURRENT_CURRENT_L       0x45    // Current current low byte
#define STS_REG_CURRENT_CURRENT_H       0x46    // Current current high
// REG 0x47 to 0x49 are reserved

// Factory Parameters (read-only)
// REG 0x50 to 0x56 are register used for factory parameters but their meaning is unknown
// and have no effect on the servomotor operation
// (see datasheet for details)

// Sero status flags
typedef enum STS_StatusFlags {
    STS_STATUS_VOLTAGE_PROTECTION       = 1 << 0,   // Voltage protection flag
    STS_STATUS_ENCODER_PROTECTION       = 1 << 1,   // Magnetic encoder protection flag
    STS_STATUS_OVERHEAT_PROTECTION      = 1 << 2,   // Overheat protection flag
    STS_STATUS_OVERCURRENT_PROTECTION   = 1 << 3,   // Overcurrent protection flag
    // Not use
    STS_STATUS_OVERLOAD_PROTECTION      = 1 << 5,   // Overload protection flag
    // Not use
    // Not use
} STS_StatusFlags;


// SERIAL COMMUNICATION PROTOCOL DEFINITIONS

#define STS_HEADER_1                    0xFF    // First header byte
#define STS_HEADER_2                    0xFF    // Second header byte

// Instruction code
#define STS_INST_PING                   0x01    // Ping instruction
#define STS_INST_READ                   0x02    // Read instruction
#define STS_INST_WRITE                  0x03    // Write instruction
#define STS_INST_REG_WRITE              0x04    // Reg write instruction
#define STS_INST_ACTION                 0x05    // Action instruction
#define STS_INST_SYNC_READ              0x82    // Sync read instruction
#define STS_INST_SYNC_WRITE             0x83    // Sync write instruction
#define STS_INST_RESET                  0x0A    // Reset instruction
#define STS_INST_POSITION_CALIBRATION   0x0B    // Position calibration instruction
#define STS_INST_RESET_PARAMETERS       0x06    // Reset parameters instruction
#define STS_INST_SAVE_PARAMETERS        0x09    // Save parameters instruction
#define STS_INST_REBOOT                 0x08    // Reboot instruction


// Constante
#define STS_UNIT_TO_DEGREE              (360.0f/4096.0f)    // Conversion factor from unit to degree (1 unit = 0.088 degree)
#define STS_UNIT_TO_RPM_1               0.732               // Conversion factor from unit to RPM (1 unit = 0.7320 RPM)
#define STS_UNIT_TO_RPM_2               0.0146              // Conversion factor from unit to RPM (1 unit = 0.0146 RPM)
#define STS_UNIT_TO_TORQUE              (19.5f/1000.0f)     // Conversion factor from unit to Torque in Ncm (1 unit = 0.0195 kg.cm)
#define STS_UNIT_TO_TEMPERATURE         1                   // Conversion factor from unit to degree Celsius (1 unit = 1 degree Celsius)
#define STS_UNIT_TO_VOLTAGE             0.1                 // Conversion factor from unit to Volt (1 unit = 0.1 Volt)
#define STS_UNIT_TO_CURRENT             6.5                 // Conversion factor from unit to mA (1 unit = 6.5 mA)

#define STS_SERIAL_BUFFER_SIZE          256      // Size of the serial communication buffer
#define STS_TIMEOUT_MS                   25      // Timeout for serial communication in milliseconds

typedef struct STS_UART_Port_t {
    UART_HandleTypeDef *huart;  // UART handle
    uint8_t __rx_buffer[STS_SERIAL_BUFFER_SIZE]; // Internal RX buffer

    bool rx_complete;          // RX complete flag
} STS_UART_Port_t;

extern STS_UART_Port_t huart_sts_port1;

typedef struct STS_Servo_t {
    STS_UART_Port_t *uart_port; // UART port
    uint8_t id;                 // Servo ID
} STS_Servo_t;

typedef struct STS_Servo_Current_raw_t {
    uint16_t position;          // Current position
    int16_t speed;              // Current speed
    int16_t load;               // Current load
    uint8_t voltage;            // Current voltage
    uint8_t temperature;        // Current temperature
    uint8_t current;            // Current current
} STS_Servo_Current_raw_t;

typedef struct STS_Servo_Current_t {
    float position;         // Current position (in degrees)
    float speed;            // Current speed (in degrees per second)
    float load;             // Current load (in kg.cm)
    float voltage;          // Current voltage (in Volts)
    float temperature;      // Current temperature (in degree Celsius)
    float current;          // Current current (in mA)
} STS_Servo_Current_t;


HAL_StatusTypeDef STS_UART_Port_Init(STS_UART_Port_t *uart_port, UART_HandleTypeDef *huart);

bool STS_UART_Port_VerifyHeader(STS_UART_Port_t *uart_port);
uint8_t STS_UART_Port_GetID(STS_UART_Port_t *uart_port);
uint8_t STS_UART_Port_GetLength(STS_UART_Port_t *uart_port);
uint8_t STS_UART_Port_GetStatus(STS_UART_Port_t *uart_port);
bool STS_UART_Port_VerifyChecksum(STS_UART_Port_t *uart_port, uint16_t data_length);

void STS_UART_Port_Callback_RX_IRQHandler(STS_UART_Port_t *uart_port, uint16_t Size);

bool STS_UART_Port_IsRXComplete(STS_UART_Port_t *uart_port);




HAL_StatusTypeDef STS_Servo_Init(STS_Servo_t *servo, STS_UART_Port_t *uart_port, uint8_t id);

bool STS_Servo_IsPacketValide(STS_Servo_t *servo, uint8_t expected_data_length);

HAL_StatusTypeDef STS_Servo_SendInstruction(STS_Servo_t *servo, uint8_t instruction, uint8_t *params, uint16_t params_length);
HAL_StatusTypeDef STS_Servo_ReceiveStatus(STS_Servo_t *servo, uint8_t *status, uint8_t *data, uint16_t data_length);

HAL_StatusTypeDef STS_Servo_Ping(STS_Servo_t *servo);
HAL_StatusTypeDef STS_Servo_ReadRegister(STS_Servo_t *servo, uint8_t reg_addr, uint8_t *data, uint16_t length);
HAL_StatusTypeDef STS_Servo_WriteRegister(STS_Servo_t *servo, uint8_t reg_addr, uint8_t *data, uint16_t length);
HAL_StatusTypeDef STS_Servo_Reset(STS_Servo_t *servo);
HAL_StatusTypeDef STS_Servo_PositionCalibration(STS_Servo_t *servo);
HAL_StatusTypeDef STS_Servo_ResetParameters(STS_Servo_t *servo);
HAL_StatusTypeDef STS_Servo_SaveParameters(STS_Servo_t *servo);
HAL_StatusTypeDef STS_Servo_Reboot(STS_Servo_t *servo);

HAL_StatusTypeDef STS_Servo_GetCurrentStatus(STS_Servo_t *servo, STS_Servo_Current_raw_t *current_status);
HAL_StatusTypeDef STS_Servo_SetGoalPosition(STS_Servo_t *servo, uint16_t position);
HAL_StatusTypeDef STS_Servo_SetGoalSpeed(STS_Servo_t *servo, int16_t speed);
HAL_StatusTypeDef STS_Servo_SetGoalLoad(STS_Servo_t *servo, int16_t load);
HAL_StatusTypeDef STS_Servo_IsMoving(STS_Servo_t *servo, uint8_t *is_moving);

void STS_Servo_raw_to_physical(const STS_Servo_Current_raw_t *raw, STS_Servo_Current_t *physical);
void STS_Servo_physical_to_raw(const STS_Servo_Current_t *physical, STS_Servo_Current_raw_t *raw);

#ifdef __cplusplus
}
#endif
#endif /* STS3215_H */