#include "STS.h"

#include "usart.h"

#include <stdint.h>
#include <string.h>
#include <math.h>



STS_UART_Port_t huart_sts_port1 = { 0 };
STS_UART_Port_t huart_sts_port2 = { 0 };

uint32_t STS_GetBaudRateValue(STS_BaudRate baudrate) {
    switch (baudrate) {
        case STS_BAUDRATE_1000000   : return 1000000;
        case STS_BAUDRATE_500000    : return 500000;
        case STS_BAUDRATE_250000    : return 250000;
        case STS_BAUDRATE_128000    : return 128000;
        case STS_BAUDRATE_115200    : return 115200;
        case STS_BAUDRATE_76800     : return 76800;
        case STS_BAUDRATE_57600     : return 57600;
        case STS_BAUDRATE_38400     : return 38400;
        default                     : return 0; // Invalid baudrate
    }
}

STS_BaudRate STS_GetBaudRateFromValue(uint32_t baudrate_value) {
    switch (baudrate_value) {
        case 1000000    : return STS_BAUDRATE_1000000;
        case 500000     : return STS_BAUDRATE_500000;
        case 250000     : return STS_BAUDRATE_250000;
        case 128000     : return STS_BAUDRATE_128000;
        case 115200     : return STS_BAUDRATE_115200;
        case 76800      : return STS_BAUDRATE_76800;
        case 57600      : return STS_BAUDRATE_57600;
        case 38400      : return STS_BAUDRATE_38400;
        default         : return STS_BAUDRATE_UNDEFINED; // Invalid baudrate value
    }
}


float STS_GetPositionInDegrees(uint16_t position_units) {
    float pos = (float)(position_units & 0x7FFF) * STS_UNIT_TO_DEGREE;
    return (position_units >> 15) & 0x01 ? -pos : pos;
}
float STS_GetSpeedInRPM(int16_t speed_units) {
    float speed = (float)(speed_units & 0x7FFF) * STS_UNIT_TO_RPM_2;
    return (speed_units >> 15) & 0x01 ? -speed : speed;
}
float STS_GetLoadInKgcm(int16_t load_units) {
    float load = (float)(load_units & 0x3FF) * STS_UNIT_TO_TORQUE;
    return (load_units >> 10) & 0x01 ? -load : load;
}
float STS_GetVoltageInVolts(uint8_t voltage_units) {
    return (float)voltage_units * STS_UNIT_TO_VOLTAGE;
}
float STS_GetTemperatureInCelsius(uint8_t temperature_units) {
    return (float)temperature_units * STS_UNIT_TO_TEMPERATURE;
}
float STS_GetCurrentInmA(uint16_t current_units) {
    return (float)current_units * STS_UNIT_TO_CURRENT;
}

uint16_t STS_GetPositionInUnits(float position_degrees) {
    uint16_t position_units = (uint16_t)(fabs(position_degrees) / STS_UNIT_TO_DEGREE);
    return (position_degrees < 0) ? (position_units | 0x8000) : position_units;
}
int16_t STS_GetSpeedInUnits(float speed_rpm) {
    int16_t speed_units = (int16_t)(fabs(speed_rpm) / STS_UNIT_TO_RPM_2);
    return (speed_rpm < 0) ? (speed_units | 0x8000) : speed_units;
}
uint16_t STS_GetLoadInUnits(float load_kgcm) {
    uint16_t load_units = (uint16_t)(fabs(load_kgcm) / STS_UNIT_TO_TORQUE);
    return (load_kgcm < 0) ? (load_units | 0x400) : load_units;
}
uint8_t STS_GetVoltageInUnits(float voltage_volts) {
    return (uint8_t)(voltage_volts / STS_UNIT_TO_VOLTAGE);
}
uint8_t STS_GetTemperatureInUnits(float temperature_celsius) {
    return (uint8_t)(temperature_celsius / STS_UNIT_TO_TEMPERATURE);
}
uint16_t STS_GetCurrentInUnits(float current_mA) {
    return (uint16_t)(current_mA / STS_UNIT_TO_CURRENT);
}


uint8_t __compute_checksum(uint8_t *data, uint8_t length) {
  uint8_t sum = 0;
  for (uint8_t i = 0; i < length; i++) {
    sum += data[i];
  }
  return ~sum;
}


HAL_StatusTypeDef STS_UART_Port_Init(STS_UART_Port_t *uart_port, UART_HandleTypeDef *huart) {
    uart_port->huart = huart;
    memset(uart_port->__rx_buffer, 0, STS_SERIAL_BUFFER_SIZE);
    uart_port->rx_complete = false;

    UART_buffer_t *buffer_obj;
    UART_get_buffer(huart, &buffer_obj);
    if (buffer_obj == NULL) {
        return HAL_ERROR; // Unknown UART instance
    }
    buffer_obj->rx_buffer = uart_port->__rx_buffer;
    buffer_obj->rx_length = STS_SERIAL_BUFFER_SIZE;

    HAL_StatusTypeDef res = HAL_HalfDuplex_EnableReceiver(huart);
    if (res != HAL_OK) {
        return res;
    }
    res = HAL_UARTEx_ReceiveToIdle_IT(uart_port->huart, uart_port->__rx_buffer, STS_SERIAL_BUFFER_SIZE);
    if (res != HAL_OK) {
        return res;
    }

    return HAL_OK;
}

bool STS_UART_Port_VerifyHeader(STS_UART_Port_t *uart_port) {
    return (uart_port->__rx_buffer[0] == STS_HEADER_1 && uart_port->__rx_buffer[1] == STS_HEADER_2);
}

uint8_t STS_UART_Port_GetID(STS_UART_Port_t *uart_port) {
    return uart_port->__rx_buffer[2];
}

uint8_t STS_UART_Port_GetLength(STS_UART_Port_t *uart_port) {
    return uart_port->__rx_buffer[3];
}

uint8_t STS_UART_Port_GetStatus(STS_UART_Port_t *uart_port) {
    return uart_port->__rx_buffer[4];
}

bool STS_UART_Port_VerifyChecksum(STS_UART_Port_t *uart_port, uint16_t data_length) {
    uint8_t received_checksum = uart_port->__rx_buffer[5 + data_length]; // Header(2) + ID(1) + Length(1) + Status(1) + Data(n)
    uint8_t computed_checksum = __compute_checksum(&uart_port->__rx_buffer[2], 3 + data_length); // ID(1) + Length(1) + Status(1) + Data(n)
    return (received_checksum == computed_checksum);
}


void STS_UART_Port_Callback_RX_IRQHandler(STS_UART_Port_t *uart_port, uint16_t Size) {
    uart_port->rx_complete = true;
}

bool STS_UART_Port_IsRXComplete(STS_UART_Port_t *uart_port) {
    return uart_port->rx_complete;
}





HAL_StatusTypeDef STS_Servo_Init(STS_Servo_t *servo, STS_UART_Port_t *uart_port, uint8_t id) {
    servo->uart_port = uart_port;
    servo->id = id;

    // Send a first ping to initialize communication (discard result since servo might be in unknown state,
    // but it will help to "wake up" the servo and get it ready for the next commands)
    STS_Servo_Ping(servo);
    HAL_Delay(10); // Small delay

    // Now send a ping and check response to verify that communication is working
    return STS_Servo_Ping(servo);
}


bool STS_Servo_IsPacketValide(STS_Servo_t *servo, uint8_t expected_data_length) {
    bool header_valid = STS_UART_Port_VerifyHeader(servo->uart_port);
    bool id_valid = (STS_UART_Port_GetID(servo->uart_port) == servo->id);
    bool length_valid = (STS_UART_Port_GetLength(servo->uart_port) == expected_data_length + 2); // Length(n+2) = Status(1) + Data(n) + Checksum(1)
    bool checksum_valid = STS_UART_Port_VerifyChecksum(servo->uart_port, expected_data_length);

    return (header_valid && id_valid && length_valid && checksum_valid);
}


HAL_StatusTypeDef STS_Servo_SendInstruction(STS_Servo_t *servo, uint8_t instruction, uint8_t *params, uint16_t params_length) {
    uint8_t packet[STS_SERIAL_BUFFER_SIZE];
    uint16_t packet_length = 0;

    // Construct packet
    packet[packet_length++] = STS_HEADER_1;
    packet[packet_length++] = STS_HEADER_2;
    packet[packet_length++] = servo->id;
    packet[packet_length++] = (uint8_t)(params_length + 2); // Length(n+2) = Instruction(1) + Params(n) + Checksum(1)
    packet[packet_length++] = instruction;

    // Add parameters
    if (params_length > 0 && params != NULL) {
        memcpy(packet + packet_length, params, params_length);
        packet_length += params_length;
    }

    // Compute checksum
    uint8_t checksum = __compute_checksum(&packet[2], 3 + params_length); // ID(1) + Length(1) + Instruction(1) + Params(n)
    packet[packet_length++] = checksum;

    // Transmit packet

    HAL_StatusTypeDef res = HAL_HalfDuplex_EnableTransmitter(servo->uart_port->huart);
    if (res != HAL_OK) {
        return res;
    }
    res = HAL_UART_Transmit(servo->uart_port->huart, packet, packet_length, HAL_MAX_DELAY);
    if (res != HAL_OK) {
        return res;
    }
    return HAL_HalfDuplex_EnableReceiver(servo->uart_port->huart);
}

HAL_StatusTypeDef STS_Servo_ReceiveStatus(STS_Servo_t *servo, uint8_t *status, uint8_t *data, uint16_t data_length) {
    uint32_t timeout = HAL_GetTick() + 10; // Max 500 us timeout so 1 ms is large enough
    uint32_t t0 = HAL_GetTick();
    HAL_UARTEx_ReceiveToIdle_IT(servo->uart_port->huart, servo->uart_port->__rx_buffer, STS_SERIAL_BUFFER_SIZE);
    while (servo->uart_port->rx_complete == false && HAL_GetTick() < timeout) {
        if (HAL_GetTick() - t0 > 1) { // After 1 ms, if we still haven't received anything, we can timeout
            HAL_UARTEx_ReceiveToIdle_IT(servo->uart_port->huart, servo->uart_port->__rx_buffer, STS_SERIAL_BUFFER_SIZE);
        }
    }
    if (servo->uart_port->rx_complete == false) {
        return HAL_TIMEOUT; // Timeout
    }
    servo->uart_port->rx_complete = false; // Reset flag

    // Check packet validity
    if (!STS_Servo_IsPacketValide(servo, data_length)) {
        return HAL_ERROR; // Invalid packet
    }

    // Extract status and data
    *status = servo->uart_port->__rx_buffer[4];
    if (data != NULL && data_length > 0) {
        memcpy(data, &servo->uart_port->__rx_buffer[5], data_length);
    }

    return HAL_OK;
}



// High-level functions (Ping, ReadRegister, WriteRegister, etc.)

HAL_StatusTypeDef STS_Servo_Ping(STS_Servo_t *servo) {
    HAL_StatusTypeDef res = STS_Servo_SendInstruction(servo, STS_INST_PING, NULL, 0);
    if (res != HAL_OK) {
        return res;
    }
    uint8_t status;
    res = STS_Servo_ReceiveStatus(servo, &status, NULL, 0);
    if (res != HAL_OK) {
        return res;
    }
    return (status == 0 && res == HAL_OK) ? HAL_OK : res;
}

HAL_StatusTypeDef STS_Servo_ReadRegister(STS_Servo_t *servo, uint8_t reg_addr, uint8_t *data, uint16_t length) {
    uint8_t params[2] = { reg_addr, (uint8_t)length };
    HAL_StatusTypeDef res = STS_Servo_SendInstruction(servo, STS_INST_READ, params, 2);
    if (res != HAL_OK) {
        return res;
    }
    uint8_t status;
    res = STS_Servo_ReceiveStatus(servo, &status, data, length);
    return (status == 0 && res == HAL_OK) ? HAL_OK : res;

}

HAL_StatusTypeDef STS_Servo_WriteRegister(STS_Servo_t *servo, uint8_t reg_addr, uint8_t *data, uint16_t length) {
    uint8_t params[STS_SERIAL_BUFFER_SIZE];
    params[0] = reg_addr;
    memcpy(&params[1], data, length);
    return STS_Servo_SendInstruction(servo, STS_INST_WRITE, params, length + 1);
}

HAL_StatusTypeDef STS_Servo_Reset(STS_Servo_t *servo) {
    return STS_Servo_SendInstruction(servo, STS_INST_RESET, NULL, 0);
}

HAL_StatusTypeDef STS_Servo_PositionCalibration(STS_Servo_t *servo, uint16_t position_units) {
    HAL_StatusTypeDef res;
    uint8_t params[2] = { position_units, position_units >> 8 };
    // res = STS_Servo_SendInstruction(servo, STS_INST_POSITION_CALIBRATION, params, 2);
    
    uint8_t lock = 0;
    res = STS_Servo_WriteRegister(servo, STS_REG_LOCK_SYMBOL, &lock, 1);
    HAL_Delay(1);
    
    uint8_t status;
    res = STS_Servo_SendInstruction(servo, STS_INST_POSITION_CALIBRATION, params, 2);
    // res = STS_Servo_SendInstruction(servo, STS_INST_POSITION_CALIBRATION, NULL, 0);
    res = STS_Servo_ReceiveStatus(servo, &status, NULL, 0);
    HAL_Delay(1);
    
    lock = 1;
    res = STS_Servo_WriteRegister(servo, STS_REG_LOCK_SYMBOL, &lock, 1);
    
    return (status == 0 && res == HAL_OK) ? HAL_OK : HAL_ERROR;
}

HAL_StatusTypeDef STS_Servo_ResetParameters(STS_Servo_t *servo) {
    return STS_Servo_SendInstruction(servo, STS_INST_RESET_PARAMETERS, NULL, 0);
}

HAL_StatusTypeDef STS_Servo_SaveParameters(STS_Servo_t *servo) {
    return STS_Servo_SendInstruction(servo, STS_INST_SAVE_PARAMETERS, NULL, 0);
}

HAL_StatusTypeDef STS_Servo_Reboot(STS_Servo_t *servo) {
    return STS_Servo_SendInstruction(servo, STS_INST_REBOOT, NULL, 0);
}


// Get current status (position, speed, load, voltage, temperature)
HAL_StatusTypeDef STS_Servo_GetCurrentStatus(STS_Servo_t *servo, STS_Servo_Current_raw_t *current_status) {
    uint8_t data[8];
    HAL_StatusTypeDef res = STS_Servo_ReadRegister(servo, STS_REG_CURRENT_POSITION_L, data, 8);
    if (res != HAL_OK) {
        return res;
    }

    current_status->position = (uint16_t)(data[0] | (data[1] << 8));
    current_status->speed = (int16_t)(data[2] | (data[3] << 8));
    current_status->load = (int16_t)(data[4] | (data[5] << 8));
    current_status->voltage = data[6];
    current_status->temperature = data[7];

    HAL_Delay(1); // Small delay before reading current

    res = STS_Servo_ReadRegister(servo, STS_REG_CURRENT_CURRENT_L, data, 2);
    if (res != HAL_OK) {
        return res;
    }
    current_status->current = (uint16_t)(data[0] | (data[1] << 8));

    return HAL_OK;
}

HAL_StatusTypeDef STS_Servo_SetGoalPosition(STS_Servo_t *servo, uint16_t position) {
    uint8_t data[2];
    data[0] = (uint8_t)(position & 0xFF);
    data[1] = (uint8_t)((position >> 8) & 0xFF);
    return STS_Servo_WriteRegister(servo, STS_REG_GOAL_POSITION_L, data, 2);
}

HAL_StatusTypeDef STS_Servo_SetGoalSpeed(STS_Servo_t *servo, int16_t speed) {
    uint8_t data[2];
    data[0] = (uint8_t)(speed & 0xFF);
    data[1] = (uint8_t)((speed >> 8) & 0xFF);
    return STS_Servo_WriteRegister(servo, STS_REG_GOAL_SPEED_L, data, 2);
}

HAL_StatusTypeDef STS_Servo_SetGoalLoad(STS_Servo_t *servo, int16_t load) {
    uint8_t data[2];
    data[0] = (uint8_t)(load & 0xFF);
    data[1] = (uint8_t)((load >> 8) & 0xFF);
    return STS_Servo_WriteRegister(servo, STS_REG_TORQUE_LIMIT_L, data, 2);
}

HAL_StatusTypeDef STS_Servo_SetOperatingMode(STS_Servo_t *servo, STS_OperatingMode mode) {
    uint8_t data = (uint8_t)mode;
    return STS_Servo_WriteRegister(servo, STS_REG_OPERATING_MODE, &data, 1);
}

HAL_StatusTypeDef STS_Servo_IsMoving(STS_Servo_t *servo, uint8_t *is_moving) {
    return STS_Servo_ReadRegister(servo, STS_REG_MOVING, is_moving, 1);
}

HAL_StatusTypeDef STS_Servo_InOverload(STS_Servo_t *servo, bool *in_overload) {
    uint8_t status;
    HAL_StatusTypeDef ret = STS_Servo_ReadRegister(servo, STS_REG_SERVO_STATUS, &status, 1);
    *in_overload = (status & 0x20) != 0;
    return ret;
}


void STS_raw_to_physical(const STS_Servo_Current_raw_t *raw, STS_Servo_Current_t *physical) {
    physical->position    = STS_GetPositionInDegrees(raw->position);
    physical->speed       = STS_GetSpeedInRPM(raw->speed);
    physical->load        = STS_GetLoadInKgcm(raw->load);
    physical->voltage     = STS_GetVoltageInVolts(raw->voltage);
    physical->temperature = STS_GetTemperatureInCelsius(raw->temperature);
    physical->current     = STS_GetCurrentInmA(raw->current);
}

void STS_physical_to_raw(const STS_Servo_Current_t *physical, STS_Servo_Current_raw_t *raw) {
    raw->position    = STS_GetPositionInUnits(physical->position);
    raw->speed       = STS_GetSpeedInUnits(physical->speed);
    raw->load        = STS_GetLoadInUnits(physical->load);
    raw->voltage     = STS_GetVoltageInUnits(physical->voltage);
    raw->temperature = STS_GetTemperatureInUnits(physical->temperature);
    raw->current     = STS_GetCurrentInUnits(physical->current);
}
