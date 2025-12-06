#include "STS.h"
#include "stm32l4xx_hal_def.h"
#include <string.h>



uint8_t __compute_checksum(uint8_t *data, uint8_t length) {
  uint8_t sum = 0;
  for (uint8_t i = 0; i < length; i++) {
    sum += data[i];
  }
  return ~sum;
}



void STS_Servo_Init(STS_Servo_t *servo, UART_HandleTypeDef *huart, uint8_t id) {
    servo->huart = huart;
    servo->id = id;
}


HAL_StatusTypeDef STS_Servo_SendInstruction(STS_Servo_t *servo, uint8_t instruction, uint8_t *params, uint16_t params_length) {
    uint8_t packet[STS_SERIAL_BUFFER_SIZE];
    uint16_t packet_length = 0;

    // Construct packet
    packet[packet_length++] = STS_HEADER_1;
    packet[packet_length++] = STS_HEADER_2;
    packet[packet_length++] = servo->id;
    packet[packet_length++] = (uint8_t)(params_length + 2); // Length
    packet[packet_length++] = instruction;

    // Add parameters
    memcpy(packet + packet_length, params, params_length);

    // Compute checksum
    uint8_t checksum = __compute_checksum(&packet[2], params_length + 3); // ID + Length + Instruction + Params
    packet[packet_length++] = checksum;

    // Transmit packet
    return HAL_UART_Transmit(servo->huart, packet, packet_length, HAL_MAX_DELAY);
}

HAL_StatusTypeDef STS_Servo_ReceiveStatus(STS_Servo_t *servo, uint8_t *status, uint8_t *data, uint16_t data_length) {
    uint8_t rx_buffer[STS_SERIAL_BUFFER_SIZE];
    uint16_t expected_length = 6 + data_length; // Header(2) + ID(1) + Length(1) + Status(1) + Data + Checksum(1)

    // Receive packet
    HAL_StatusTypeDef res = HAL_UART_Receive(servo->huart, rx_buffer, expected_length, HAL_MAX_DELAY);
    if (res != HAL_OK) {
        return res;
    }

    // Verify checksum
    uint8_t received_checksum = rx_buffer[expected_length - 1];
    uint8_t computed_checksum = __compute_checksum(&rx_buffer[2], expected_length - 3); // ID + Length + Status + Data
    if (received_checksum != computed_checksum) {
        return HAL_ERROR; // Checksum mismatch
    }

    // Extract status and data
    *status = rx_buffer[4];
    memcpy(data, &rx_buffer[5], data_length);
    return HAL_OK;
}



// High-level functions (Ping, ReadRegister, WriteRegister, etc.)

HAL_StatusTypeDef STS_Servo_Ping(STS_Servo_t *servo) {
    return STS_Servo_SendInstruction(servo, STS_INST_PING, NULL, 0);
}

HAL_StatusTypeDef STS_Servo_ReadRegister(STS_Servo_t *servo, uint8_t reg_addr, uint8_t *data, uint16_t length) {
    uint8_t params[2] = { reg_addr, (uint8_t)length };
    HAL_StatusTypeDef res = STS_Servo_SendInstruction(servo, STS_INST_READ, params, 2);
    if (res != HAL_OK) {
        return res;
    }
    uint8_t status;
    return STS_Servo_ReceiveStatus(servo, &status, data, length);
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

HAL_StatusTypeDef STS_Servo_PositionCalibration(STS_Servo_t *servo) {
    return STS_Servo_SendInstruction(servo, STS_INST_POSITION_CALIBRATION, NULL, 0);
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
HAL_StatusTypeDef STS_Servo_GetCurrentStatus(STS_Servo_t *servo, STS_Servo_Current_t *current_status) {
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

    return HAL_OK;
}
