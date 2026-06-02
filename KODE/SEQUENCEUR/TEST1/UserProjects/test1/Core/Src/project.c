#include "project.h"

#include "STS.h"
#include "cmsis_gcc.h"
#include "stm32f0xx_hal.h"
#include "usart.h"

STS_Servo_t servo1, servo2;

void setup() {
    
    HAL_Delay(2000); // Wait for peripherals to stabilize

	HAL_StatusTypeDef status;

	status = STS_UART_Port_Init(&huart_sts_port2, &huart3);

	status = STS_Servo_Init(&servo1, &huart_sts_port2, 1);
	HAL_Delay(1);
	status = STS_Servo_Init(&servo2, &huart_sts_port2, 2);
	HAL_Delay(1);

    __NOP();

    // uint8_t protection_time = 2; // 2 = 20ms
    // status = STS_Servo_WriteRegister(&servo1, STS_REG_PROTECTION_TIME, &protection_time, sizeof(protection_time));
	// HAL_Delay(1);

    // uint8_t overload_threshold = 80; // 80 % of max torque
    // status = STS_Servo_WriteRegister(&servo1, STS_REG_OVERLOAD_TORQUE, &overload_threshold, sizeof(overload_threshold));
    // HAL_Delay(1);

    // uint8_t maintain_torque = 20; // 20 % of max torque
    // status = STS_Servo_WriteRegister(&servo1, STS_REG_MAINTAIN_TORQUE, &maintain_torque, sizeof(maintain_torque));
    // HAL_Delay(1);


    // // Invert direction of servo 1
    // uint8_t phase;
    // status = STS_Servo_ReadRegister(&servo1, STS_REG_PHASE, &phase, sizeof(phase));
    // HAL_Delay(1);
    // phase |= 0x01; // Set bit 0 to invert direction
    // status = STS_Servo_WriteRegister(&servo1, STS_REG_PHASE, &phase, sizeof(phase));
    // HAL_Delay(1);



    // bool overload = false;

    // status = STS_Servo_SetOperatingMode(&servo1, STS_OP_MODE_SPEED_CONTROL);
    // HAL_Delay(1);

    // status = STS_Servo_SetGoalSpeed(&servo1, STS_Servo_GetSpeedInUnits(-10.0f)); // 10 RPM
    // HAL_Delay(1);

    // // Wait for a bit and check overload status
    // while (!overload) {
    //     STS_Servo_InOverload(&servo1, &overload);
    //     HAL_Delay(10); // Check every 10ms
    // }

    // // Stop the servo
    // status = STS_Servo_SetGoalSpeed(&servo1, 0);
    // HAL_Delay(1);


	STS_Servo_SetOperatingMode(&servo1, STS_OP_MODE_POSITION_CONTROL);
	HAL_Delay(1);
	STS_Servo_PositionCalibration(&servo1, 0);
    HAL_Delay(1);
    STS_Servo_SetGoalPosition(&servo1, STS_Servo_GetPositionInUnits(30.0f));
    HAL_Delay(1);

}

void loop() {
    // Main loop code here
}