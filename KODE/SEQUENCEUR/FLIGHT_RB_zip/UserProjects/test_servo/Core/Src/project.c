#include "project.h"

#include "main.h"
#include "stm32f0xx_hal.h"
#include "stm32f0xx_hal_def.h"
#include "stm32f0xx_hal_gpio.h"
#include "stm32f0xx_hal_uart.h"
#include "usbd_cdc_if.h"

#include "usart.h"

#include "STS.h"

#include <stdint.h>
#include <stdio.h>

static char buffer[256];
static uint32_t t0;
static STS_Servo_t servo;

void setup() {

    STS_UART_Port_Init(&huart_sts_port1, &huart2); // Pour etre raccord par rapport à la serigraphie du PCB
    HAL_Delay(1);

    HAL_StatusTypeDef res = STS_Servo_Init(&servo, &huart_sts_port1, 4);
	if (res != HAL_OK) { Error_Handler(); }

    HAL_Delay(1);

    STS_Servo_SetOperatingMode(&servo, STS_OP_MODE_POSITION_CONTROL);
    HAL_Delay(1);
    STS_Servo_PositionCalibration(&servo, STS_GetPositionInUnits(180));
    HAL_Delay(1);

    STS_Servo_SetGoalPosition(&servo, STS_GetPositionInUnits(0));
    HAL_Delay(3000);

    STS_Servo_SetGoalPosition(&servo, STS_GetPositionInUnits(360));
    HAL_Delay(3000);

    STS_Servo_SetGoalPosition(&servo, STS_GetPositionInUnits(180));
    HAL_Delay(3000);
}

void loop() {

}
