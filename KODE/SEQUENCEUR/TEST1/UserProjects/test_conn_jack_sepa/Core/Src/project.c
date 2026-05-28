#include "project.h"

#include "main.h"
#include "stm32f0xx_hal.h"
#include "stm32f0xx_hal_gpio.h"
#include "usbd_cdc_if.h"
#include <stdio.h>


static bool jack_connected = false;
static bool stage_connected = false;

static char buffer[64];

static uint32_t t0;


void setup() {
    t0 = HAL_GetTick();
}

void loop() {

    jack_connected = HAL_GPIO_ReadPin(IN_TRG_N1_GPIO_Port, IN_TRG_N1_Pin) == GPIO_PIN_SET;
    stage_connected = HAL_GPIO_ReadPin(IN_TRG_N2_GPIO_Port, IN_TRG_N2_Pin) == GPIO_PIN_SET;   

    snprintf(buffer, sizeof(buffer), "Jack: %s, Stage: %s\r\n", jack_connected ? "Connected" : "Disconnected", stage_connected ? "Connected" : "Disconnected");
    CDC_Transmit_FS((uint8_t*)buffer, strlen(buffer));

    HAL_GPIO_WritePin(LED1G_GPIO_Port, LED1G_Pin, jack_connected ? GPIO_PIN_RESET : GPIO_PIN_SET);
    HAL_GPIO_WritePin(LED1B_GPIO_Port, LED1B_Pin, stage_connected ? GPIO_PIN_RESET : GPIO_PIN_SET);

    if (HAL_GetTick() - t0 > 500) {
        HAL_GPIO_TogglePin(LED1R_GPIO_Port, LED1R_Pin);
        t0 = HAL_GetTick();
    }

    HAL_Delay(1);
}
