#include "project.h"

#include "main.h"
#include "stm32f0xx_hal.h"
#include "usbd_cdc_if.h"
#include <stdio.h>

void setup() {

}

static bool jack_connected = false;
static bool stage_connected = false;

static char buffer[64];

void loop() {

    jack_connected = HAL_GPIO_ReadPin(IN_TRG_N1_GPIO_Port, IN_TRG_N1_Pin) == GPIO_PIN_SET;
    stage_connected = HAL_GPIO_ReadPin(IN_TRG_N2_GPIO_Port, IN_TRG_N2_Pin) == GPIO_PIN_SET;   

    snprintf(buffer, sizeof(buffer), "Jack: %s, Stage: %s\r\n", jack_connected ? "Connected" : "Disconnected", stage_connected ? "Connected" : "Disconnected");

    CDC_Transmit_FS((uint8_t *)buffer, strlen(buffer));

    HAL_Delay(1);
}
