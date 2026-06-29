#include "project.h"

#include "main.h"
#include "stm32f0xx_hal.h"
#include "stm32f0xx_hal_gpio.h"

void setup() {

    for (size_t i = 0; i < 10; i++) {
        HAL_GPIO_TogglePin(LED1B_GPIO_Port, LED1B_Pin);
        HAL_Delay(500);
    }

    HAL_GPIO_WritePin(LED1R_GPIO_Port, LED1R_Pin, GPIO_PIN_RESET);
    HAL_GPIO_WritePin(OUT_N2_GPIO_Port, OUT_N2_Pin, GPIO_PIN_SET);

}

void loop() {

}
