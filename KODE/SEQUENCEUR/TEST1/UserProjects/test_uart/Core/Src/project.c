#include "project.h"

#include "main.h"

#include "stm32f0xx_hal.h"
#include "stm32f0xx_hal_gpio.h"
#include "usart.h"

bool flag_rx = false;

uint8_t uart_rx_buffer[64]; // Buffer de réception pour les données UART

void setup() {

    uart_buffer_4.rx_buffer = uart_rx_buffer;
    uart_buffer_4.rx_length = sizeof(uart_rx_buffer);

    HAL_UARTEx_ReceiveToIdle_IT(&huart4, uart_rx_buffer, sizeof(uart_rx_buffer));

    HAL_GPIO_WritePin(TX_OPTO_N1_GPIO_Port, TX_OPTO_N1_Pin, GPIO_PIN_RESET);

}

static uint32_t t0 = 0;

void loop() {
    if (flag_rx) {
        HAL_GPIO_TogglePin(LED1B_GPIO_Port, LED1B_Pin);
        HAL_Delay(100);
        HAL_GPIO_TogglePin(LED1B_GPIO_Port, LED1B_Pin);
        HAL_UART_Transmit(&huart4, (const uint8_t *)"Hello, from sequenceur!\r\n", 26, HAL_MAX_DELAY);
        HAL_GPIO_WritePin(TX_OPTO_N1_GPIO_Port, TX_OPTO_N1_Pin, GPIO_PIN_RESET);
        flag_rx = false;
    }

    if (HAL_GetTick() - t0 >= 3000) {
        HAL_GPIO_WritePin(TX_OPTO_N1_GPIO_Port, TX_OPTO_N1_Pin, GPIO_PIN_SET);
        t0 = HAL_GetTick();
    }
}