#include "project.h"

#include "usart.h"


uint8_t uart_rx_buffer[64]; // Buffer de réception pour les données UART

void setup() {
    HAL_UARTEx_ReceiveToIdle_DMA(&huart4, uart_rx_buffer, sizeof(uart_rx_buffer));
}

void loop() {
    __NOP();
}