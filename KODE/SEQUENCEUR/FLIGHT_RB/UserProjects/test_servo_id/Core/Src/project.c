#include "project.h"

#include "main.h"
#include "usbd_cdc_if.h"

#include "usart.h"

#include "STS.h"

#include <stdint.h>
#include <stdio.h>

static char buffer[256];
static uint32_t t0;
static STS_UART_Port_t *port;

static uint8_t uart2_buffer[STS_SERIAL_BUFFER_SIZE];
static uint8_t uart3_buffer[STS_SERIAL_BUFFER_SIZE];

static bool active_id[2][256] = { false }; // active_id[port][baudrate][id]

void setup() {

    HAL_Delay(10000);

	UART_buffer_init(&huart2, uart2_buffer, sizeof(uart2_buffer));
	UART_buffer_init(&huart3, uart3_buffer, sizeof(uart3_buffer));

    STS_UART_Port_Init(&huart_sts_port1, &huart2); // Pour etre raccord par rapport à la serigraphie du PCB
    STS_UART_Port_Init(&huart_sts_port2, &huart3); // Pour etre raccord par rapport à la serigraphie du PCB
    t0 = HAL_GetTick();

    for (int i_port = 0; i_port < 2; i_port++) {
        // Select UART port
        port = i_port == 0 ? &huart_sts_port1 : &huart_sts_port2;
        for (int id = 0; id < 256; id++) {
            // Init servo struct-
            STS_Servo_t servo;
            // Ping servo
            snprintf(buffer, sizeof(buffer), "Port %d; ID %d: ", i_port + 1, id);
            if (STS_Servo_Init(&servo, port, id) == HAL_OK) {
                active_id[i_port][id] = true;
                snprintf(buffer, sizeof(buffer), "%sFOUND", buffer);
            }
            snprintf(buffer, sizeof(buffer), "%s\r\n", buffer);
            CDC_Transmit_FS((uint8_t*)buffer, strlen(buffer));
        }
    }
}

void loop() {

    if (HAL_GetTick() - t0 > 500) {
        HAL_GPIO_TogglePin(LED1R_GPIO_Port, LED1R_Pin);
        t0 = HAL_GetTick();
    }

    HAL_Delay(1);

    active_id;
}
