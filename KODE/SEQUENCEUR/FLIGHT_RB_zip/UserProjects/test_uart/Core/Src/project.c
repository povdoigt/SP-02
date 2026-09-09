#include "project.h"

#include "event_uart.h"
#include "main.h"

#include "stm32f0xx_hal_gpio.h"
#include "usart.h"
#include <string.h>


bool flag_rx = false;
bool flag_tx = false;

uint8_t uart4_buffer[sizeof(event_uart_msg_t)];

void setup() {

    UART_buffer_init(&huart4, uart4_buffer, sizeof(uart4_buffer));

	event_uart_producer_init(&event_uart_producer, TX_OPTO_N1_GPIO_Port, TX_OPTO_N1_Pin, &huart4);

}

static uint32_t t0 = 0;
static uint32_t new_state = 0;

void loop() {

    if (HAL_GetTick() - t0 >= 1000) {
        event_uart_producer_add_event(&event_uart_producer, event_uart_msg_format(
            HAL_GetTick(),
            EVENT_UART_TYPE_STATE_MACHINE_STATE_CHANGE,
			(event_uart_payload_u){.state_change_payload = {
				.state_machine_id = 0,
				.old_state_id = new_state++,
				.new_state_id = new_state
			}}
		));
        HAL_GPIO_TogglePin(LED1G_GPIO_Port, LED1G_Pin);
        HAL_Delay(100);
        HAL_GPIO_TogglePin(LED1G_GPIO_Port, LED1G_Pin);
        t0 = HAL_GetTick();
    }

    event_uart_producer_send_events(&event_uart_producer);
}
