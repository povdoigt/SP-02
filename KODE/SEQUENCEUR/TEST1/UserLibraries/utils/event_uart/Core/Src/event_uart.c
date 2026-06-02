#include "event_uart.h"

#if EVENT_UART_PRODUCER


event_uart_msg_t event_uart_msg_format(uint32_t timestamp, event_uart_type_t type, event_uart_payload_u payload) {
    return (event_uart_msg_t){
        .header = EVENT_UART_HEADER_MAGIC,
        .timestamp = timestamp,
        .type = type,
        .payload = payload
    };
}

event_uart_producer_t event_uart_producer;

void event_uart_producer_init(event_uart_producer_t *producer,
								GPIO_TypeDef *gpio_port, uint16_t gpio_pin,
								UART_HandleTypeDef *huart) {
    producer->gpio_port = gpio_port;
    producer->gpio_pin = gpio_pin;
    producer->huart = huart;
    cb_init(&producer->cb, producer->queue, sizeof(event_uart_msg_t),
            EVENT_UART_MAX_COUNT, CB_OVERWRITE_OLDEST);
}

void event_uart_producer_add_event(event_uart_producer_t *producer, event_uart_msg_t event_msg) {
    cb_push(&producer->cb, &event_msg);
}

void event_uart_producer_send_events(event_uart_producer_t *producer) {
    if (producer->cb.count > 0) {
        if (!producer->notif_sent) {
            producer->ready_receive = false; // Reset the ready flag before sending notification
            // Notify the consumer by toggling the GPIO pin
            HAL_GPIO_WritePin(producer->gpio_port, producer->gpio_pin, GPIO_PIN_SET);
            producer->notif_sent = true;
        } else if (producer->ready_receive) {
            // Send the next event message over UART
            event_uart_msg_t event_msg;
            cb_pop(&producer->cb, &event_msg);
            HAL_UART_Transmit(producer->huart, (uint8_t *)&event_msg, sizeof(event_uart_msg_t), HAL_MAX_DELAY);
            producer->notif_sent = false; // Reset the notification flag after sending
        }   
    }
}

void event_uart_callback(void) {
    event_uart_producer.ready_receive = true;
}

#endif /* EVENT_UART_PRODUCER */

#if EVENT_UART_CONSUMER

#endif /* EVENT_UART_CONSUMER */