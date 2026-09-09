#include "event_uart.h"
#include "usart.h"


event_uart_msg_t event_uart_msg_format(uint32_t timestamp, event_uart_type_t type, event_uart_payload_u payload) {
    return (event_uart_msg_t){
        .header = EVENT_UART_HEADER_MAGIC,
        .timestamp = timestamp,
        .type = type,
        .payload = payload
    };
}

#if EVENT_UART_PRODUCER

event_uart_producer_t event_uart_producer;

HAL_StatusTypeDef event_uart_producer_init(event_uart_producer_t *producer,
								           GPIO_TypeDef *gpio_port, uint16_t gpio_pin,
								           UART_HandleTypeDef *huart) {
    producer->gpio_port = gpio_port;
    producer->gpio_pin = gpio_pin;
    producer->huart = huart;

    UART_buffer_t *buffer_obj = UART_buffer_get(huart);
    if (buffer_obj == NULL) {
        return HAL_ERROR; // Unknown UART instance
    }
    if (buffer_obj->rx_length < sizeof(event_uart_msg_t)) {
        return HAL_ERROR; // Buffer size is too small
    }

    cb_init(&producer->cb, producer->queue, sizeof(event_uart_msg_t),
            EVENT_UART_MAX_COUNT, CB_OVERWRITE_OLDEST);
    
    HAL_GPIO_WritePin(producer->gpio_port, producer->gpio_pin, GPIO_PIN_RESET);
    return HAL_OK;
}

void event_uart_producer_add_event(event_uart_producer_t *producer, event_uart_msg_t event_msg) {
    cb_push(&producer->cb, &event_msg);
}

void event_uart_producer_run(event_uart_producer_t *producer) {
    switch (producer->state) {
        case EVENT_UART_PRODUCER_STATE_WAITING_FOR_EVENT: {
            if (producer->cb.count > 0) {
                producer->state = EVENT_UART_PRODUCER_STATE_SEND_NOTIFICATION;
            }
            break;
        }
        case EVENT_UART_PRODUCER_STATE_SEND_NOTIFICATION: {
            UART_buffer_t *uart_buffer = UART_buffer_get(producer->huart);
            HAL_UARTEx_ReceiveToIdle_IT(producer->huart, uart_buffer->rx_buffer, uart_buffer->rx_length);
            HAL_GPIO_WritePin(producer->gpio_port, producer->gpio_pin, GPIO_PIN_SET);
            producer->state = EVENT_UART_PRODUCER_STATE_WAITING_FOR_READY;
            break;
        }
        case EVENT_UART_PRODUCER_STATE_WAITING_FOR_READY: {
            // Wait for the consumer to be ready to receive
            break;
        }
        case EVENT_UART_PRODUCER_STATE_SEND_MSG: {
            event_uart_msg_t event_msg;
            cb_pop(&producer->cb, &event_msg);
            HAL_UART_Transmit(producer->huart, (uint8_t *)&event_msg, sizeof(event_uart_msg_t), HAL_MAX_DELAY);
            HAL_GPIO_WritePin(producer->gpio_port, producer->gpio_pin, GPIO_PIN_RESET);
            producer->state = EVENT_UART_PRODUCER_STATE_WAITING_FOR_EVENT;
            break;
        }
        default: {
            // WTF ???
            break;
        }
    }
}

void event_uart_producer_callback(uint16_t Size) {
    if (event_uart_producer.state == EVENT_UART_PRODUCER_STATE_WAITING_FOR_READY) {
        if (Size == sizeof(event_uart_msg_t)) {
            UART_buffer_t *uart_buffer = UART_buffer_get(event_uart_producer.huart);
            event_uart_msg_t *received_msg = (event_uart_msg_t *)uart_buffer->rx_buffer;
            if (received_msg->header == EVENT_UART_HEADER_MAGIC) {
                if (received_msg-> type == EVENT_UART_TYPE_READY) {
                    event_uart_producer.state = EVENT_UART_PRODUCER_STATE_SEND_MSG;
                }
            }
        }
    }
}

#endif /* EVENT_UART_PRODUCER */

#if EVENT_UART_CONSUMER

event_uart_consumer_t event_uart_consumer;

HAL_StatusTypeDef event_uart_consumer_init(event_uart_consumer_t *consumer,
										   GPIO_TypeDef	*gpio_port, uint16_t gpio_pin,
										   uart_mux_t *mux, uart_mux_channel_t channel,
										   UART_HandleTypeDef *huart) {
    consumer->gpio_port = gpio_port;
    consumer->gpio_pin = gpio_pin;
    consumer->mux = mux;
    consumer->channel = channel;
    consumer->huart = huart;

    UART_buffer_t *uart_buffer = UART_buffer_get(huart);
    if (uart_buffer == NULL) {
        return HAL_ERROR; // Unknown UART instance
    }
    cb_init(&consumer->cb, consumer->queue, sizeof(event_uart_msg_t),
            EVENT_UART_MAX_COUNT, CB_OVERWRITE_OLDEST);
    
    consumer->state = EVENT_UART_CONSUMER_STATE_WAITING_FOR_NOTFICATION;

    return HAL_OK;
}

void event_uart_consumer_run(event_uart_consumer_t *consumer) {
    switch (consumer->state) {
        case EVENT_UART_CONSUMER_STATE_WAITING_FOR_NOTFICATION: {
            // Wait for notification from producer
            if (HAL_GPIO_ReadPin(consumer->gpio_port, consumer->gpio_pin) == GPIO_PIN_RESET) {
                consumer->state = EVENT_UART_CONSUMER_STATE_SEND_READY;
            }
            break;
        }
        case EVENT_UART_CONSUMER_STATE_SEND_READY: {
            consumer->last_event_time = HAL_GetTick();
            uart_mux_set(consumer->mux, consumer->channel);
            HAL_Delay(10); // Allow time for the UART mux to switch channels
            HAL_UART_Transmit(consumer->huart, (const uint8_t *)&(event_uart_msg_t){
                .header = EVENT_UART_HEADER_MAGIC,
                .timestamp = HAL_GetTick(),
                .type = EVENT_UART_TYPE_READY,
            }, sizeof(event_uart_msg_t), HAL_MAX_DELAY);
            UART_buffer_t *uart_buffer = UART_buffer_get(consumer->huart);
            HAL_UARTEx_ReceiveToIdle_IT(consumer->huart, uart_buffer->rx_buffer, uart_buffer->rx_length);
            consumer->state = EVENT_UART_CONSUMER_STATE_WAITING_FOR_MSG;
            break;
        }
        case EVENT_UART_CONSUMER_STATE_WAITING_FOR_MSG: {
            if (HAL_GetTick() - consumer->last_event_time > EVENT_UART_CONSUMER_TIMEOUT_MS) {
                consumer->state = EVENT_UART_CONSUMER_STATE_COMEPLETED;
            }
            break;
        }
        case EVENT_UART_CONSUMER_STATE_COMEPLETED: {
            uart_mux_channel_t other_channel = (consumer->channel == UART_MUX_CHANNEL_0) ? UART_MUX_CHANNEL_1 : UART_MUX_CHANNEL_0;
            uart_mux_set(consumer->mux, other_channel);
            consumer->state = EVENT_UART_CONSUMER_STATE_WAITING_FOR_NOTFICATION;
            break;
        }
        default: {
            // WTF ???
            break;
        }
    }
}

void event_uart_consumer_callback(uint16_t Size) {
    if (Size == sizeof(event_uart_msg_t)) {
        UART_buffer_t *uart_buffer = UART_buffer_get(event_uart_consumer.huart);
        event_uart_msg_t *received_msg = (event_uart_msg_t *)uart_buffer->rx_buffer;
        if (received_msg->header == EVENT_UART_HEADER_MAGIC) {
            cb_push(&event_uart_consumer.cb, received_msg);
        }
    }
}

#endif /* EVENT_UART_CONSUMER */