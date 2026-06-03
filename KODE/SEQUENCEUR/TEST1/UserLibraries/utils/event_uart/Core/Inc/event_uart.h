#ifndef EVENT_UART_H
#define EVENT_UART_H

#include <stdint.h>
#include <stdbool.h>

#include "circular_buffer.h"

#define EVENT_UART_MAX_COUNT 10
#define EVENT_UART_HEADER_MAGIC 0xAA55

#define EVENT_UART_PRODUCER 1
#define EVENT_UART_CONSUMER 0

// ======== Event UART Message Structure ========

/**
 * @brief Event UART message types
 * Define different types of events that can be sent over UART. This can be extended as needed.
 * For example, you might have events for state machine changes, timing windows, etc.
 */
typedef enum event_uart_type_t {
	EVENT_UART_TYPE_READY = 0, /*< Ready event to indicate the system is ready */
	EVENT_UART_TYPE_STATE_MACHINE_STATE_CHANGE, /*< State machine state change event */
	EVENT_UART_TYPE_WINDOW_TIME, /*< Timing window event */
} event_uart_type_t;

// ========= Event UART STATE MACHINE STATE CHANGE PAYLOAD STRUCTURE ========

/**
 * @brief Event UART state change payload structure
 * Define the payload structure for state machine state change events.
 */
typedef struct event_uart_state_change_payload_t {
	uint8_t state_machine_id;
	uint8_t old_state_id;
	uint8_t new_state_id;
} event_uart_state_change_payload_t;

// ========= Event UART TIMING WINDOW PAYLOAD STRUCTURE ========

/**
 * @brief Event UART window time result types
 * Define different results for timing window events.
 */
typedef enum event_uart_window_time_result_t {
	EVENT_UART_WINDOW_TIME_RESULT_PASS,
	EVENT_UART_WINDOW_TIME_RESULT_TIMEOUT,
} event_uart_window_time_result_t;

/**
 * @brief Event UART window time payload structure
 * Define the payload structure for timing window events.
 */
typedef struct event_uart_window_time_payload_t {
	uint8_t 						window_id;			/*< Window ID */
	event_uart_window_time_result_t result;				/*< Result of the timing window (pass/fail) */
} event_uart_window_time_payload_t;

// ========= Event UART Message Structure ========

/**
 * @brief Event UART payload union
 * Define the union for different payload types.
 */
typedef union event_uart_payload_u {
	event_uart_state_change_payload_t	state_change_payload;
	event_uart_window_time_payload_t	window_time_payload;
	// Add more payload types here as needed
} event_uart_payload_u;

/**
 * @brief Event UART message structure
 * Define the structure for the complete UART message.
 */
typedef struct event_uart_msg_t {
	uint16_t				header;		// Magic number to identify the start of a message
	uint32_t				timestamp;	// Timestamp of the event (e.g., in milliseconds)
	event_uart_type_t		type;		// Type of the event
	event_uart_payload_u	payload;	// Payload of the event
} event_uart_msg_t;

event_uart_msg_t event_uart_msg_format(uint32_t timestamp, event_uart_type_t type, event_uart_payload_u payload);

#if EVENT_UART_PRODUCER

// ======== Event producer =======

#include "stm32f0xx_hal.h"

typedef struct event_uart_producer_t {
	GPIO_TypeDef		*gpio_port; 					// GPIO port to notify the event
	uint16_t		 	 gpio_pin;						// GPIO pin to notify the event
	UART_HandleTypeDef	*huart;							// USART instance to send the event message
	circular_buffer_t	 cb;							// Circular buffer to store event messages to send
	event_uart_msg_t	 queue[EVENT_UART_MAX_COUNT];	// Circular buffer for event messages
	bool				 notif_sent;					// Flag to indicate if notification has been sent
	bool				 ready_receive;					// Flag to indicate if receive a ready flag from the consumer
} event_uart_producer_t;

extern event_uart_producer_t event_uart_producer;

void event_uart_producer_init(event_uart_producer_t *producer,
								GPIO_TypeDef *gpio_port, uint16_t gpio_pin,
								UART_HandleTypeDef *huart);
void event_uart_producer_add_event(event_uart_producer_t *producer, event_uart_msg_t event_msg);
void event_uart_producer_send_events(event_uart_producer_t *producer);
void event_uart_callback(void);

#endif /* EVENT_UART_PRODUCER */

#if EVENT_UART_CONSUMER

// ======== Event consumer =======

#include "stm32f4xx_hal.h"

#endif /* EVENT_UART_CONSUMER */

#endif /* EVENT_UART_H */