#include "full_seq_utils.h"
#include "event_uart.h"
#include "stm32f0xx_hal.h"




/* ===================================================
   PHASE TRANSITION
   =================================================== */

void phase_transition_init(stage_phase_transition_t *transition, stage_phase_type_t stage_phase_type, uint8_t *phase_variable) {
	transition->stage_phase_type = stage_phase_type;
	transition->phase_variable = phase_variable;
}

void change_state_and_notify(stage_phase_transition_t *transition, uint8_t new_state) {
	if (transition->phase_variable != NULL) {
		event_uart_producer_add_event(&event_uart_producer, event_uart_msg_format(
			HAL_GetTick(),
			EVENT_UART_TYPE_STATE_MACHINE_STATE_CHANGE,
			(event_uart_payload_u){.state_change_payload = {
				.state_machine_id = transition->stage_phase_type,
				.old_state_id = *(transition->phase_variable),
				.new_state_id = new_state
			}}
		));
		*(transition->phase_variable) = new_state;
	}
}




/* ===================================================
   ROCKET
   =================================================== */

void init_rocket_state(rocket_state_t *rocket_state, led_rgb_t *led_rgb) {
	rocket_state->stage = ROCKET_STAGE_NO_SET;
	rocket_state->dynamics = (rocket_dynamics_t){ 0 };
	rocket_state->stage_phase_transition = (stage_phase_transition_t){ 0 };
	rocket_state->led_rgb = led_rgb;
	rocket_state->t_launch = 0;
	rocket_state->is_launch_confirmed = false;
	rocket_state->is_separation_confirmed = false;
	rocket_state->is_second_burn_confirmed = false;
}




/* ===================================================
   WAITING BUTTON STATE
   =================================================== */

void waiting_button_init(waiting_button_t *waiting_button, GPIO_TypeDef *button_gpio_port, uint16_t button_gpio_pin, led_rgb_t *led_rgb, bool inverted_logic) {
	waiting_button->state = WAITING_BUTTON_STATE_WAITING_PRESS;
	waiting_button->led_rgb = led_rgb;
	waiting_button->button_gpio_port = button_gpio_port;
	waiting_button->button_gpio_pin = button_gpio_pin;
	waiting_button->inverted_logic = inverted_logic;
}

bool waiting_button_play(waiting_button_t *waiting_button) {
	bool button_pressed = false;

	if (HAL_GetTick() - waiting_button->t_led >= 500) {
		waiting_button->t_led = HAL_GetTick();
		if (waiting_button->led_on) {
			LED_RGB_SetColor(waiting_button->led_rgb, FLOAT3_ZERO);
		} else {
			LED_RGB_SetColor(waiting_button->led_rgb, (float3_t){ .x = 1.0f, .y = 1.0f, .z = 0.0f });
		}
		waiting_button->led_on = !waiting_button->led_on;
	}
	
	switch (HAL_GPIO_ReadPin(waiting_button->button_gpio_port, waiting_button->button_gpio_pin)) {
		case GPIO_PIN_RESET: {
			button_pressed = waiting_button->inverted_logic ? true : false;
			break;
		}
		case GPIO_PIN_SET: {
			button_pressed = waiting_button->inverted_logic ? false : true;
			break;
		}
	}
	
	switch (waiting_button->state) {
		case WAITING_BUTTON_STATE_WAITING_PRESS: {
			if (button_pressed) {
				waiting_button->state = WAITING_BUTTON_STATE_WAITING_RELEASE;
				waiting_button->led_on = false;
			}
			break;
		}
		case WAITING_BUTTON_STATE_WAITING_RELEASE: {
			if (!button_pressed) {
				waiting_button->state = WAITING_BUTTON_STATE_WAITING_PRESS;
				LED_RGB_SetColor(waiting_button->led_rgb, FLOAT3_ZERO);
				return true;
			}
			break;
		}
	}
	return false;
}