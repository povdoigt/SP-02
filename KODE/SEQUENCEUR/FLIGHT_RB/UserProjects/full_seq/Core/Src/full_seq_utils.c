#include <string.h>

#include "full_seq_utils.h"
#include "event_uart.h"
#include "float3.h"

#include "waveform.h"
#include "waveform_def.h"
#include "waveform_space_arithmetic.h"




/* ===================================================
   GPIO MAPPING
   =================================================== */

void update_gpio_input_states(GPIO_TypeDef *input_gpio_port[], uint16_t input_gpio_pin[], GPIO_PinState input_gpio_state[], uint8_t max_input_number) {
	for (size_t i = 0; i < max_input_number; i++) {
		if (input_gpio_port[i]) {
			input_gpio_state[i] = HAL_GPIO_ReadPin(input_gpio_port[i], input_gpio_pin[i]);
		}
	}
}




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

void rocket_state_init(rocket_state_t *rocket_state, led_rgb_t *led_rgb) {
	memset(rocket_state, 0, sizeof(rocket_state_t));
	rocket_state->stage = ROCKET_STAGE_NO_SET;
	rocket_state->dynamics = (rocket_dynamics_t){ 0 };
	rocket_state->stage_phase_transition = (stage_phase_transition_t){ 0 };
	rocket_state->led_rgb = led_rgb;
	rocket_state->t_launch = 0;
	rocket_state->is_launch_confirmed = false;
	rocket_state->is_separation_confirmed = false;
	rocket_state->is_second_burn_confirmed = false;
}

void rocket_state_setup_gpio(rocket_state_t *rocket_state, GPIO_TypeDef *input_gpio_port[], uint16_t input_gpio_pin[]) {
	memcpy(rocket_state->input_gpio_ports, input_gpio_port, sizeof(GPIO_TypeDef *) * MAX_INPUT_NUMBER);
	memcpy(rocket_state->input_gpio_pins, input_gpio_pin, sizeof(uint16_t) * MAX_INPUT_NUMBER);
}





/* ===================================================
   LED STATE
   =================================================== */

waveform_space_t waveform_wait_button;
waveform_space_t waveform_wait_jack_ready;
waveform_space_t waveform_wait_jack_launch;
waveform_space_t waveform_wait_sepa;
waveform_space_t waveform_wait_jack_launch_sepa;
waveform_space_t waveform_perform_sepa;

const waveform_space_mult_const_t waveform_wait_button_ctx = {
	.wave_function_mult = waveform_sine,
	.ctx_mult = &(waveform_sine_t){
		.start = 0.0f,
		.end = 1.0f
	},
	.factor = (float3_t){ .x = 1.0f, .y = 1.0f, .z = 0.0f } // yellow
};

const waveform_space_mult_const_t waveform_wait_jack_ready_ctx = {
	.wave_function_mult = waveform_sine,
	.ctx_mult = &(waveform_sine_t){
		.start = 0.0f,
		.end = 1.0f
	},
	.factor = (float3_t){ .x = 1.0f, .y = 0.0f, .z = 1.0f } // purple
};

const waveform_space_mult_const_t waveform_wait_jack_launch_ctx = {
	.wave_function_mult = waveform_sine,
	.ctx_mult = &(waveform_sine_t){
		.start = 0.0f,
		.end = 0.5f
	},
	.factor = FLOAT3_UNIT_Y // green
};

const waveform_space_mult_const_t waveform_wait_sepa_ctx = {
	.wave_function_mult = waveform_sine,
	.ctx_mult = &(waveform_sine_t){
		.start = 0.5f,
		.end = 1.0f
	},
	.factor = FLOAT3_UNIT_Z // blue
};

const waveform_space_add_t waveform_wait_jack_launch_sepa_ctx = {
	.wave_function_add_1 = waveform_space_mult_const,
	.ctx_add_1 = &waveform_wait_jack_launch_ctx,
	.wave_function_add_2 = waveform_space_mult_const,
	.ctx_add_2 = &waveform_wait_sepa_ctx,
};

const waveform_space_mult_const_t waveform_perform_sepa_ctx = {
	.wave_function_mult = waveform_scale_add,
	.ctx_mult = &(waveform_scale_add_t){
		.wave_function_add_1 = waveform_gate,
		.wave_function_add_2 = waveform_gate,
		.ctx_add_1 = &(waveform_gate_t){
			.start = 0.0f,
			.end = 0.1f
		},
		.ctx_add_2 = &(waveform_gate_t){
			.start = 0.2f,
			.end = 0.3f
		}
	},
	.factor = FLOAT3_UNIT_Z // blue
};



/* ===================================================
   WAITING BUTTON STATE
   =================================================== */

void waiting_button_init(waiting_button_t *waiting_button, GPIO_TypeDef *button_gpio_port, uint16_t button_gpio_pin, waveform_space_t **rocket_waveform_space, bool inverted_logic) {
	waiting_button->state = WAITING_BUTTON_STATE_WAITING_START;
	waiting_button->rocket_waveform_space = rocket_waveform_space;
	waiting_button->button_gpio_port = button_gpio_port;
	waiting_button->button_gpio_pin = button_gpio_pin;
	waiting_button->inverted_logic = inverted_logic;

	Waveform_Init_Space(
		&waiting_button->wait_waveform_space,
		waveform_space_mult_const,
		&waveform_wait_button_ctx,
		1000,
		true
	);
}

bool waiting_button_play(waiting_button_t *waiting_button) {
	bool button_pressed = false;
	
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
		case WAITING_BUTTON_STATE_WAITING_START: {
			// Launch the waveform for the LED to indicate that we are waiting for the button press
			*(waiting_button->rocket_waveform_space) = &waiting_button->wait_waveform_space;
			waiting_button->state = WAITING_BUTTON_STATE_WAITING_PRESS;
			// no break, intentionally fall through to the next case
		}
		case WAITING_BUTTON_STATE_WAITING_PRESS: {
			if (button_pressed) {
				waiting_button->state = WAITING_BUTTON_STATE_WAITING_RELEASE;
				waiting_button->led_on = false;
			}
			break;
		}
		case WAITING_BUTTON_STATE_WAITING_RELEASE: {
			if (!button_pressed) {
				waiting_button->state = WAITING_BUTTON_STATE_WAITING_START;
				*(waiting_button->rocket_waveform_space) = NULL; // Remove the waveform from the rocket's LED
				return true;
			}
			break;
		}
	}
	return false;
}