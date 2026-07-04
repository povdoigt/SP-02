#include <string.h>

#include "main.h"

#include "full_seq_utils.h"

#include "event_uart.h"
#include "float3.h"
#include "led_scheduler.h"
#include "waveform.h"
#include "waveform_def.h"




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

static waveform_space_t waveform_wait_button;
waveform_space_t waveform_wait_actuator;
waveform_space_t waveform_error;
waveform_space_t waveform_prgm0_start;
waveform_space_t waveform_wait_jack_ready;
waveform_space_t waveform_wait_jack_launch;
waveform_space_t waveform_wait_sepa;
waveform_space_t waveform_perform_sepa;
waveform_space_t waveform_wait_launch;
waveform_space_t waveform_in_flight;
waveform_space_t waveform_2nd_burn;
waveform_space_t waveform_apogee;

static const waveform_space_mult_const_t waveform_wait_button_ctx = {
	.wave_function_mult = waveform_sine,
	.ctx_mult = &(waveform_sine_t){
		.start = 0.0f,
		.end = 1.0f
	},
	.factor = (float3_t){ .x = 1.0f, .y = 1.0f, .z = 0.0f } // yellow
};

const waveform_space_mult_const_t waveform_wait_actuator_ctx = {
	.wave_function_mult = waveform_scale_add,
	.ctx_mult = &(waveform_scale_add_t){
		.wave_function_add_1 = waveform_gate,
		.ctx_add_1 = &(waveform_gate_t){
			.start = 0.0f,
			.end = 0.1f
		},
		.wave_function_add_2 = waveform_gate,
		.ctx_add_2 = &(waveform_gate_t){
			.start = 0.2f,
			.end = 0.3f
		}
	},
	.factor = (float3_t){ .x = 1.0f, .y = 1.0f, .z = 0.0f } // yellow
};

const waveform_space_mult_const_t waveform_error_ctx = {
	.wave_function_mult = waveform_scale_add,
	.ctx_mult = &(waveform_scale_add_t){
		.wave_function_add_1 = waveform_gate,
		.ctx_add_1 = &(waveform_gate_t){
			.start = 0.0f,
			.end = 0.1f
		},
		.wave_function_add_2 = waveform_gate,
		.ctx_add_2 = &(waveform_gate_t){
			.start = 0.2f,
			.end = 0.3f
		}
	},
	.factor = FLOAT3_UNIT_X // red
};

const waveform_space_add_t waveform_prgm0_start_ctx = {
	.wave_function_add_1 = waveform_space_add,
	.ctx_add_1 = &(waveform_space_add_t){
		.wave_function_add_1 = waveform_space_mult_const,
		.ctx_add_1 = &(waveform_space_mult_const_t){
			.wave_function_mult = waveform_gate,
			.ctx_mult = &(waveform_gate_t){
				.start = 0.00f,
				.end = 1.0f/3.0f
			},
			.factor = FLOAT3_UNIT_X // red
		},
		.wave_function_add_2 = waveform_space_mult_const,
		.ctx_add_2 = &(waveform_space_mult_const_t){
			.wave_function_mult = waveform_gate,
			.ctx_mult = &(waveform_gate_t){
				.start = 1.0f/3.0f,
				.end = 2.0f/3.0f
			},
			.factor = FLOAT3_UNIT_Y // green
		}	
	},
	.wave_function_add_2 = waveform_space_mult_const,
	.ctx_add_2 = &(waveform_space_mult_const_t){
		.wave_function_mult = waveform_gate,
		.ctx_mult = &(waveform_gate_t){
			.start = 2.0f/3.0f,
			.end = 1.0f
		},
		.factor = FLOAT3_UNIT_Z // blue
	}
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

const waveform_space_mult_const_t waveform_perform_sepa_ctx = {
	.wave_function_mult = waveform_scale_add,
	.ctx_mult = &(waveform_scale_add_t){
		.wave_function_add_1 = waveform_gate,
		.ctx_add_1 = &(waveform_gate_t){
			.start = 0.0f,
			.end = 0.1f
		},
		.wave_function_add_2 = waveform_gate,
		.ctx_add_2 = &(waveform_gate_t){
			.start = 0.2f,
			.end = 0.3f
		}
	},
	.factor = FLOAT3_UNIT_Z // blue
};

const waveform_space_mult_const_t waveform_wait_launch_ctx = {
	.wave_function_mult = waveform_scale_add,
	.ctx_mult = &(waveform_scale_add_t){
		.wave_function_add_1 = waveform_gate,
		.ctx_add_1 = &(waveform_gate_t){
			.start = 0.0f,
			.end = 0.1f
		},
		.wave_function_add_2 = waveform_gate,
		.ctx_add_2 = &(waveform_gate_t){
			.start = 0.2f,
			.end = 0.3f
		}
	},
	.factor = FLOAT3_UNIT_X // red
};

const waveform_space_mult_const_t waveform_in_flight_ctx = {
	.wave_function_mult = waveform_sine,
	.ctx_mult = &(waveform_sine_t){
		.start = 0.0f,
		.end = 0.5f
	},
	.factor = FLOAT3_UNIT_X // red
};

const waveform_space_mult_const_t waveform_2nd_burn_ctx = {
	.wave_function_mult = waveform_scale_add,
	.ctx_mult = &(waveform_scale_add_t){
		.wave_function_add_1 = waveform_gate,
		.ctx_add_1 = &(waveform_gate_t){
			.start = 0.0f,
			.end = 0.1f
		},
		.wave_function_add_2 = waveform_gate,
		.ctx_add_2 = &(waveform_gate_t){
			.start = 0.2f,
			.end = 0.3f
		}
	},
	.factor = (float3_t){ .x = 1.0f, .y = 1.0f, .z = 1.0f } // white
};

const waveform_space_mult_const_t waveform_apogee_ctx = {
	.wave_function_mult = waveform_scale_add,
	.ctx_mult = &(waveform_scale_add_t){
		.wave_function_add_1 = waveform_scale_add,
		.ctx_add_1 = &(waveform_scale_add_t){
			.wave_function_add_1 = waveform_scale_add,
			.ctx_add_1 = &(waveform_scale_add_t){
				.wave_function_add_1 = waveform_gate,
				.ctx_add_1 = &(waveform_gate_t){
					.start = 0.0f,
					.end = 0.1f
				},
				.wave_function_add_2 = waveform_gate,
				.ctx_add_2 = &(waveform_gate_t){
					.start = 0.2f,
					.end = 0.3f
				}
			},
			.wave_function_add_2 = waveform_gate,
			.ctx_add_2 = &(waveform_gate_t){
				.start = 0.7f,
				.end = 0.8f
			}
		},
		.wave_function_add_2 = waveform_gate,
		.ctx_add_2 = &(waveform_gate_t){
			.start = 0.9f,
			.end = 1.0f
		}
	},
	.factor = (float3_t){ .x = 1.0f, .y = 1.0f, .z = 1.0f } // white
};

/* One-shot confirmation flashes (remplacent les anciens toggles LED1R/G/B) */
waveform_space_t waveform_flash_green;
waveform_space_t waveform_flash_blue;
waveform_space_t waveform_flash_red;

const waveform_space_mult_const_t waveform_flash_green_ctx = {
	.wave_function_mult = waveform_gate,
	.ctx_mult = &(waveform_gate_t){
		.start = 0.0f,
		.end = 1.0f
	},
	.factor = FLOAT3_UNIT_Y // green
};

const waveform_space_mult_const_t waveform_flash_blue_ctx = {
	.wave_function_mult = waveform_gate,
	.ctx_mult = &(waveform_gate_t){
		.start = 0.0f,
		.end = 1.0f
	},
	.factor = FLOAT3_UNIT_Z // blue
};

const waveform_space_mult_const_t waveform_flash_red_ctx = {
	.wave_function_mult = waveform_gate,
	.ctx_mult = &(waveform_gate_t){
		.start = 0.0f,
		.end = 1.0f
	},
	.factor = FLOAT3_UNIT_X // red
};

#define LED_FLASH_DURATION_MS 200

void led_states_init(void) {

	Waveform_Init_Space(&waveform_wait_actuator, waveform_space_mult_const, &waveform_wait_actuator_ctx, 1000, true);	
	Waveform_Init_Space(&waveform_error, waveform_space_mult_const, &waveform_error_ctx, 1000, false);
	Waveform_Init_Space(&waveform_prgm0_start, waveform_space_add, &waveform_prgm0_start_ctx, 1000, false);
	Waveform_Init_Space(&waveform_wait_jack_ready, waveform_space_mult_const, &waveform_wait_jack_ready_ctx, 1000, true);
	Waveform_Init_Space(&waveform_wait_jack_launch, waveform_space_mult_const, &waveform_wait_jack_launch_ctx, 1000, true);
	Waveform_Init_Space(&waveform_wait_sepa, waveform_space_mult_const, &waveform_wait_sepa_ctx, 1000, true);
	Waveform_Init_Space(&waveform_perform_sepa, waveform_space_mult_const, &waveform_perform_sepa_ctx, 1000, true);
	Waveform_Init_Space(&waveform_wait_launch, waveform_space_mult_const, &waveform_wait_launch_ctx, 1000, true);
	Waveform_Init_Space(&waveform_in_flight, waveform_space_mult_const, &waveform_in_flight_ctx, 1000, true);

	Waveform_Init_Space(&waveform_flash_green, waveform_space_mult_const, &waveform_flash_green_ctx, LED_FLASH_DURATION_MS, false);
	Waveform_Init_Space(&waveform_flash_blue, waveform_space_mult_const, &waveform_flash_blue_ctx, LED_FLASH_DURATION_MS, false);
	Waveform_Init_Space(&waveform_flash_red, waveform_space_mult_const, &waveform_flash_red_ctx, LED_FLASH_DURATION_MS, false);
}



/* ===================================================
   WAITING BUTTON STATE
   =================================================== */

waiting_button_t waiting_button;

void waiting_button_init(waiting_button_t *waiting_button, bool inverted_logic) {
	waiting_button->state = WAITING_BUTTON_STATE_WAITING_START;
	waiting_button->inverted_logic = inverted_logic;
	waiting_button->led_evt = LED_SCHED_HANDLE_INVALID;

	Waveform_Init_Space(
		&waveform_wait_button,
		waveform_space_mult_const,
		&waveform_wait_button_ctx,
		1000,
		true
	);
}

bool waiting_button_play(waiting_button_t *waiting_button) {
	bool button_pressed = false;
	
	switch (HAL_GPIO_ReadPin(PRGM_RUN_GPIO_Port, PRGM_RUN_Pin)) {
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
			// set_led_state(waiting_button->rocket_led_evt, &waveform_wait_button, 1);
			waiting_button->led_evt = LedSched_Add(&waveform_wait_button, 1, false, HAL_MAX_DELAY, LED_SCHED_NO_FORCE);
			waiting_button->state = WAITING_BUTTON_STATE_WAITING_PRESS;
			// no break, intentionally fall through to the next case
		}
		case WAITING_BUTTON_STATE_WAITING_PRESS: {
			if (button_pressed) {
				waiting_button->state = WAITING_BUTTON_STATE_WAITING_RELEASE;
			}
			break;
		}
		case WAITING_BUTTON_STATE_WAITING_RELEASE: {
			if (!button_pressed) {
				waiting_button->state = WAITING_BUTTON_STATE_WAITING_START;
				LedSched_Remove(waiting_button->led_evt);
				return true;
			}
			break;
		}
	}
	return false;
}
