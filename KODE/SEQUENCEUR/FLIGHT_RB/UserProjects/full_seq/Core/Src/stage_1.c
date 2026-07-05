#include "stage_1.h"

#include "float3.h"
#include "full_seq_utils.h"
#include "led.h"
#include "STS.h"
#include "WT901B.h"

#include "actuator.h"
#include "event_uart.h"
#include "led_scheduler.h"
#include "stm32f0xx_hal.h"
#include "waveform_def.h"
#include "window_time.h"

#include <stdint.h>




typedef enum GPIO_idx_name_t {
	JACK_LAUNCH	= 0,
	SEPARATION	= 1,
	JACK_READY	= 2,
} GPIO_idx_name_t;


/* ===================================================
   STATIC VARIABLES
   =================================================== */

static const window_time_t window_time_alpha_beta_sepa = {
	.id = 0,
	.start_time_ms = T_ALPHA_BETA_1,
	.duration_ms = 1000
};

static const window_time_t window_time_alpha_apogee_sepa = {
	.id = 1,
	.start_time_ms = T_ALPHA_0,
	.duration_ms = 3000
};

static const window_time_t window_time_alpha_beta_apogee_no_sepa = {
	.id = 2,
	.start_time_ms = T_ALPHA_BETA_2,
	.duration_ms = 3000
};

static window_time_t window_time_alpha_apogee;


static STS_Servo_t servo1 = { 0 };
static STS_Servo_t servo2 = { 0 };
static STS_Servo_t servo3 = { 0 };


static first_stage_initialisation_phase_t first_stage_init_phase;
static first_stage_initialisation_phase_t first_stage_init_next_phase;
static first_stage_flight_phase_t first_stage_flight_phase;

static uint32_t t0_init;

/* ============================================================
   Semantic position aliases  (local to this file)
   ============================================================ */
 
/* Aerobrake (servo1) — 3 positions */
typedef enum Aerobrake_Position_t {
	AF_POS_MOUNT  = 0,   /* Hard-stop side, for assembly / disassembly     */
	AF_POS_CLOSED = 1,   /* Nominal closed position during flight          */
	AF_POS_OPEN   = 2,   /* Fully deployed                                 */
} Aerobrake_Position_t;
 
/* Parachute hatch stage 1 (servo2) — 3 positions */
typedef enum Hatch1_Position_t {
	HATCH1_POS_CLOSED  = 0,   /* Fully closed                                   */
	HATCH1_POS_PARTIAL = 1,   /* Partially open                                 */
	HATCH1_POS_OPEN    = 2,   /* Fully open                                     */
} Hatch1_Position_t;
 
/* Separation system (servo3) — 3 positions */
typedef enum Separation_Position_t {
	SEPA_POS_LOCKED     = 0,
	SEPA_POS_UNLOCKED   = 1,
	SEPA_POS_RELEASED   = 2,
} Separation_Position_t;
 
 
/* ============================================================
   Static configurations  (adapt angles to real geometry)
   ============================================================ */
 
static const Actuator_Config_t config_aerobrake = {
    .homing_speed_rpm       = -10.0f,
    .homing_calibration_idx = AF_POS_MOUNT,   /* Hard-stop = 0°; flight range 30° ~ 120° */ 
    .num_positions          = 3,
    .positions_deg          = {
        [AF_POS_MOUNT]  =   0.0f,
        [AF_POS_CLOSED] =  30.0f,
        [AF_POS_OPEN]   = 120.0f,
    },
};
 
static const Actuator_Config_t config_hatch1 = {
    .homing_speed_rpm       = -2.0f,
    .homing_calibration_idx = HATCH1_POS_CLOSED,
    .num_positions          = 3,
    .positions_deg          = {
        [HATCH1_POS_CLOSED]  =   0.0f,
        [HATCH1_POS_PARTIAL] =  15.0f,
        [HATCH1_POS_OPEN]    =  30.0f,
    },
};
 
static const Actuator_Config_t config_separation = {
    .homing_speed_rpm       = 10.0f,   /* CW toward upper hard-stop            */
    .homing_calibration_idx = SEPA_POS_LOCKED,   /* Hard-stop = 0°     */
    .num_positions          = 3,
    .positions_deg          = {
        [SEPA_POS_LOCKED]   =    0.0f,
        [SEPA_POS_UNLOCKED] = -540.0f,
        [SEPA_POS_RELEASED] = -690.0f, // 2 full turns
    },
};
 

/* ============================================================
   Actuator instances
   ============================================================ */
 
static Actuator_t actuator_aerobrake;   /* servo1 */
static Actuator_t actuator_hatch1;      /* servo2 */
static Actuator_t actuator_separation;  /* servo3 */




/* ===================================================
   GROUND FUNCTIONS
   =================================================== */

ground_func_state_t ground_func_1_stage_1(rocket_state_t *rocket_state) {
	return __ground_func_homming(rocket_state, &actuator_aerobrake);
}

ground_func_state_t ground_func_2_stage_1(rocket_state_t *rocket_state) {
	return __ground_func_homming(rocket_state, &actuator_hatch1);
}

ground_func_state_t ground_func_3_stage_1(rocket_state_t *rocket_state) {
	return __ground_func_homming(rocket_state, &actuator_separation);
}

ground_func_state_t ground_func_4_stage_1(rocket_state_t *rocket_state) {
	static ground_func_play_actuator_direction_t direction = DIR_FORWARD;
	return __ground_func_play_actuator(rocket_state, &actuator_aerobrake, &direction);
}

ground_func_state_t ground_func_5_stage_1(rocket_state_t *rocket_state) {
	static ground_func_play_actuator_direction_t direction = DIR_FORWARD;
	return __ground_func_play_actuator(rocket_state, &actuator_hatch1, &direction);
}

ground_func_state_t ground_func_6_stage_1(rocket_state_t *rocket_state) {
	static ground_func_play_actuator_direction_t direction = DIR_FORWARD;
	return __ground_func_play_actuator(rocket_state, &actuator_separation, &direction);
}

ground_func_state_t ground_func_7_stage_1(rocket_state_t *rocket_state) {
	(void)rocket_state;
	return GROUND_FUNC_STATE_DONE;
}

ground_func_state_t ground_func_8_stage_1(rocket_state_t *rocket_state) {
	(void)rocket_state;
	return GROUND_FUNC_STATE_DONE;
}

ground_func_state_t ground_func_9_stage_1(rocket_state_t *rocket_state) {
	(void)rocket_state;
	return GROUND_FUNC_STATE_DONE;
}

ground_func_state_t ground_func_10_stage_1(rocket_state_t *rocket_state) {
	(void)rocket_state;
	return GROUND_FUNC_STATE_DONE;
}

ground_func_state_t ground_func_11_stage_1(rocket_state_t *rocket_state) {
	(void)rocket_state;
	return GROUND_FUNC_STATE_DONE;
}

ground_func_state_t ground_func_12_stage_1(rocket_state_t *rocket_state) {
	(void)rocket_state;
	return GROUND_FUNC_STATE_DONE;
}

ground_func_state_t ground_func_13_stage_1(rocket_state_t *rocket_state) {
	(void)rocket_state;
	return GROUND_FUNC_STATE_DONE;
}

ground_func_state_t ground_func_14_stage_1(rocket_state_t *rocket_state) {
	(void)rocket_state;
	return GROUND_FUNC_STATE_DONE;
}

ground_func_state_t ground_func_15_stage_1(rocket_state_t *rocket_state) {
	(void)rocket_state;
	return GROUND_FUNC_STATE_DONE;
}




void setup_stage_1(rocket_state_t *rocket_state) {
	
	LED_RGB_SetColor(rocket_state->led_rgb, FLOAT3_UNIT_Y); // green
	HAL_Delay(1500);
	LED_RGB_SetColor(rocket_state->led_rgb, FLOAT3_ZERO);
	HAL_Delay(500);

	GPIO_TypeDef *stage1_input_gpio_port[MAX_INPUT_NUMBER] = {
		[JACK_LAUNCH] = JACK_LAUNCH_GPIO_Port,
		[SEPARATION] = SEPARATION_GPIO_Port,
		[JACK_READY] = JACK_READY_GPIO_Port,
	};

	uint16_t stage1_input_gpio_pin[MAX_INPUT_NUMBER] = {
		[JACK_LAUNCH] = JACK_LAUNCH_Pin,
		[SEPARATION] = SEPARATION_Pin,
		[JACK_READY] = JACK_READY_Pin,
	};

	rocket_state_setup_gpio(rocket_state, stage1_input_gpio_port, stage1_input_gpio_pin);

    setup_servomotors_stage_1(rocket_state);

	first_stage_init_next_phase = FIRST_STAGE_INIT_WAIT_JACK_READY;
	first_stage_init_phase = FIRST_STAGE_INIT_IDLE;

    phase_transition_init(&rocket_state->stage_phase_transition, STAGE_PHASE_FLIGHT, (uint8_t*)&first_stage_flight_phase);
	change_state_and_notify(&rocket_state->stage_phase_transition, FIRST_STAGE_FLIGHT_INITIALISATION);
	phase_transition_init(&rocket_state->stage_phase_transition, STAGE_PHASE_INIT, (uint8_t*)&first_stage_init_phase);
	change_state_and_notify(&rocket_state->stage_phase_transition, FIRST_STAGE_INIT_WAIT_BUTTON);
}

void loop_stage_1(rocket_state_t *rocket_state) {
    // WT901B_Parse_Frames(&wt901b);

    // on_new_pressure_frame(rocket_state, &pressure_sub);

	first_stage_flight_state_machine(rocket_state);
}


/* ===================================================
   SETUP FUNCTIONS
   =================================================== */

void setup_servomotors_stage_1(rocket_state_t *rocket_state) {
	HAL_StatusTypeDef res;

	/* STS_UART_Port_Init() make sure that the UART ports are properly initialized */
	res = STS_UART_Port_Init(&huart_sts_port1, &huart2);
	if (res != HAL_OK) { goto error; }
	res = STS_UART_Port_Init(&huart_sts_port2, &huart3);
	if (res != HAL_OK) { goto error; }

	/* STS_Servo_Init() initializes the servo motors */
	res = STS_Servo_Init(&servo1, &huart_sts_port1, 1);
	if (res != HAL_OK) { goto error; }
	res = STS_Servo_Init(&servo2, &huart_sts_port1, 2);
	if (res != HAL_OK) { goto error; }
	res = STS_Servo_Init(&servo3, &huart_sts_port2, 3);
	if (res != HAL_OK) { goto error; }

    /* Actuator_Init() writes CW/CCW EPROM limits derived from each config */
    res = Actuator_Init(&actuator_aerobrake,  &servo1, &config_aerobrake);
    if (res != HAL_OK) { goto error; }
    res = Actuator_Init(&actuator_hatch1,     &servo2, &config_hatch1);
    if (res != HAL_OK) { goto error; }
    res = Actuator_Init(&actuator_separation, &servo3, &config_separation);
    if (res != HAL_OK) { goto error; }

	HAL_Delay(100);
    return;

error:
    LED_RGB_SetColor(rocket_state->led_rgb, FLOAT3_UNIT_X); // red
    Error_Handler();
}




/* ===================================================
   FIRST STAGE STATE MACHINE
   =================================================== */


void first_stage_init_state_machine(rocket_state_t *rocket_state) {
	switch (first_stage_init_phase) {
		case FIRST_STAGE_INIT_IDLE: { break; }
		case FIRST_STAGE_INIT_WAIT_JACK_READY: {
			static led_evt_handle_t led_evt_handle = LED_SCHED_HANDLE_INVALID;
			if (rocket_state->input_gpio_states[JACK_READY] == GPIO_PIN_SET) {
				LedSched_Remove(led_evt_handle);
				first_stage_init_next_phase = FIRST_STAGE_INIT_AF_ZERO;
				first_stage_init_phase = FIRST_STAGE_INIT_WAIT_BUTTON;
			} else {
				if (!LedSched_IsHandleValid(led_evt_handle)) {
					led_evt_handle = LedSched_Add(&waveform_wait_jack_ready, 0, false, 0, LED_SCHED_NO_FORCE);
				}
			}
			break;
		}

		case FIRST_STAGE_INIT_AF_ZERO: {
			static led_evt_handle_t led_evt_handle = LED_SCHED_HANDLE_INVALID;
			static bool homing_started = false;
			if (!homing_started) {
				Actuator_HomingStart(&actuator_aerobrake);
				led_evt_handle = LedSched_Add(&waveform_wait_actuator, 0, false, 0, LED_SCHED_NO_FORCE);
				homing_started = true;
			}
			switch (Actuator_HomingProcess(&actuator_aerobrake)) {
				case ACTUATOR_HOMING_IDLE: {
					Actuator_HomingStart(&actuator_aerobrake);
				}
				case ACTUATOR_HOMING_IN_PROGRESS: {
					break;
				}
				case ACTUATOR_HOMING_ERROR: {
					LED_RGB_SetColor(rocket_state->led_rgb, FLOAT3_UNIT_X); // red
					Error_Handler();
					break;
				}
				case ACTUATOR_HOMING_DONE: {
					Actuator_GoToPosition(&actuator_aerobrake, AF_POS_CLOSED);
					LedSched_Remove(led_evt_handle);

					first_stage_init_next_phase = FIRST_STAGE_INIT_PARA_ZERO;
					change_state_and_notify(&rocket_state->stage_phase_transition, FIRST_STAGE_INIT_WAIT_BUTTON);
					homing_started = false;
					break;
				}
			}
			break;
		}
		case FIRST_STAGE_INIT_PARA_ZERO: {
			static led_evt_handle_t led_evt_handle = LED_SCHED_HANDLE_INVALID;
			static bool homing_started = false;
			if (!homing_started) {
				Actuator_HomingStart(&actuator_hatch1);
				led_evt_handle = LedSched_Add(&waveform_wait_actuator, 0, false, 0, LED_SCHED_NO_FORCE);
				homing_started = true;
			}
			switch (Actuator_HomingProcess(&actuator_hatch1)) {
				case ACTUATOR_HOMING_IDLE: {
					Actuator_HomingStart(&actuator_hatch1);
				}
				case ACTUATOR_HOMING_IN_PROGRESS: {
					break;
				}
				case ACTUATOR_HOMING_ERROR: {
					LED_RGB_SetColor(rocket_state->led_rgb, FLOAT3_UNIT_X); // red
					Error_Handler();
					break;
				}
				case ACTUATOR_HOMING_DONE: {
					Actuator_GoToPosition(&actuator_hatch1, HATCH1_POS_CLOSED);
					LedSched_Remove(led_evt_handle);

					first_stage_init_next_phase = FIRST_STAGE_INIT_SEPA_ZERO;
					change_state_and_notify(&rocket_state->stage_phase_transition, FIRST_STAGE_INIT_WAIT_BUTTON);
					homing_started = false;
					break;
				}
			}
			break;
		}
		case FIRST_STAGE_INIT_SEPA_ZERO: {
			static led_evt_handle_t led_evt_handle = LED_SCHED_HANDLE_INVALID;
			static bool homing_started = false;
			if (!homing_started) {
				Actuator_HomingStart(&actuator_separation);
				led_evt_handle = LedSched_Add(&waveform_wait_actuator, 0, false, 0, LED_SCHED_NO_FORCE);
				homing_started = true;
			}
			switch (Actuator_HomingProcess(&actuator_separation)) {
				case ACTUATOR_HOMING_IDLE: {
					Actuator_HomingStart(&actuator_separation);
				}
				case ACTUATOR_HOMING_IN_PROGRESS: {
					break;
				}
				case ACTUATOR_HOMING_ERROR: {
					LED_RGB_SetColor(rocket_state->led_rgb, FLOAT3_UNIT_X); // red
					Error_Handler();
					break;
				}
				case ACTUATOR_HOMING_DONE: {
					Actuator_GoToPosition(&actuator_separation, SEPA_POS_UNLOCKED);
					LedSched_Remove(led_evt_handle);
					change_state_and_notify(&rocket_state->stage_phase_transition, FIRST_STAGE_INIT_WAIT_ALL_GOOD);
					homing_started = false;
					break;
				}
			}
			break;
		}

		case FIRST_STAGE_INIT_WAIT_ALL_GOOD: {

			static led_evt_handle_t jack_launch_led_evt_handle = LED_SCHED_HANDLE_INVALID;
			static led_evt_handle_t wait_sepa_led_evt_handle = LED_SCHED_HANDLE_INVALID;
			static led_evt_handle_t jack_ready_led_evt_handle = LED_SCHED_HANDLE_INVALID;
			
			static sepa_assembly_phase_t sepa_assembly_phase = SEPA_ASSEMBLY_UNLOCK_STAGE;
			static led_evt_handle_t perform_sepa_evt_handle = LED_SCHED_HANDLE_INVALID;

			bool is_waiting_jack_launch = (rocket_state->input_gpio_states[JACK_LAUNCH] == GPIO_PIN_RESET);
			bool is_waiting_sepa = (
				rocket_state->input_gpio_states[SEPARATION] == GPIO_PIN_RESET
				|| actuator_separation.current_position != SEPA_POS_LOCKED
				|| sepa_assembly_phase != SEPA_ASSEMBLY_LOCKED_AND_STABLE
			);
	
			if (is_waiting_jack_launch) {
				if (!LedSched_IsHandleValid(jack_launch_led_evt_handle)) {
					jack_launch_led_evt_handle = LedSched_Add(&waveform_wait_jack_launch, 0, true, 0, LED_SCHED_NO_FORCE);
					Waveform_Restart((waveform_generic_t*)&waveform_wait_sepa); // resync the waveforms
				}
			} else {
				LedSched_Remove(jack_launch_led_evt_handle);
			}

			if (is_waiting_sepa) {
				if (!LedSched_IsHandleValid(wait_sepa_led_evt_handle)) {
					wait_sepa_led_evt_handle = LedSched_Add(&waveform_wait_sepa, 0, true, 0, LED_SCHED_NO_FORCE);
					Waveform_Restart((waveform_generic_t*)&waveform_wait_jack_launch); // resync the waveforms
				}
			} else {
				LedSched_Remove(wait_sepa_led_evt_handle);
			}

			switch (sepa_assembly_phase) {
				case SEPA_ASSEMBLY_UNLOCK_STAGE: {
					Actuator_GoToPosition(&actuator_separation, SEPA_POS_UNLOCKED);
					sepa_assembly_phase = SEPA_ASSEMBLY_WAIT_STAGE_CONNECTION;
					break;
				}
				case SEPA_ASSEMBLY_WAIT_STAGE_CONNECTION: {
					if (rocket_state->input_gpio_states[SEPARATION] == GPIO_PIN_SET) {
						t0_init = HAL_GetTick();
						sepa_assembly_phase = SEPA_ASSEMBLY_WAIT_STAGE_CONNECTION_STABLE;
						perform_sepa_evt_handle = LedSched_Add(&waveform_perform_sepa, 1, false, 0, LED_SCHED_NO_FORCE);
					}
					break;
				}
				case SEPA_ASSEMBLY_WAIT_STAGE_CONNECTION_STABLE: {
					if (rocket_state->input_gpio_states[SEPARATION] == GPIO_PIN_RESET) {
						LedSched_Remove(perform_sepa_evt_handle);
						sepa_assembly_phase = SEPA_ASSEMBLY_WAIT_STAGE_CONNECTION;
					} else if (HAL_GetTick() - t0_init > 1000) { // Wait 1s to ensure that the stage is properly connected
						Actuator_GoToPosition(&actuator_separation, SEPA_POS_LOCKED);
						t0_init = HAL_GetTick();
						sepa_assembly_phase = SEPA_ASSEMBLY_WAIT_LOCK_STAGE;
					}
					break;
				}
				case SEPA_ASSEMBLY_WAIT_LOCK_STAGE: {
					if (rocket_state->input_gpio_states[SEPARATION] == GPIO_PIN_RESET) {
						LedSched_Remove(perform_sepa_evt_handle);
						sepa_assembly_phase = SEPA_ASSEMBLY_UNLOCK_STAGE;
					} else {
						if (HAL_GetTick() - t0_init > 2000) { // Wait 2s to ensure that the stage is properly locked
							LedSched_Remove(perform_sepa_evt_handle);
							sepa_assembly_phase = SEPA_ASSEMBLY_LOCKED_AND_STABLE;
						}
					}
					break;
				}
				case SEPA_ASSEMBLY_LOCKED_AND_STABLE: {
					if (rocket_state->input_gpio_states[SEPARATION] == GPIO_PIN_RESET) {
						sepa_assembly_phase = SEPA_ASSEMBLY_UNLOCK_STAGE;
					}
				}
			}

			if (!(is_waiting_jack_launch || is_waiting_sepa)) {
				if (rocket_state->input_gpio_states[JACK_READY] == GPIO_PIN_RESET) {
					t0_init = HAL_GetTick();
					LedSched_Remove(jack_ready_led_evt_handle);
					change_state_and_notify(&rocket_state->stage_phase_transition, FIRST_STAGE_INIT_WAIT_ALL_GOOD_STABLE);
					sepa_assembly_phase = SEPA_ASSEMBLY_UNLOCK_STAGE; // Reset sepa_assembly_phase if we're go back in later
				} else {
					if (!LedSched_IsHandleValid(jack_ready_led_evt_handle)) {
						jack_ready_led_evt_handle = LedSched_Add(&waveform_wait_jack_ready, 0, false, 0, LED_SCHED_NO_FORCE);
					}
				}
			} else {
				LedSched_Remove(jack_ready_led_evt_handle);
			}

			if (waiting_button_play(&waiting_button, false) && get_prgm() == SEPARATION_GROUND_FUNC_ID) {
				t0_init = HAL_GetTick();
				Actuator_GoToPosition(&actuator_separation, SEPA_POS_UNLOCKED);
				change_state_and_notify(&rocket_state->stage_phase_transition, FIRST_STAGE_INIT_UNLOCK_STAGE_COMMAND);
				sepa_assembly_phase = SEPA_ASSEMBLY_UNLOCK_STAGE; // Reset sepa_assembly_phase if we're go back in later
			}

			break;
		}
		case FIRST_STAGE_INIT_WAIT_ALL_GOOD_STABLE: {
			if (rocket_state->input_gpio_states[JACK_LAUNCH] != GPIO_PIN_SET ||
				rocket_state->input_gpio_states[SEPARATION] != GPIO_PIN_SET) {
				change_state_and_notify(&rocket_state->stage_phase_transition, FIRST_STAGE_INIT_WAIT_ALL_GOOD);
			} else if (HAL_GetTick() - t0_init > 5000) {
				phase_transition_init(&rocket_state->stage_phase_transition, STAGE_PHASE_FLIGHT, (uint8_t *)&first_stage_flight_phase);
				change_state_and_notify(&rocket_state->stage_phase_transition, FIRST_STAGE_FLIGHT_WAIT_LAUNCH_CONFIRMATION);
			}
			LedSched_Clear();
			break;
		}

		case FIRST_STAGE_INIT_UNLOCK_STAGE_COMMAND: {
			static led_evt_handle_t led_evt_handle = LED_SCHED_HANDLE_INVALID;
			if (!LedSched_IsHandleValid(led_evt_handle)) {
				led_evt_handle = LedSched_Add(&waveform_wait_actuator, 1, false, 0, LED_SCHED_NO_FORCE);
			}
			if (HAL_GetTick() - t0_init > SEPARATION_GROUND_FUNC_DELAY) {
				LedSched_Remove(led_evt_handle);
				change_state_and_notify(&rocket_state->stage_phase_transition, FIRST_STAGE_INIT_WAIT_ALL_GOOD);
			}
			break;
		}

		case FIRST_STAGE_INIT_WAIT_BUTTON: {
			if (waiting_button_play(&waiting_button, true)) {
				change_state_and_notify(&rocket_state->stage_phase_transition, first_stage_init_next_phase);
			}
			break;
		}
	}
}

void first_stage_flight_state_machine(rocket_state_t *rocket_state) {

	switch (first_stage_flight_phase) {
		case FIRST_STAGE_FLIGHT_INITIALISATION: {
			first_stage_init_state_machine(rocket_state);
			break;
		}
		case FIRST_STAGE_FLIGHT_WAIT_LAUNCH_CONFIRMATION: {
			static led_evt_handle_t led_evt_handle = LED_SCHED_HANDLE_INVALID;
			if (!LedSched_IsHandleValid(led_evt_handle)) {
				led_evt_handle = LedSched_Add(&waveform_wait_launch, 0, false, 0, LED_SCHED_NO_FORCE);
			}
			// Disarm the system if the separation button is pressed during the launch wait phase
			// Go back to the initialisation phase
			if (waiting_button_play(&waiting_button, false) && get_prgm() == SEPARATION_GROUND_FUNC_ID) {
				t0_init = HAL_GetTick();
				LedSched_Remove(led_evt_handle);
				Actuator_GoToPosition(&actuator_separation, SEPA_POS_UNLOCKED);
				change_state_and_notify(&rocket_state->stage_phase_transition, FIRST_STAGE_FLIGHT_INITIALISATION);
				phase_transition_init(&rocket_state->stage_phase_transition, STAGE_PHASE_INIT, (uint8_t *)&first_stage_init_phase);
				change_state_and_notify(&rocket_state->stage_phase_transition, FIRST_STAGE_INIT_UNLOCK_STAGE_COMMAND);
			}
			// Wait for launch confirmation
			if (rocket_state->input_gpio_states[JACK_LAUNCH] == GPIO_PIN_RESET) {
				rocket_state->t_launch = HAL_GetTick();
				rocket_state->is_launch_confirmed = true;
				LedSched_Remove(led_evt_handle);
				LedSched_Add(&waveform_in_flight, 0, false, 0, LED_SCHED_NO_FORCE);
				LedSched_Add(&waveform_flash_green, 1, false, 0, LED_SCHED_HARD_FORCE);
				change_state_and_notify(&rocket_state->stage_phase_transition, FIRST_STAGE_FLIGHT_WAIT_BURN_END);
			}
			break;
		}
		case FIRST_STAGE_FLIGHT_WAIT_BURN_END: {
			// Besoin de le garder ???
			if (HAL_GetTick() - rocket_state->t_launch > T_ALPHA_BETA_0) {
				change_state_and_notify(&rocket_state->stage_phase_transition, FIRST_STAGE_FLIGHT_SEPARATION);
			}
			break;
		}
		case FIRST_STAGE_FLIGHT_SEPARATION: {
			if (HAL_GetTick() - rocket_state->t_launch > T_ALPHA_BETA_1) {
				LedSched_Add(&waveform_flash_green, 1, false, 0, LED_SCHED_HARD_FORCE);
				Actuator_GoToPosition(&actuator_aerobrake, AF_POS_OPEN);
				Actuator_GoToPosition(&actuator_separation, SEPA_POS_RELEASED);
				change_state_and_notify(&rocket_state->stage_phase_transition, FIRST_STAGE_FLIGHT_WAIT_SEPARATION_CONFIRMATION);
			}
			break;
		}
		case FIRST_STAGE_FLIGHT_WAIT_SEPARATION_CONFIRMATION: {
			switch (window_time_get_state(&window_time_alpha_beta_sepa, HAL_GetTick() - rocket_state->t_launch)) {
				case WINDOW_TIME_STATE_WAITING: {
					// Still waiting for separation confirmation
					break;
				}
				case WINDOW_TIME_STATE_ACTIVE: {
					if (rocket_state->input_gpio_states[SEPARATION] == GPIO_PIN_RESET) {

						LedSched_Add(&waveform_flash_blue, 1, false, 0, LED_SCHED_HARD_FORCE);

						rocket_state->is_separation_confirmed = true;
						window_time_alpha_apogee = window_time_alpha_apogee_sepa;
						event_uart_producer_add_event(&event_uart_producer, event_uart_msg_format(
							HAL_GetTick(), EVENT_UART_TYPE_WINDOW_TIME, (event_uart_payload_u){ .window_time_payload = {
								.window_id = window_time_alpha_beta_sepa.id,
								.result = EVENT_UART_WINDOW_TIME_RESULT_PASS
							}}
						));
						change_state_and_notify(&rocket_state->stage_phase_transition, FIRST_STAGE_FLIGHT_WAIT_APOGEE_CONFIRMATION);
					}
					break;
				}
				case WINDOW_TIME_STATE_EXPIRED: {
					rocket_state->is_separation_confirmed = false;
					window_time_alpha_apogee = window_time_alpha_beta_apogee_no_sepa;
					event_uart_producer_add_event(&event_uart_producer, event_uart_msg_format(
						HAL_GetTick(), EVENT_UART_TYPE_WINDOW_TIME, (event_uart_payload_u){ .window_time_payload = {
							.window_id = window_time_alpha_beta_sepa.id,
							.result = EVENT_UART_WINDOW_TIME_RESULT_TIMEOUT
						}}
					));
					change_state_and_notify(&rocket_state->stage_phase_transition, FIRST_STAGE_FLIGHT_WAIT_APOGEE_CONFIRMATION);
					break;
				}
			}
			break;
		}
		case FIRST_STAGE_FLIGHT_WAIT_APOGEE_CONFIRMATION: {
			switch (window_time_get_state(&window_time_alpha_apogee, HAL_GetTick() - rocket_state->t_launch)) {
				case WINDOW_TIME_STATE_WAITING: {
					// Still waiting for apogee confirmation
					break;
				}
				case WINDOW_TIME_STATE_ACTIVE: {
					if (rocket_state->dynamics.pressure_variation_pa_s > 0.0f) { // If we are near apogee, variation of pressure should be positive (going down)
						// Apogee confirmed
						event_uart_producer_add_event(&event_uart_producer, event_uart_msg_format(
							HAL_GetTick(), EVENT_UART_TYPE_WINDOW_TIME, (event_uart_payload_u){ .window_time_payload = {
								.window_id = window_time_alpha_apogee.id,
								.result = EVENT_UART_WINDOW_TIME_RESULT_PASS
							}}
						));
						change_state_and_notify(&rocket_state->stage_phase_transition, FIRST_STAGE_FLIGHT_APOGEE);
					}
					break;
				}
				case WINDOW_TIME_STATE_EXPIRED: {
					// Apogee not confirmed, but we can consider it passed
					event_uart_producer_add_event(&event_uart_producer, event_uart_msg_format(
						HAL_GetTick(), EVENT_UART_TYPE_WINDOW_TIME, (event_uart_payload_u){ .window_time_payload = {
							.window_id = window_time_alpha_apogee.id,
							.result = EVENT_UART_WINDOW_TIME_RESULT_TIMEOUT
						}}
					));
					change_state_and_notify(&rocket_state->stage_phase_transition, FIRST_STAGE_FLIGHT_APOGEE);
					break;
				}
			}
			break;
		}
		case FIRST_STAGE_FLIGHT_APOGEE: {
			Actuator_GoToPosition(&actuator_hatch1, HATCH1_POS_OPEN);
			LedSched_Add(&waveform_apogee, 1, false, 0, LED_SCHED_NO_FORCE);
			change_state_and_notify(&rocket_state->stage_phase_transition, FIRST_STAGE_FLIGHT_IDLE);
			break;
		}
		case FIRST_STAGE_FLIGHT_WAIT_DROGUE_CONFIRMATION: {
			break;
		}
		case FIRST_STAGE_FLIGHT_WAIT_LANDING_CONFIRMATION: {
			break;
		}
		case FIRST_STAGE_FLIGHT_IDLE: {
			break;
		}
	}
}
