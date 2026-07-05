#include "stage_2.h"

#include "float3.h"
#include "full_seq_utils.h"
#include "led.h"
#include "STS.h"
#include "WT901B.h"

#include "actuator.h"
#include "quaternion_dynamics.h"
#include "data_topic.h"
#include "event_uart.h"
#include "iir_filter.h"
#include "led_scheduler.h"
#include "stm32f0xx_hal.h"
#include "usbd_cdc_if.h"
#include "waveform.h"
#include "window_time.h"

#include <math.h>
#include <stdint.h>




typedef enum GPIO_idx_name_t {
	JACK_LAUNCH	= 0,
	SEPARATION	= 1,
	JACK_READY	= 2,
} GPIO_idx_name_t;


/* ===================================================
   STATIC VARIABLES
   =================================================== */

static const window_time_t window_time_sepa = {
	.id = 0,
	.start_time_ms = T_ALPHA_BETA_1,
	.duration_ms = 5000
};

static const window_time_t window_time_beta_ignition = {
	.id = 1,
	.start_time_ms = T_BETA_0,
	.duration_ms = 5000
};

static const window_time_t window_time_beta_ignition_confirm = {
	.id = 2,
	.start_time_ms = T_BETA_1,
	.duration_ms = 5000
};

static const window_time_t window_time_beta_apogee_sepa_ignition = {
	.id = 3,
	.start_time_ms = T_BETA_2,
	.duration_ms = 3000
};

static const window_time_t window_time_beta_apogee_sepa_no_ignition = {
	.id = 4,
	.start_time_ms = T_BETA_4,
	.duration_ms = 3000
};

static const window_time_t window_time_alpha_beta_apogee_no_sepa = {
	.id = 5,
	.start_time_ms = T_ALPHA_BETA_2,
	.duration_ms = 3000
};

static window_time_t window_time_beta_apogee;

static STS_Servo_t servo4 = { 0 };

static data_sub_t accel_sub = { 0 };
static data_sub_t gyro_sub = { 0 };
static data_sub_t pressure_sub = { 0 };

static second_stage_initialisation_phase_t second_stage_init_phase;
static second_stage_initialisation_phase_t second_stage_init_next_phase;
static second_stage_flight_phase_t second_stage_flight_phase;

static uint32_t t0_init;

/* ============================================================
   Semantic position aliases  (local to this file)
   ============================================================ */
 
/* Parachute hatch stage 2 (servo2) — 3 positions */
typedef enum Hatch2_Position_t {
	HATCH2_POS_CLOSED  = 0,   /* Fully closed                                   */
	HATCH2_POS_PARTIAL = 1,   /* Partially open                                 */
	HATCH2_POS_OPEN    = 2,   /* Fully open                                     */
} Hatch2_Position_t;
 
 
/* ============================================================
   Static configurations  (adapt angles to real geometry)
   ============================================================ */
 
static const Actuator_Config_t config_hatch2 = {
    .homing_speed_rpm       = -2.0f,
    .homing_calibration_idx = HATCH2_POS_CLOSED,
    .num_positions          = 3,
    .positions_deg          = {
        [HATCH2_POS_CLOSED]  =   0.0f,
        [HATCH2_POS_PARTIAL] =  15.0f,
        [HATCH2_POS_OPEN]    =  30.0f,
    },
};
 

/* ============================================================
   Actuator instances
   ============================================================ */
 
static Actuator_t actuator_hatch2;      /* servo2 */



static float b_coef[5]; // Moving average filter coefficients

static float wx_inp_storage[sizeof(b_coef) / sizeof(float) - 1];
static float wy_inp_storage[sizeof(b_coef) / sizeof(float) - 1];
static float wz_inp_storage[sizeof(b_coef) / sizeof(float) - 1];

static iir_filter_t wx_iir_filter;
static iir_filter_t wy_iir_filter;
static iir_filter_t wz_iir_filter;




/* ===================================================
   GROUND FUNCTIONS
   =================================================== */

ground_func_state_t ground_func_1_stage_2(rocket_state_t *rocket_state) {
	return __ground_func_homming(rocket_state, &actuator_hatch2);
}

ground_func_state_t ground_func_2_stage_2(rocket_state_t *rocket_state) {
	static ground_func_play_actuator_direction_t direction = DIR_FORWARD;
	return __ground_func_play_actuator(rocket_state, &actuator_hatch2, &direction);
}

ground_func_state_t ground_func_3_stage_2(rocket_state_t *rocket_state) {
	(void)rocket_state;
	return GROUND_FUNC_STATE_DONE;
}

ground_func_state_t ground_func_4_stage_2(rocket_state_t *rocket_state) {
	(void)rocket_state;
	return GROUND_FUNC_STATE_DONE;
}

ground_func_state_t ground_func_5_stage_2(rocket_state_t *rocket_state) {
	(void)rocket_state;
	return GROUND_FUNC_STATE_DONE;
}

ground_func_state_t ground_func_6_stage_2(rocket_state_t *rocket_state) {
	(void)rocket_state;
	return GROUND_FUNC_STATE_DONE;
}

ground_func_state_t ground_func_7_stage_2(rocket_state_t *rocket_state) {
	(void)rocket_state;
	return GROUND_FUNC_STATE_DONE;
}

ground_func_state_t ground_func_8_stage_2(rocket_state_t *rocket_state) {
	(void)rocket_state;
	return GROUND_FUNC_STATE_DONE;
}

ground_func_state_t ground_func_9_stage_2(rocket_state_t *rocket_state) {
	(void)rocket_state;
	return GROUND_FUNC_STATE_DONE;
}

ground_func_state_t ground_func_10_stage_2(rocket_state_t *rocket_state) {
	(void)rocket_state;
	return GROUND_FUNC_STATE_DONE;
}

ground_func_state_t ground_func_11_stage_2(rocket_state_t *rocket_state) {
	(void)rocket_state;
	return GROUND_FUNC_STATE_DONE;
}

ground_func_state_t ground_func_12_stage_2(rocket_state_t *rocket_state) {
	(void)rocket_state;
	return GROUND_FUNC_STATE_DONE;
}

ground_func_state_t ground_func_13_stage_2(rocket_state_t *rocket_state) {
	(void)rocket_state;
	return GROUND_FUNC_STATE_DONE;
}

ground_func_state_t ground_func_14_stage_2(rocket_state_t *rocket_state) {
	(void)rocket_state;
	return GROUND_FUNC_STATE_DONE;
}

ground_func_state_t ground_func_15_stage_2(rocket_state_t *rocket_state) {
	(void)rocket_state;
	return GROUND_FUNC_STATE_DONE;
}

// FROM LAUNCH !!!!!! use -> get_beta_target_deg_over_time(HAL_GetTick() - rocket_state->t_launch) to get beta rad at time t
static float get_beta_target_deg_over_time(const window_time_t * const window_time_beta_ignition, uint32_t t_ms) {
	
	// Adjust t_ms based on the state of the window_time_beta_ignition
	switch (window_time_get_state(window_time_beta_ignition, t_ms)) {
		case WINDOW_TIME_STATE_WAITING: {
			t_ms = window_time_beta_ignition->start_time_ms;
		}
		case WINDOW_TIME_STATE_ACTIVE: {
			break;
		}
		case WINDOW_TIME_STATE_EXPIRED: {
			t_ms = (window_time_beta_ignition->start_time_ms + window_time_beta_ignition->duration_ms);
			break;
		}
	}
	
	float t_s = (float)t_ms / 1000.0f;

	return (
		ATIITUDE_ELEVATION_OVER_TIME_COEF0 * (t_s * t_s * t_s) +
		ATTITUDE_ELEVATION_OVER_TIME_COEF1 * (t_s * t_s) +
		ATTITUDE_ELEVATION_OVER_TIME_COEF2 * (t_s) +
		ATTITUDE_ELEVATION_OVER_TIME_COEF3
	) * RAD_TO_DEG;
}


void compute_elevation_azimut(rocket_state_t *rocket_state) {
	if (!rocket_state->is_launch_confirmed) {
		rocket_state->dynamics.q = quatf_from_2_vec3(rocket_state->dynamics.accel_g, FLOAT3_UNIT_Z);
		rocket_state->dynamics.q_init = rocket_state->dynamics.q;
	}

	float3_t v_up_body_earth_init = float3_normalized(quatf_rotate_vector(rocket_state->dynamics.q_init, FLOAT3_UNIT_Y));
	float3_t v_up_body_earth = float3_normalized(quatf_rotate_vector(rocket_state->dynamics.q, FLOAT3_UNIT_Y));
	rocket_state->dynamics.elevation_deg = 90 - acosf(v_up_body_earth.z) * RAD_TO_DEG;

	float3_t v_up_body_earth_init_proj = float3_normalized(float3_sub(v_up_body_earth_init, float3_scale(FLOAT3_UNIT_Z, v_up_body_earth_init.z)));
	float3_t v_up_body_earth_proj = float3_normalized(float3_sub(v_up_body_earth, float3_scale(FLOAT3_UNIT_Z, v_up_body_earth.z)));
	rocket_state->dynamics.azimuth_deg = acosf(float3_dot(v_up_body_earth_init_proj, v_up_body_earth_proj)) * RAD_TO_DEG;
}






void setup_stage_2(rocket_state_t *rocket_state) {

	LED_RGB_SetColor(rocket_state->led_rgb, FLOAT3_UNIT_Z); // blue
	HAL_Delay(1500);
	LED_RGB_SetColor(rocket_state->led_rgb, FLOAT3_ZERO);
	HAL_Delay(500);

	GPIO_TypeDef *stage2_input_gpio_port[MAX_INPUT_NUMBER] = {
		[JACK_LAUNCH] = JACK_LAUNCH_GPIO_Port,
		[SEPARATION] = SEPARATION_GPIO_Port,
		[JACK_READY] = JACK_READY_GPIO_Port,
	};

	uint16_t stage2_input_gpio_pin[MAX_INPUT_NUMBER] = {
		[JACK_LAUNCH] = JACK_LAUNCH_Pin,
		[SEPARATION] = SEPARATION_Pin,
		[JACK_READY] = JACK_READY_Pin,
	};

	rocket_state_setup_gpio(rocket_state, stage2_input_gpio_port, stage2_input_gpio_pin);

    setup_servomotors_stage_2(rocket_state);
    setup_data_acquisition_stage_2(rocket_state);
	setup_iir_filters_stage_2();

    second_stage_init_next_phase = SECOND_STAGE_INIT_WAIT_JACK_READY;

    phase_transition_init(&rocket_state->stage_phase_transition, STAGE_PHASE_FLIGHT, (uint8_t*)&second_stage_init_phase);
	change_state_and_notify(&rocket_state->stage_phase_transition, SECOND_STAGE_FLIGHT_INITIALISATION);
	phase_transition_init(&rocket_state->stage_phase_transition, STAGE_PHASE_INIT, (uint8_t*)&second_stage_init_phase);
	second_stage_init_phase = SECOND_STAGE_INIT_WAIT_BUTTON;
}

void loop_stage_2(rocket_state_t *rocket_state) {
    WT901B_Parse_Frames(&wt901b);

    on_new_accel_frame(rocket_state, &accel_sub);
    on_new_gyro_frame(rocket_state, &gyro_sub);
    on_new_pressure_frame(rocket_state, &pressure_sub);

	compute_elevation_azimut(rocket_state);

	second_stage_flight_state_machine(rocket_state);
}




/* ===================================================
   SETUP FUNCTIONS
   =================================================== */

void setup_servomotors_stage_2(rocket_state_t *rocket_state) {
	HAL_StatusTypeDef res;

	res = STS_UART_Port_Init(&huart_sts_port2, &huart3);
	if (res != HAL_OK) { goto error; }

	res = STS_Servo_Init(&servo4, &huart_sts_port2, 4);
	if (res != HAL_OK) { goto error; }

    /* Actuator_Init() writes CW/CCW EPROM limits derived from each config */
    res = Actuator_Init(&actuator_hatch2, &servo4, &config_hatch2);
    if (res != HAL_OK) { goto error; }

	HAL_Delay(100);
    return;

error:
    LED_RGB_SetColor(rocket_state->led_rgb, FLOAT3_UNIT_X); // red
    Error_Handler();
}

void setup_data_acquisition_stage_2(rocket_state_t *rocket_state) {
	WT901B_status_t wt_res = WT901B_Init(&wt901b, &huart1);
	if (wt_res != WT901B_OK) { goto error; }

	WT901B_Write(&wt901b, WT901B_REG_RRATE, WT901B_RRATE_20HZ);
	WT901B_Write(&wt901b, WT901B_REG_RSW, WT901B_RSW_ACC_BIT | WT901B_RSW_GYRO_BIT | WT901B_RSW_PRESSURE_BIT);

	data_status_t data_res = DT_OK;
	data_res = data_sub_attach(&accel_sub, &wt901b.data_topic, DATA_ATTACH_FROM_OLDEST);
	if (data_res != DT_OK) { goto error; }
	data_res = data_sub_attach(&gyro_sub, &wt901b.data_topic, DATA_ATTACH_FROM_OLDEST);
	if (data_res != DT_OK) { goto error; }
	data_res = data_sub_attach(&pressure_sub, &wt901b.data_topic, DATA_ATTACH_FROM_OLDEST);
	if (data_res != DT_OK) { goto error; }

	return;

error:
	LED_RGB_SetColor(rocket_state->led_rgb, FLOAT3_UNIT_X); // red
	Error_Handler();
}

void setup_iir_filters_stage_2() {
	size_t b_order = sizeof(b_coef) / sizeof(float) - 1;
	// size_t a_order = sizeof(a_coef) / sizeof(float) - 1;

	for (int i = 0; i < b_order + 1; i++) {
		b_coef[i] = 1.0f / (b_order + 1); // Moving average filter coefficients
	}

	iir_init(&wx_iir_filter, 0, b_order, NULL, b_coef, wx_inp_storage, NULL);
	iir_init(&wy_iir_filter, 0, b_order, NULL, b_coef, wy_inp_storage, NULL);
	iir_init(&wz_iir_filter, 0, b_order, NULL, b_coef, wz_inp_storage, NULL);
}




/* ===================================================
   SENSOR DATA CALLBACKS
   =================================================== */

void on_new_accel_frame(rocket_state_t *rocket_state, data_sub_t *sub) {
	if (data_sub_num_to_read(sub) > 0) {
		WT901B_Frame_t frame;
		data_sub_read(sub, &frame);
		if (frame.type == WT901B_FRAME_ACCEL) {
			rocket_state->dynamics.accel_g = (float3_t){
				.x = frame.data.accel.ax_g,
				.y = frame.data.accel.ay_g,
				.z = frame.data.accel.az_g
			};
		}
	}
}

void on_new_gyro_frame(rocket_state_t *rocket_state, data_sub_t *sub) {
	if (data_sub_num_to_read(sub) > 0) {
		WT901B_Frame_t frame;
		data_sub_read(sub, &frame);
		if (frame.type == WT901B_FRAME_GYRO) {
			quatdyn_propagate_ip_3(
				&rocket_state->dynamics.q,
				(float3_t){
					.x = iir_process(&wx_iir_filter, frame.data.gyro.gx_dps * DEG_TO_RAD), /* convert to radians/s */
					.y = iir_process(&wy_iir_filter, frame.data.gyro.gy_dps * DEG_TO_RAD),
					.z = iir_process(&wz_iir_filter, frame.data.gyro.gz_dps * DEG_TO_RAD)
				},
				WT901B_PERIOD_S
			);
		}
	}
}

void on_new_pressure_frame(rocket_state_t *rocket_state, data_sub_t *sub) {
	if (data_sub_num_to_read(sub) > 0) {
		WT901B_Frame_t frame;
		data_sub_read(sub, &frame);
		if (frame.type == WT901B_FRAME_PRESSURE) {
			// Assuming we receive pressure in Pa, we can compute the variation if we store the previous pressure
			static float previous_pressure_pa = 0.0f;
			float current_pressure_pa = (float)frame.data.pressure.pressure_pa;
			if (previous_pressure_pa > 0.0f) {
				rocket_state->dynamics.pressure_variation_pa_s = (current_pressure_pa - previous_pressure_pa) / WT901B_PERIOD_S; // variation en Pa/s
			}
			previous_pressure_pa = current_pressure_pa;
		}
	}
}




/* ===================================================
   SECOND STAGE STATE MACHINE
   =================================================== */

void second_stage_init_state_machine(rocket_state_t *rocket_state) {
	switch (second_stage_init_phase) {
		case SECOND_STAGE_INIT_WAIT_JACK_READY: {
			static led_evt_handle_t led_evt_handle = LED_SCHED_HANDLE_INVALID;
			if (rocket_state->input_gpio_states[JACK_READY] == GPIO_PIN_SET) {
				LedSched_Remove(led_evt_handle);
				second_stage_init_next_phase = SECOND_STAGE_INIT_PARA_ZERO;
				second_stage_init_phase = SECOND_STAGE_INIT_WAIT_BUTTON;
			} else {
				if (!LedSched_IsHandleValid(led_evt_handle)) {
					led_evt_handle = LedSched_Add(&waveform_wait_jack_ready, 0, false, 0, LED_SCHED_NO_FORCE);
				}
			}
			break;
		}

		case SECOND_STAGE_INIT_PARA_ZERO: {
			static led_evt_handle_t led_evt_handle = LED_SCHED_HANDLE_INVALID;
			static bool homing_started = false;
			if (!homing_started) {
				Actuator_HomingStart(&actuator_hatch2);
				led_evt_handle = LedSched_Add(&waveform_wait_actuator, 0, false, 0, LED_SCHED_NO_FORCE);
				homing_started = true;
			}
			switch (Actuator_HomingProcess(&actuator_hatch2)) {
				case ACTUATOR_HOMING_IDLE: {
					Actuator_HomingStart(&actuator_hatch2);
					led_evt_handle = LedSched_Add(&waveform_wait_actuator, 0, false, 0, LED_SCHED_NO_FORCE);
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
					Actuator_GoToPosition(&actuator_hatch2, HATCH2_POS_CLOSED);
					LedSched_Remove(led_evt_handle);

					second_stage_init_next_phase = SECOND_STAGE_INIT_WAIT_ALL_GOOD;
					second_stage_init_phase = SECOND_STAGE_INIT_WAIT_BUTTON;
					break;
				}
			}
			break;
		}

		case SECOND_STAGE_INIT_WAIT_ALL_GOOD: {

			static led_evt_handle_t jack_launch_led_evt_handle = LED_SCHED_HANDLE_INVALID;
			static led_evt_handle_t wait_sepa_led_evt_handle = LED_SCHED_HANDLE_INVALID;
			static led_evt_handle_t jack_ready_led_evt_handle = LED_SCHED_HANDLE_INVALID;

			bool is_waiting_jack_launch = (rocket_state->input_gpio_states[JACK_LAUNCH] == GPIO_PIN_RESET);
			bool is_waiting_sepa = (rocket_state->input_gpio_states[SEPARATION] == GPIO_PIN_RESET);
	
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

			if (!(is_waiting_jack_launch || is_waiting_sepa)) {
				if (rocket_state->input_gpio_states[JACK_READY] == GPIO_PIN_RESET) {
					t0_init = HAL_GetTick();
					LedSched_Remove(jack_ready_led_evt_handle);
					change_state_and_notify(&rocket_state->stage_phase_transition, SECOND_STAGE_INIT_WAIT_ALL_GOOD_STABLE);
				} else {
					if (!LedSched_IsHandleValid(jack_ready_led_evt_handle)) {
						jack_ready_led_evt_handle = LedSched_Add(&waveform_wait_jack_ready, 0, false, 0, LED_SCHED_NO_FORCE);
					}
				}
			} else {
				LedSched_Remove(jack_ready_led_evt_handle);
			}
			break;
		}
		case SECOND_STAGE_INIT_WAIT_ALL_GOOD_STABLE: {
			if (rocket_state->input_gpio_states[JACK_LAUNCH] != GPIO_PIN_SET ||
				rocket_state->input_gpio_states[SEPARATION] != GPIO_PIN_SET) {
				change_state_and_notify(&rocket_state->stage_phase_transition, SECOND_STAGE_INIT_WAIT_ALL_GOOD);
			} else if (HAL_GetTick() - t0_init > 5000) {
				phase_transition_init(&rocket_state->stage_phase_transition, STAGE_PHASE_FLIGHT, (uint8_t *)&second_stage_flight_phase);
				change_state_and_notify(&rocket_state->stage_phase_transition, SECOND_STAGE_FLIGHT_WAIT_LAUNCH_CONFIRMATION);
			}
			LedSched_Clear();
			break;
		}

		case SECOND_STAGE_INIT_WAIT_BUTTON: {
			if (waiting_button_play(&waiting_button, true)) {
				change_state_and_notify(&rocket_state->stage_phase_transition, second_stage_init_next_phase);
			}
			break;
		}
	}
}

void second_stage_flight_state_machine(rocket_state_t *rocket_state) {

	switch (second_stage_flight_phase) {
		case SECOND_STAGE_FLIGHT_INITIALISATION: {
			second_stage_init_state_machine(rocket_state);
			break;
		}
		case SECOND_STAGE_FLIGHT_WAIT_LAUNCH_CONFIRMATION: {
			static led_evt_handle_t led_evt_handle = LED_SCHED_HANDLE_INVALID;
			if (!LedSched_IsHandleValid(led_evt_handle)) {
				led_evt_handle = LedSched_Add(&waveform_wait_launch, 0, false, 0, LED_SCHED_NO_FORCE);
			}
			// Disarm the system if the separation button is pressed during the launch wait phase
			// Go back to the initialisation phase
			if (waiting_button_play(&waiting_button, false) && get_prgm() == SEPARATION_GROUND_FUNC_ID) {
				t0_init = HAL_GetTick();
				LedSched_Remove(led_evt_handle);
				change_state_and_notify(&rocket_state->stage_phase_transition, SECOND_STAGE_FLIGHT_INITIALISATION);
				phase_transition_init(&rocket_state->stage_phase_transition, STAGE_PHASE_INIT, (uint8_t *)&second_stage_init_phase);
				change_state_and_notify(&rocket_state->stage_phase_transition, SECOND_STAGE_INIT_WAIT_ALL_GOOD);
			}
			// Wait for launch confirmation
			if (rocket_state->input_gpio_states[JACK_LAUNCH] == GPIO_PIN_RESET) {
				rocket_state->t_launch = HAL_GetTick();
				rocket_state->is_launch_confirmed = true;
				LedSched_Remove(led_evt_handle);
				LedSched_Add(&waveform_in_flight, 0, false, 0, LED_SCHED_NO_FORCE);
				LedSched_Add(&waveform_flash_green, 1, false, 0, LED_SCHED_HARD_FORCE);
				change_state_and_notify(&rocket_state->stage_phase_transition, SECOND_STAGE_FLIGHT_WAIT_SEPARATION_CONFIRMATION);
			}
			break;
		}
		case SECOND_STAGE_FLIGHT_WAIT_SEPARATION_CONFIRMATION: {
			switch (window_time_get_state(&window_time_sepa, HAL_GetTick() - rocket_state->t_launch)) {
				case WINDOW_TIME_STATE_WAITING: {
					// Still waiting
					break;
				}
				case WINDOW_TIME_STATE_ACTIVE: {
					if (rocket_state->input_gpio_states[SEPARATION] == GPIO_PIN_RESET) {
						LedSched_Add(&waveform_flash_blue, 1, false, 0, LED_SCHED_HARD_FORCE);
						rocket_state->is_separation_confirmed = true;

						event_uart_producer_add_event(&event_uart_producer, event_uart_msg_format(
							HAL_GetTick(), EVENT_UART_TYPE_WINDOW_TIME, (event_uart_payload_u){ .window_time_payload = {
								.window_id = window_time_sepa.id,
								.result = EVENT_UART_WINDOW_TIME_RESULT_PASS
							}}
						));
						change_state_and_notify(&rocket_state->stage_phase_transition, SECOND_STAGE_FLIGHT_WAIT_ATTITUDE_CONFIRMATION);
					}
					break;
				}
				case WINDOW_TIME_STATE_EXPIRED: {
					rocket_state->is_separation_confirmed = false;
					window_time_beta_apogee = window_time_alpha_beta_apogee_no_sepa;
					event_uart_producer_add_event(&event_uart_producer, event_uart_msg_format(
						HAL_GetTick(), EVENT_UART_TYPE_WINDOW_TIME, (event_uart_payload_u){ .window_time_payload = {
							.window_id = window_time_sepa.id,
							.result = EVENT_UART_WINDOW_TIME_RESULT_TIMEOUT
						}}
					));
					change_state_and_notify(&rocket_state->stage_phase_transition, SECOND_STAGE_FLIGHT_WAIT_APOGEE_CONFIRMATION);
					break;
				}
			}
			break;
		}
		case SECOND_STAGE_FLIGHT_WAIT_ATTITUDE_CONFIRMATION: {
			switch (window_time_get_state(&window_time_beta_ignition, HAL_GetTick() - rocket_state->t_launch)) {
				case WINDOW_TIME_STATE_WAITING: {
					// Still waiting
					break;
				}
				case WINDOW_TIME_STATE_ACTIVE: {
					float att_elevation_target_deg = get_beta_target_deg_over_time(&window_time_beta_ignition, HAL_GetTick() - rocket_state->t_launch);
					uint8_t buf[64];
					snprintf((char*)buf, sizeof(buf), "Elv=%-3.2f Trg=%-3.2f Azi=%-3.2f\r\n",
						rocket_state->dynamics.elevation_deg, att_elevation_target_deg, rocket_state->dynamics.azimuth_deg);
					CDC_Transmit_FS(buf, strlen((char*)buf));
					if (fabsf(rocket_state->dynamics.elevation_deg - att_elevation_target_deg) < 10.0f) {
						if (fabsf(rocket_state->dynamics.azimuth_deg - ATTITUDE_AZIMUTH_GOAL_DEG) < 45.0f) {
							// Attitude confirmed = command ignition
							event_uart_producer_add_event(&event_uart_producer, event_uart_msg_format(
								HAL_GetTick(), EVENT_UART_TYPE_WINDOW_TIME, (event_uart_payload_u){ .window_time_payload = {
									.window_id = window_time_beta_ignition.id,
									.result = EVENT_UART_WINDOW_TIME_RESULT_PASS
								}}
							));
							change_state_and_notify(&rocket_state->stage_phase_transition, SECOND_STAGE_FLIGHT_IGNITION);
						}
					}
					break;
				}
				case WINDOW_TIME_STATE_EXPIRED: {
					// Attitude not confirmed = no burn... go back to home
					window_time_beta_apogee = window_time_beta_apogee_sepa_no_ignition;
					event_uart_producer_add_event(&event_uart_producer, event_uart_msg_format(
						HAL_GetTick(), EVENT_UART_TYPE_WINDOW_TIME, (event_uart_payload_u){ .window_time_payload = {
							.window_id = window_time_beta_ignition.id,
							.result = EVENT_UART_WINDOW_TIME_RESULT_TIMEOUT
						}}
					));
					change_state_and_notify(&rocket_state->stage_phase_transition, SECOND_STAGE_FLIGHT_WAIT_APOGEE_CONFIRMATION);
					break;
				}
			}
			break;
		}
		case SECOND_STAGE_FLIGHT_IGNITION: {
			// Command ignition (e.g., by sending a signal to the engine)
			HAL_GPIO_WritePin(STAGE2_IGNITION_GPIO_Port, STAGE2_IGNITION_Pin, GPIO_PIN_SET);

			LedSched_Add(&waveform_flash_green, 1, false, 0, LED_SCHED_HARD_FORCE);
			change_state_and_notify(&rocket_state->stage_phase_transition, SECOND_STAGE_FLIGHT_WAIT_IGNITION_CONFIRMATION);
			break;
		}
		case SECOND_STAGE_FLIGHT_WAIT_IGNITION_CONFIRMATION: {
			switch (window_time_get_state(&window_time_beta_ignition_confirm, HAL_GetTick() - rocket_state->t_launch)) {
				case WINDOW_TIME_STATE_WAITING: {
					// Still waiting
					break;
				}
				case WINDOW_TIME_STATE_ACTIVE: {
					if (rocket_state->dynamics.accel_g.y > 0.3f) {
						rocket_state->is_second_burn_confirmed = true;
						window_time_beta_apogee = window_time_beta_apogee_sepa_ignition;
						LedSched_Add(&waveform_2nd_burn, 1, false, 0, LED_SCHED_NO_FORCE);
						event_uart_producer_add_event(&event_uart_producer, event_uart_msg_format(
							HAL_GetTick(), EVENT_UART_TYPE_WINDOW_TIME, (event_uart_payload_u){ .window_time_payload = {
								.window_id = window_time_beta_ignition_confirm.id,
								.result = EVENT_UART_WINDOW_TIME_RESULT_PASS
							}}
						));
						change_state_and_notify(&rocket_state->stage_phase_transition, SECOND_STAGE_FLIGHT_WAIT_APOGEE_CONFIRMATION);
					}
					break;
				}
				case WINDOW_TIME_STATE_EXPIRED: {
					window_time_beta_apogee = window_time_beta_apogee_sepa_no_ignition;
					event_uart_producer_add_event(&event_uart_producer, event_uart_msg_format(
						HAL_GetTick(), EVENT_UART_TYPE_WINDOW_TIME, (event_uart_payload_u){ .window_time_payload = {
							.window_id = window_time_beta_ignition_confirm.id,
							.result = EVENT_UART_WINDOW_TIME_RESULT_TIMEOUT
						}}
					));
					change_state_and_notify(&rocket_state->stage_phase_transition, SECOND_STAGE_FLIGHT_WAIT_APOGEE_CONFIRMATION);
				}
			}
		}
		case SECOND_STAGE_FLIGHT_WAIT_APOGEE_CONFIRMATION: {
			switch (window_time_get_state(&window_time_beta_apogee, HAL_GetTick() - rocket_state->t_launch)) {
				case WINDOW_TIME_STATE_WAITING: {
					// Still waiting for separation confirmation
					break;
				}
				case WINDOW_TIME_STATE_ACTIVE: {
					if (rocket_state->dynamics.pressure_variation_pa_s < 5.0f) { // If we are near apogee (low pressure variation), confirm apogee
						// Apogee confirmed
						event_uart_producer_add_event(&event_uart_producer, event_uart_msg_format(
							HAL_GetTick(), EVENT_UART_TYPE_WINDOW_TIME, (event_uart_payload_u){ .window_time_payload = {
								.window_id = window_time_beta_apogee.id,
								.result = EVENT_UART_WINDOW_TIME_RESULT_PASS
							}}
						));
						change_state_and_notify(&rocket_state->stage_phase_transition, SECOND_STAGE_FLIGHT_APOGEE);
					}
					break;
				}
				case WINDOW_TIME_STATE_EXPIRED: {
					// Apogee not confirmed, but we can consider it passed
					event_uart_producer_add_event(&event_uart_producer, event_uart_msg_format(
						HAL_GetTick(), EVENT_UART_TYPE_WINDOW_TIME, (event_uart_payload_u){ .window_time_payload = {
							.window_id = window_time_beta_apogee.id,
							.result = EVENT_UART_WINDOW_TIME_RESULT_TIMEOUT
						}}
					));
					change_state_and_notify(&rocket_state->stage_phase_transition, SECOND_STAGE_FLIGHT_APOGEE);
					break;
				}
			}
			break;
		}
		case SECOND_STAGE_FLIGHT_APOGEE: {
			Actuator_GoToPosition(&actuator_hatch2, HATCH2_POS_OPEN);
			LedSched_Add(&waveform_apogee, 1, false, 0, LED_SCHED_NO_FORCE);
			change_state_and_notify(&rocket_state->stage_phase_transition, SECOND_STAGE_FLIGHT_IDLE);
			break;
		}
		case SECOND_STAGE_FLIGHT_WAIT_DROGUE_CONFIRMATION: {
			break;
		}
		case SECOND_STAGE_FLIGHT_WAIT_LANDING_CONFIRMATION: {
			break;
		}
		case SECOND_STAGE_FLIGHT_IDLE: {
			break;
		}
	}
}
