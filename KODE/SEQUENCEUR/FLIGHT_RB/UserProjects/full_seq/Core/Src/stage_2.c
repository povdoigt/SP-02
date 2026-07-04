#include "stage_2.h"

#include "float3.h"
#include "full_seq_utils.h"
#include "led.h"
#include "STS.h"
#include "WT901B.h"

#include "quaternion_dynamics.h"
#include "data_topic.h"
#include "event_uart.h"
#include "window_time.h"

#include <math.h>





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
	.duration_ms = 2000
};

static const window_time_t window_time_beta_ignition_confirm = {
	.id = 2,
	.start_time_ms = T_BETA_1,
	.duration_ms = 2000
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
static second_stage_flight_phase_t second_stage_flight_phase;




/* ===================================================
   BOARD FUNCTIONS
   =================================================== */

void board_func_1_stage_2(rocket_state_t *rocket_state) {
	(void)rocket_state;
}

void board_func_2_stage_2(rocket_state_t *rocket_state) {
	(void)rocket_state;
}

void board_func_3_stage_2(rocket_state_t *rocket_state) {
	(void)rocket_state;
}

void board_func_4_stage_2(rocket_state_t *rocket_state) {
	(void)rocket_state;
}

void board_func_5_stage_2(rocket_state_t *rocket_state) {
	(void)rocket_state;
}

void board_func_6_stage_2(rocket_state_t *rocket_state) {
	(void)rocket_state;
}

void board_func_7_stage_2(rocket_state_t *rocket_state) {
	(void)rocket_state;
}

void board_func_8_stage_2(rocket_state_t *rocket_state) {
	(void)rocket_state;
}

void board_func_9_stage_2(rocket_state_t *rocket_state) {
	(void)rocket_state;
}

void board_func_10_stage_2(rocket_state_t *rocket_state) {
	(void)rocket_state;
}

void board_func_11_stage_2(rocket_state_t *rocket_state) {
	(void)rocket_state;
}

void board_func_12_stage_2(rocket_state_t *rocket_state) {
	(void)rocket_state;
}

void board_func_13_stage_2(rocket_state_t *rocket_state) {
	(void)rocket_state;
}

void board_func_14_stage_2(rocket_state_t *rocket_state) {
	(void)rocket_state;
}

void board_func_15_stage_2(rocket_state_t *rocket_state) {
	(void)rocket_state;
}







void setup_stage_2(rocket_state_t *rocket_state) {

	LED_RGB_SetColor(rocket_state->led_rgb, (float3_t){ .x = 0.0, .y = 0.0, .z = 1.0});
	HAL_Delay(500);
	LED_RGB_SetColor(rocket_state->led_rgb, FLOAT3_ZERO);
	HAL_Delay(500);

    setup_servomotors_stage_2(rocket_state);
    setup_data_acquisition_stage_2(rocket_state);
    setup_attitude_stage_2(rocket_state);

    second_stage_flight_phase = SECOND_STAGE_FLIGHT_INITIALISATION;
    second_stage_init_phase = SECOND_STAGE_INIT_PARA_ZERO;
    phase_transition_init(&rocket_state->stage_phase_transition, STAGE_PHASE_STAGE_INIT, (uint8_t*)&second_stage_init_phase);
}

void loop_stage_2(rocket_state_t *rocket_state) {
    WT901B_Parse_Frames(&wt901b);

    on_new_accel_frame(rocket_state, &accel_sub);
    on_new_gyro_frame(rocket_state, &gyro_sub);
    on_new_pressure_frame(rocket_state, &pressure_sub);

	second_stage_flight_state_machine(rocket_state);
}




/* ===================================================
   SETUP FUNCTIONS
   =================================================== */

void setup_servomotors_stage_2(rocket_state_t *rocket_state) {
	HAL_StatusTypeDef res;

	res = STS_UART_Port_Init(&huart_sts_port2, &huart3);
	if (res != HAL_OK) { Error_Handler(); }

	res = STS_Servo_Init(&servo4, &huart_sts_port2, 4);
	if (res != HAL_OK) { Error_Handler(); }
	HAL_Delay(1);
}

void setup_data_acquisition_stage_2(rocket_state_t *rocket_state) {
	WT901B_status_t wt_res = WT901B_Init(&wt901b, &huart1);
	if (wt_res != WT901B_OK) { Error_Handler(); }

	data_sub_attach(&accel_sub, &wt901b.data_topic, DATA_ATTACH_FROM_OLDEST);
	data_sub_attach(&gyro_sub, &wt901b.data_topic, DATA_ATTACH_FROM_OLDEST);
	data_sub_attach(&pressure_sub, &wt901b.data_topic, DATA_ATTACH_FROM_OLDEST);
}

void setup_attitude_stage_2(rocket_state_t *rocket_state) {
	rocket_state->dynamics.v_up_body = FLOAT3_UNIT_Y; /* axe “up” de la fusée aligné avec l’axe Y du body */
	rocket_state->dynamics.initial_elevation_deg = 0.0f * DEG_TO_RAD; /* fusée initialement à 80° */
	rocket_state->dynamics.q = quatf_from_axis_angle(FLOAT3_UNIT_Z, - M_PI / 2); // rotation initiale pour aligner l’axe forward du repère Terre avec l’axe Y du body
	rocket_state->dynamics.q = quatf_mul(rocket_state->dynamics.q, quatf_from_axis_angle(FLOAT3_UNIT_X, rocket_state->dynamics.initial_elevation_deg)); // rotation initiale d’élévation
	rocket_state->dynamics.elevation_deg = 0.0f;
	rocket_state->dynamics.azimuth_deg = 0.0f;
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
	if (rocket_state->is_launch_confirmed && data_sub_num_to_read(sub) > 0) {
		WT901B_Frame_t frame;
		data_sub_read(sub, &frame);
		if (frame.type == WT901B_FRAME_GYRO) {
			quatdyn_propagate_ip(
				&rocket_state->dynamics.q,
				(float3_t){
					.x = frame.data.gyro.gx_dps * DEG_TO_RAD, /* convert to radians/s */
					.y = frame.data.gyro.gy_dps * DEG_TO_RAD,
					.z = frame.data.gyro.gz_dps * DEG_TO_RAD
				},
				WT901B_PERIOD_S
			);
			float3_t v_body_earth = quatf_rotate_vector(rocket_state->dynamics.q, rocket_state->dynamics.v_up_body);
			rocket_state->dynamics.elevation_deg = 90 - acosf(v_body_earth.z) * RAD_TO_DEG;
			float3_t v_body_earth_proj = float3_sub(v_body_earth, float3_scale(FLOAT3_UNIT_Z, v_body_earth.z));
			rocket_state->dynamics.azimuth_deg = atan2f(v_body_earth_proj.y, v_body_earth_proj.x) * RAD_TO_DEG;
		}
	}
}

void on_new_pressure_frame(rocket_state_t *rocket_state, data_sub_t *sub) {
	if (rocket_state->is_launch_confirmed && data_sub_num_to_read(sub) > 0) {
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
		case SECOND_STAGE_INIT_PARA_ZERO: {
			STS_Servo_SetOperatingMode(&servo4, STS_OP_MODE_SPEED_CONTROL);
			STS_Servo_SetGoalSpeed(&servo4, STS_GetSpeedInUnits(-10));
			second_stage_init_phase = SECOND_STAGE_INIT_WAIT_PARA_ZERO;
			break;
		}
		case SECOND_STAGE_INIT_WAIT_PARA_ZERO: {
			bool in_overload = false;
			STS_Servo_InOverload(&servo4, &in_overload);
			if (true) {
				STS_Servo_SetGoalSpeed(&servo4, 0);
				STS_Servo_PositionCalibration(&servo4, STS_GetPositionInUnits(0));
				STS_Servo_SetOperatingMode(&servo4, STS_OP_MODE_POSITION_CONTROL);
				STS_Servo_SetGoalPosition(&servo4, STS_GetPositionInUnits(30));
				second_stage_init_phase = SECOND_STAGE_INIT_WAIT_STAGE_ASSEMBLY_CONFIRMATION;
			}
			break;
		}
		case SECOND_STAGE_INIT_WAIT_STAGE_ASSEMBLY_CONFIRMATION: {
			if (HAL_GPIO_ReadPin(IN_TRG_N2_GPIO_Port, IN_TRG_N2_Pin) == GPIO_PIN_SET) {
				HAL_GPIO_WritePin(LED1B_GPIO_Port, LED1B_Pin, GPIO_PIN_RESET);
				HAL_Delay(3000);
				if (HAL_GPIO_ReadPin(IN_TRG_N2_GPIO_Port, IN_TRG_N2_Pin) == GPIO_PIN_SET) {
					HAL_GPIO_WritePin(LED1B_GPIO_Port, LED1B_Pin, GPIO_PIN_SET);
					second_stage_init_phase = SECOND_STAGE_INIT_WAIT_JACK;
				}
			} else {
				HAL_GPIO_TogglePin(LED1B_GPIO_Port, LED1B_Pin);
				HAL_Delay(500);
				HAL_GPIO_TogglePin(LED1B_GPIO_Port, LED1B_Pin);
				HAL_Delay(500);
			}
			break;
		}
		case SECOND_STAGE_INIT_WAIT_JACK: {
			if (HAL_GPIO_ReadPin(IN_TRG_N1_GPIO_Port, IN_TRG_N1_Pin) == GPIO_PIN_SET) {
				HAL_Delay(1000);
				HAL_GPIO_TogglePin(LED1R_GPIO_Port, LED1R_Pin);
				HAL_Delay(500);
				HAL_GPIO_TogglePin(LED1R_GPIO_Port, LED1R_Pin);
				// second_stage_flight_phase = SECOND_STAGE_FLIGHT_WAIT_LAUNCH_CONFIRMATION;
                phase_transition_init(&rocket_state->stage_phase_transition, STAGE_PHASE_STAGE_FLIGHT, (uint8_t *)&second_stage_flight_phase);
                change_state_and_notify(&rocket_state->stage_phase_transition, SECOND_STAGE_FLIGHT_WAIT_LAUNCH_CONFIRMATION);
            } else {
				HAL_GPIO_TogglePin(LED1G_GPIO_Port, LED1G_Pin);
				HAL_Delay(500);
				HAL_GPIO_TogglePin(LED1G_GPIO_Port, LED1G_Pin);
				HAL_Delay(500);
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
			// Wait for launch confirmation
			if (HAL_GPIO_ReadPin(IN_TRG_N1_GPIO_Port, IN_TRG_N1_Pin) == GPIO_PIN_RESET) {
				rocket_state->t_launch = HAL_GetTick();

				HAL_GPIO_TogglePin(LED1G_GPIO_Port, LED1G_Pin);
				HAL_Delay(100);
				HAL_GPIO_TogglePin(LED1G_GPIO_Port, LED1G_Pin);
				rocket_state->is_launch_confirmed = true;
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
					if (HAL_GPIO_ReadPin(IN_TRG_N2_GPIO_Port, IN_TRG_N2_Pin) == GPIO_PIN_RESET) {

						HAL_GPIO_TogglePin(LED1B_GPIO_Port, LED1B_Pin);
						HAL_Delay(100);
						HAL_GPIO_TogglePin(LED1B_GPIO_Port, LED1B_Pin);

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
					if (fabsf(rocket_state->dynamics.elevation_deg - ATTITUDE_ELEVATION_GOAL_DEG	) < 10.0f || true) {
						if (fabsf(rocket_state->dynamics.azimuth_deg - ATTITUDE_AZIMUTH_GOAL_DEG) < 45.0f || true) {
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
			HAL_GPIO_TogglePin(OUT_N2_GPIO_Port, OUT_N2_Pin);

			HAL_GPIO_TogglePin(LED1R_GPIO_Port, LED1R_Pin);
			HAL_Delay(100);
			HAL_GPIO_TogglePin(LED1R_GPIO_Port, LED1R_Pin);
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
					if (float3_norm(rocket_state->dynamics.accel_g)	> 1.5f || true) {
						rocket_state->is_second_burn_confirmed = true;
						window_time_beta_apogee = window_time_beta_apogee_sepa_ignition;
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
			STS_Servo_SetGoalPosition(&servo4, STS_GetPositionInUnits(120));
			HAL_GPIO_TogglePin(LED1R_GPIO_Port, LED1R_Pin);
			HAL_Delay(500);
			HAL_GPIO_TogglePin(LED1R_GPIO_Port, LED1R_Pin);
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
