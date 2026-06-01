#include "project.h"

#include "main.h"
#include "stm32f0xx_hal.h"

#include "STS.h"
#include "WT901B.h"

#include "quaternion_dynamics.h"
#include "data_topic.h"
#include "event_uart.h"
#include "stm32f0xx_hal_gpio.h"
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

static rocket_stage_t current_stage;

static second_stage_initialisation_phase_t second_stage_init_phase;
static second_stage_flight_phase_t second_stage_flight_phase;

static bool is_launch_confirmed;
static bool is_separation_confirmed;
static bool is_second_burn_confirmed;

static uint32_t t0;

static uint32_t t_launch; // Time refered to the launch, in seconds since boot
static uint32_t t_2nd_burn; // Time referef to the 2nd burn

static float current_accel_g; // Latest acceleration norm in g, used for state machine decisions
static float current_pressure_variation_pa_s; // Latest pressure variation in Pa/s, used for state machine decisions

static rocket_attitude_dynamics_t attitude;


static stage_phase_transition_t current_stage_phase_transition;


void setup() {

	// Looking for which stage we are
	if (HAL_GPIO_ReadPin(ETAGE2_GPIO_Port, ETAGE2_Pin) == GPIO_PIN_RESET) {
		current_stage = ROCKET_SECOND_STAGE;
		HAL_GPIO_TogglePin(LED1G_GPIO_Port, LED1G_Pin);
		HAL_Delay(500);
		HAL_GPIO_TogglePin(LED1G_GPIO_Port, LED1G_Pin);

		HAL_Delay(1000);

		// setup_servomotors_stage_2();
		setup_data_acquisition_stage_2();
		setup_attitude();
		
		// second_stage_flight_phase = SECOND_STAGE_FLIGHT_INITIALISATION;
		second_stage_flight_phase = SECOND_STAGE_FLIGHT_INITIALISATION;
		second_stage_init_phase = SECOND_STAGE_INIT_WAIT_STAGE_ASSEMBLY_CONFIRMATION;
	} else {
		Error_Handler();
	}

	// setup_event_uart();

	// Set up initial confirmations
	is_launch_confirmed = false;
	is_separation_confirmed = false;
	is_second_burn_confirmed = false;

}

void loop() {
	
	on_new_accel_frame(&accel_sub);
	on_new_gyro_frame(&gyro_sub);
	on_new_pressure_frame(&pressure_sub);

	second_stage_flight_state_machine();

	// event_uart_producer_send_events(&event_uart_producer);
}



/* ===================================================
   SETUP FUNCTIONS
   =================================================== */

void setup_servomotors_stage_2(void) {
	HAL_StatusTypeDef res;

	res = STS_UART_Port_Init(&huart_sts_port2, &huart3);
	if (res != HAL_OK) { Error_Handler(); }

	res = STS_Servo_Init(&servo4, &huart_sts_port2, 4);
	if (res != HAL_OK) { Error_Handler(); }
	HAL_Delay(1);
}

void setup_data_acquisition_stage_2(void) {
	WT901B_status_t wt_res = WT901B_Init(&wt901b, &huart1);
	if (wt_res != WT901B_OK) { Error_Handler(); }

	data_sub_attach(&accel_sub, &wt901b.data_topic, DATA_ATTACH_FROM_OLDEST);
	data_sub_attach(&gyro_sub, &wt901b.data_topic, DATA_ATTACH_FROM_OLDEST);
	data_sub_attach(&pressure_sub, &wt901b.data_topic, DATA_ATTACH_FROM_OLDEST);
}

void setup_attitude(void) {
	attitude.v_up_body = FLOAT3_UNIT_Y; /* axe “up” de la fusée aligné avec l’axe Y du body */
	attitude.initial_elevation_deg = 0.0f * DEG_TO_RAD; /* fusée initialement à 80° */
	attitude.q = quatf_from_axis_angle(FLOAT3_UNIT_Z, - M_PI / 2); // rotation initiale pour aligner l’axe forward du repère Terre avec l’axe Y du body
	attitude.q = quatf_mul(attitude.q, quatf_from_axis_angle(FLOAT3_UNIT_X, attitude.initial_elevation_deg)); // rotation initiale d’élévation
	attitude.elevation_deg = 0.0f;
	attitude.azimuth_deg = 0.0f;
}

void setup_event_uart(void) {
	// event_uart_producer_init(&event_uart_producer, TX_OPTO_N1_GPIO_Port, TX_OPTO_N1_Pin, &huart4);
}






/* ===================================================
   SENSOR DATA CALLBACKS
   =================================================== */

void on_new_accel_frame(data_sub_t *sub) {
	if (data_sub_num_to_read(sub) > 0) {
		WT901B_Frame_t frame;
		data_sub_read(sub, &frame);
		if (frame.type == WT901B_FRAME_ACCEL) {
			current_accel_g = sqrtf(
				frame.data.accel.ax_g * frame.data.accel.ax_g +
				frame.data.accel.ay_g * frame.data.accel.ay_g +
				frame.data.accel.az_g * frame.data.accel.az_g
			);
		}
	}
}

void on_new_gyro_frame(data_sub_t *sub) {
	if (is_launch_confirmed && data_sub_num_to_read(sub) > 0) {
		WT901B_Frame_t frame;
		data_sub_read(sub, &frame);
		if (frame.type == WT901B_FRAME_GYRO) {
			quatdyn_propagate_ip(
				&attitude.q,
				(float3_t){
					.x = frame.data.gyro.gx_dps * DEG_TO_RAD, /* convert to radians/s */
					.y = frame.data.gyro.gy_dps * DEG_TO_RAD,
					.z = frame.data.gyro.gz_dps * DEG_TO_RAD
				},
				WT901B_PERIOD_S
			);
			float3_t v_body_earth = quatf_rotate_vector(attitude.q, attitude.v_up_body);
			attitude.elevation_deg = 90 - acosf(v_body_earth.z) * RAD_TO_DEG;
			float3_t v_body_earth_proj = float3_sub(v_body_earth, float3_scale(FLOAT3_UNIT_Z, v_body_earth.z));
			attitude.azimuth_deg = atan2f(v_body_earth_proj.y, v_body_earth_proj.x) * RAD_TO_DEG;
		}
	}
}

void on_new_pressure_frame(data_sub_t *sub) {
	if (is_launch_confirmed && data_sub_num_to_read(sub) > 0) {
		WT901B_Frame_t frame;
		data_sub_read(sub, &frame);
		if (frame.type == WT901B_FRAME_PRESSURE) {
			// Assuming we receive pressure in Pa, we can compute the variation if we store the previous pressure
			static float previous_pressure_pa = 0.0f;
			float current_pressure_pa = (float)frame.data.pressure.pressure_pa;
			if (previous_pressure_pa > 0.0f) {
				current_pressure_variation_pa_s = (current_pressure_pa - previous_pressure_pa) / WT901B_PERIOD_S; // variation en Pa/s
			}
			previous_pressure_pa = current_pressure_pa;
		}
	}
}


/* ===================================================
   SECOND STAGE STATE MACHINE
   =================================================== */

void second_stage_init_state_machine(void) {

	current_stage_phase_transition.stage_phase_type = STAGE_PHASE_SECOND_STAGE_INIT;
	current_stage_phase_transition.phase_variable = (uint8_t *)&second_stage_init_phase;

	// TODO: fill this in with actual initialisation steps for the second stage

	
	// SECOND_STAGE_INIT_PARA_ZERO,
	// SECOND_STAGE_INIT_WAIT_PARA_ZERO,
	// SECOND_STAGE_INIT_WAIT_STAGE_ASSEMBLY_CONFIRMATION,
	// SECOND_STAGE_INIT_WAIT_JACK,

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
				second_stage_flight_phase = SECOND_STAGE_FLIGHT_WAIT_LAUNCH_CONFIRMATION;
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

void second_stage_flight_state_machine(void) {
	
	// SECOND_STAGE_FLIGHT_INITIALISATION,
	// SECOND_STAGE_FLIGHT_WAIT_LAUNCH_CONFIRMATION,
	// SECOND_STAGE_FLIGHT_WAIT_SEPARATION_CONFIRMATION,
	// SECOND_STAGE_FLIGHT_IGNITION,
	// SECOND_STAGE_FLIGHT_WAIT_IGNITION_CONFIRMATION,
	// SECOND_STAGE_FLIGHT_WAIT_APOGEE_CONFIRMATION,
	// SECOND_STAGE_FLIGHT_APOGEE,
	// SECOND_STAGE_FLIGHT_WAIT_DROGUE_CONFIRMATION,
	// SECOND_STAGE_FLIGHT_WAIT_LANDING_CONFIRMATION,
	// SECOND_STAGE_FLIGHT_IDLE,

	current_stage_phase_transition.stage_phase_type = STAGE_PHASE_SECOND_STAGE_FLIGHT;
	current_stage_phase_transition.phase_variable = (uint8_t *)&second_stage_flight_phase;

	switch (second_stage_flight_phase) {
		case SECOND_STAGE_FLIGHT_INITIALISATION: {
			second_stage_init_state_machine();
			break;
		}
		case SECOND_STAGE_FLIGHT_WAIT_LAUNCH_CONFIRMATION: {
			// Wait for launch confirmation
			if (HAL_GPIO_ReadPin(IN_TRG_N1_GPIO_Port, IN_TRG_N1_Pin) == GPIO_PIN_RESET) {
				t_launch = HAL_GetTick();

				HAL_GPIO_TogglePin(LED1G_GPIO_Port, LED1G_Pin);
				HAL_Delay(100);
				HAL_GPIO_TogglePin(LED1G_GPIO_Port, LED1G_Pin);
				is_launch_confirmed = true;
				change_state_and_notify(SECOND_STAGE_FLIGHT_WAIT_SEPARATION_CONFIRMATION);
			}
			break;
		}
		case SECOND_STAGE_FLIGHT_WAIT_SEPARATION_CONFIRMATION: {
			switch (window_time_get_state(&window_time_sepa, HAL_GetTick() - t_launch)) {
				case WINDOW_TIME_STATE_WAITING: {
					// Still waiting
					break;
				}
				case WINDOW_TIME_STATE_ACTIVE: {
					if (HAL_GPIO_ReadPin(IN_TRG_N2_GPIO_Port, IN_TRG_N2_Pin) == GPIO_PIN_RESET) {

						HAL_GPIO_TogglePin(LED1B_GPIO_Port, LED1B_Pin);
						HAL_Delay(100);
						HAL_GPIO_TogglePin(LED1B_GPIO_Port, LED1B_Pin);

						is_separation_confirmed = true;
						event_uart_producer_add_event(&event_uart_producer, event_uart_msg_format(
							HAL_GetTick(), EVENT_UART_TYPE_WINDOW_TIME, (event_uart_payload_u){ .window_time_payload = {
								.window_id = window_time_sepa.id,
								.result = EVENT_UART_WINDOW_TIME_RESULT_PASS
							}}
						));
						change_state_and_notify(SECOND_STAGE_FLIGHT_WAIT_ATTITUDE_CONFIRMATION);
					}
					break;
				}
				case WINDOW_TIME_STATE_EXPIRED: {
					is_separation_confirmed = false;
					window_time_beta_apogee = window_time_alpha_beta_apogee_no_sepa;
					event_uart_producer_add_event(&event_uart_producer, event_uart_msg_format(
						HAL_GetTick(), EVENT_UART_TYPE_WINDOW_TIME, (event_uart_payload_u){ .window_time_payload = {
							.window_id = window_time_sepa.id,
							.result = EVENT_UART_WINDOW_TIME_RESULT_TIMEOUT
						}}
					));
					change_state_and_notify(SECOND_STAGE_FLIGHT_WAIT_APOGEE_CONFIRMATION);
					break;
				}
			}
			break;
		}
		case SECOND_STAGE_FLIGHT_WAIT_ATTITUDE_CONFIRMATION: {
			switch (window_time_get_state(&window_time_beta_ignition, HAL_GetTick() - t_launch)) {
				case WINDOW_TIME_STATE_WAITING: {
					// Still waiting
					break;
				}
				case WINDOW_TIME_STATE_ACTIVE: {
					if (fabsf(attitude.elevation_deg - ATTITUDE_ELEVATION_GOAL_DEG	) < 10.0f) {
						if (fabsf(attitude.azimuth_deg - ATTITUDE_AZIMUTH_GOAL_DEG) < 45.0f) {
							// Attitude confirmed = command ignition
							event_uart_producer_add_event(&event_uart_producer, event_uart_msg_format(
								HAL_GetTick(), EVENT_UART_TYPE_WINDOW_TIME, (event_uart_payload_u){ .window_time_payload = {
									.window_id = window_time_beta_ignition.id,
									.result = EVENT_UART_WINDOW_TIME_RESULT_PASS
								}}
							));
							change_state_and_notify(SECOND_STAGE_FLIGHT_IGNITION);
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
					change_state_and_notify(SECOND_STAGE_FLIGHT_WAIT_APOGEE_CONFIRMATION);
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
			change_state_and_notify(SECOND_STAGE_FLIGHT_WAIT_IGNITION_CONFIRMATION);
			break;
		}
		case SECOND_STAGE_FLIGHT_WAIT_IGNITION_CONFIRMATION: {
			switch (window_time_get_state(&window_time_beta_ignition_confirm, HAL_GetTick() - t_launch)) {
				case WINDOW_TIME_STATE_WAITING: {
					// Still waiting
					break;
				}
				case WINDOW_TIME_STATE_ACTIVE: {
					if (current_accel_g	> 1.5f) {
						is_second_burn_confirmed = true;
						t_2nd_burn = HAL_GetTick();
						window_time_beta_apogee = window_time_beta_apogee_sepa_ignition;
						event_uart_producer_add_event(&event_uart_producer, event_uart_msg_format(
							HAL_GetTick(), EVENT_UART_TYPE_WINDOW_TIME, (event_uart_payload_u){ .window_time_payload = {
								.window_id = window_time_beta_ignition_confirm.id,
								.result = EVENT_UART_WINDOW_TIME_RESULT_PASS
							}}
						));
						change_state_and_notify(SECOND_STAGE_FLIGHT_WAIT_APOGEE_CONFIRMATION);
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
					change_state_and_notify(SECOND_STAGE_FLIGHT_WAIT_APOGEE_CONFIRMATION);
				}
			}
		}
		case SECOND_STAGE_FLIGHT_WAIT_APOGEE_CONFIRMATION: {
			switch (window_time_get_state(&window_time_beta_apogee, HAL_GetTick() - t_launch)) {
				case WINDOW_TIME_STATE_WAITING: {
					// Still waiting for separation confirmation
					break;
				}
				case WINDOW_TIME_STATE_ACTIVE: {
					if (current_pressure_variation_pa_s < 5.0f) { // If we are near apogee (low pressure variation), confirm apogee
						// Apogee confirmed
						event_uart_producer_add_event(&event_uart_producer, event_uart_msg_format(
							HAL_GetTick(), EVENT_UART_TYPE_WINDOW_TIME, (event_uart_payload_u){ .window_time_payload = {
								.window_id = window_time_beta_apogee.id,
								.result = EVENT_UART_WINDOW_TIME_RESULT_PASS
							}}
						));
						change_state_and_notify(SECOND_STAGE_FLIGHT_APOGEE);
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
					change_state_and_notify(SECOND_STAGE_FLIGHT_APOGEE);
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
			change_state_and_notify(SECOND_STAGE_FLIGHT_IDLE);
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




/* ===================================================
   UTILS
   =================================================== */


void change_state_and_notify(uint8_t new_state) {
	if (current_stage_phase_transition.phase_variable != NULL) {
		event_uart_producer_add_event(&event_uart_producer, event_uart_msg_format(HAL_GetTick(), EVENT_UART_TYPE_STATE_MACHINE_STATE_CHANGE,
			(event_uart_payload_u){.state_change_payload = {
				.state_machine_id = current_stage_phase_transition.stage_phase_type,
				.old_state_id = *(current_stage_phase_transition.phase_variable),
				.new_state_id = new_state
			}}
		));
		*(current_stage_phase_transition.phase_variable) = new_state;
	}
}
