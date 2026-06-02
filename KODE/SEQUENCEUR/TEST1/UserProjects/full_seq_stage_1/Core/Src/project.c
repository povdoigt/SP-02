#include "project.h"

#include "main.h"

#include "stm32f0xx_hal.h"
#include "usart.h"

#include "STS.h"
// #include "WT901B.h"

// #include "data_topic.h"
#include "event_uart.h"
#include "window_time.h"
#include <stdbool.h>
#include <stdint.h>

/* ===================================================
   STATIC VARIABLES
   =================================================== */

static const window_time_t window_time_sepa = {
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

// static data_sub_t pressure_sub = { 0 };

static rocket_stage_t current_stage;

static first_stage_initialisation_phase_t first_stage_init_phase;
static first_stage_flight_phase_t first_stage_flight_phase;

static bool is_launch_confirmed;
static bool is_separation_confirmed;

static uint32_t t0;

static uint32_t t_launch; // Time refered to the launch, in seconds since boot

static float current_pressure_variation_pa_s; // Latest pressure variation in Pa/s, used for state machine decisions

static stage_phase_transition_t current_stage_phase_transition;


void setup() {

	// Looking for which stage we are
	if (HAL_GPIO_ReadPin(ETAGE1_GPIO_Port, ETAGE1_Pin) == GPIO_PIN_RESET) {
		current_stage = ROCKET_FIRST_STAGE;
		HAL_GPIO_TogglePin(LED1B_GPIO_Port, LED1B_Pin);
		HAL_Delay(500);
		HAL_GPIO_TogglePin(LED1B_GPIO_Port, LED1B_Pin);

		HAL_Delay(1000);

		setup_servomotors_stage_1();
		// setup_data_acquisition_stage_1();

		first_stage_init_phase = FIRST_STAGE_INIT_AF_ZERO;
		first_stage_flight_phase = FIRST_STAGE_FLIGHT_INITIALISATION;
	} else {
		Error_Handler();
	}

	// setup_event_uart();

	// Set up initial confirmations
	is_launch_confirmed = false;
	is_separation_confirmed = false;
}

void loop() {
	// on_new_pressure_frame(&pressure_sub);

	first_stage_flight_state_machine();

	// event_uart_producer_send_events(&event_uart_producer);
}



/* ===================================================
   SETUP FUNCTIONS
   =================================================== */

void setup_servomotors_stage_1(void) {
	HAL_StatusTypeDef res;

	res = STS_UART_Port_Init(&huart_sts_port1, &huart2);
	if (res != HAL_OK) { Error_Handler(); }
	res = STS_UART_Port_Init(&huart_sts_port2, &huart3);
	if (res != HAL_OK) { Error_Handler(); }

	res = STS_Servo_Init(&servo1, &huart_sts_port1, 1);
	if (res != HAL_OK) { Error_Handler(); }
	res = STS_Servo_Init(&servo2, &huart_sts_port1, 2);
	if (res != HAL_OK) { Error_Handler(); }
	res = STS_Servo_Init(&servo3, &huart_sts_port2, 3);
	if (res != HAL_OK) { Error_Handler(); }
}

// void setup_data_acquisition_stage_1(void) {
// 	WT901B_status_t wt_res = WT901B_Init(&wt901b, &huart3);
// 	if (wt_res != WT901B_OK) { Error_Handler(); }

// 	data_sub_attach(&pressure_sub, &wt901b.data_topic, DATA_ATTACH_FROM_OLDEST);
// }

void setup_event_uart(void) {
	event_uart_producer_init(&event_uart_producer, TX_OPTO_N1_GPIO_Port, TX_OPTO_N1_Pin, &huart4);
}






/* ===================================================
   SENSOR DATA CALLBACKS
   =================================================== */

// void on_new_pressure_frame(data_sub_t *sub) {
// 	if (is_launch_confirmed && data_sub_num_to_read(sub) > 0) {
// 		WT901B_Frame_t frame;
// 		data_sub_read(sub, &frame);
// 		if (frame.type == WT901B_FRAME_PRESSURE) {
// 			// Assuming we receive pressure in Pa, we can compute the variation if we store the previous pressure
// 			static float previous_pressure_pa = 0.0f;
// 			float current_pressure_pa = (float)frame.data.pressure.pressure_pa;
// 			if (previous_pressure_pa > 0.0f) {
// 				current_pressure_variation_pa_s = (current_pressure_pa - previous_pressure_pa) / WT901B_PERIOD_S; // variation en Pa/s
// 			}
// 			previous_pressure_pa = current_pressure_pa;
// 		}
// 	}
// }


/* ===================================================
   FIRST STAGE STATE MACHINE
   =================================================== */


void first_stage_init_state_machine(void) {

	current_stage_phase_transition.stage_phase_type = STAGE_PHASE_FIRST_STAGE_INIT;
	current_stage_phase_transition.phase_variable = (uint8_t *)&first_stage_init_phase;

	switch (first_stage_init_phase) {
		case FIRST_STAGE_INIT_AF_ZERO: {
			STS_Servo_SetOperatingMode(&servo1, STS_OP_MODE_SPEED_CONTROL);
			STS_Servo_SetGoalSpeed(&servo1, STS_GetSpeedInUnits(-10));
			first_stage_init_phase = FIRST_STAGE_INIT_WAIT_AF_ZERO;
			break;
		}
		case FIRST_STAGE_INIT_WAIT_AF_ZERO: {
			bool in_overload = false;
			STS_Servo_InOverload(&servo1, &in_overload);
			if (in_overload) {
				STS_Servo_SetGoalSpeed(&servo1, 0);
				STS_Servo_PositionCalibration(&servo1, STS_GetPositionInUnits(0));
				STS_Servo_SetOperatingMode(&servo1, STS_OP_MODE_POSITION_CONTROL);
				STS_Servo_SetGoalPosition(&servo1, STS_GetPositionInUnits(30));
				first_stage_init_phase = FIRST_STAGE_INIT_PARA_ZERO;
			}
			break;
		}
		case FIRST_STAGE_INIT_PARA_ZERO: {
			STS_Servo_SetOperatingMode(&servo2, STS_OP_MODE_SPEED_CONTROL);
			STS_Servo_SetGoalSpeed(&servo2, STS_GetSpeedInUnits(-10));
			first_stage_init_phase = FIRST_STAGE_INIT_WAIT_PARA_ZERO;
			break;
		}
		case FIRST_STAGE_INIT_WAIT_PARA_ZERO: {
			bool in_overload = false;
			STS_Servo_InOverload(&servo2, &in_overload);
			if (true) {
				STS_Servo_SetGoalSpeed(&servo2, 0);
				STS_Servo_PositionCalibration(&servo2, STS_GetPositionInUnits(0));
				STS_Servo_SetOperatingMode(&servo2, STS_OP_MODE_POSITION_CONTROL);
				STS_Servo_SetGoalPosition(&servo2, STS_GetPositionInUnits(30));
				first_stage_init_phase = FIRST_STAGE_INIT_SEPA_ZERO;
			}
			break;
		}
		case FIRST_STAGE_INIT_SEPA_ZERO: {
			STS_Servo_SetOperatingMode(&servo3, STS_OP_MODE_SPEED_CONTROL);
			STS_Servo_SetGoalSpeed(&servo3, STS_GetSpeedInUnits(-10));
			first_stage_init_phase = FIRST_STAGE_INIT_WAIT_SEPA_ZERO;
			break;
		}
		case FIRST_STAGE_INIT_WAIT_SEPA_ZERO: {
			bool in_overload = false;
			STS_Servo_InOverload(&servo3, &in_overload);
			if (true) {
				STS_Servo_SetGoalSpeed(&servo3, 0);
				STS_Servo_PositionCalibration(&servo3, STS_GetPositionInUnits(0));
				STS_Servo_SetOperatingMode(&servo3, STS_OP_MODE_POSITION_CONTROL);
				STS_Servo_SetGoalPosition(&servo3, STS_GetPositionInUnits(30));
				first_stage_init_phase = FIRST_STAGE_INIT_WAIT_JACK;
			}
			break;
		}
		case FIRST_STAGE_INIT_WAIT_JACK: {
			if (HAL_GPIO_ReadPin(IN_TRG_N1_GPIO_Port, IN_TRG_N1_Pin) == GPIO_PIN_SET) {
				first_stage_init_phase = FIRST_STAGE_INIT_WAIT_STAGE_ASSEMBLY_CONFIRMATION;
			} else {
				HAL_GPIO_TogglePin(LED1G_GPIO_Port, LED1G_Pin);
				HAL_Delay(500);
				HAL_GPIO_TogglePin(LED1G_GPIO_Port, LED1G_Pin);
				HAL_Delay(500);
			}
			break;
		}
		case FIRST_STAGE_INIT_WAIT_STAGE_ASSEMBLY_CONFIRMATION: {
			if (HAL_GPIO_ReadPin(IN_TRG_N2_GPIO_Port, IN_TRG_N2_Pin) == GPIO_PIN_SET) {
				HAL_GPIO_WritePin(LED1B_GPIO_Port, LED1B_Pin, GPIO_PIN_RESET);
				HAL_Delay(3000);
				if (HAL_GPIO_ReadPin(IN_TRG_N2_GPIO_Port, IN_TRG_N2_Pin) == GPIO_PIN_SET) {
					HAL_GPIO_WritePin(LED1B_GPIO_Port, LED1B_Pin, GPIO_PIN_SET);
					first_stage_init_phase = FIRST_STAGE_INIT_LOCK_STAGE;
				}
			} else {
				HAL_GPIO_TogglePin(LED1B_GPIO_Port, LED1B_Pin);
				HAL_Delay(500);
				HAL_GPIO_TogglePin(LED1B_GPIO_Port, LED1B_Pin);
				HAL_Delay(500);
			}
			break;
		}
		case FIRST_STAGE_INIT_LOCK_STAGE: {
			HAL_Delay(1000);
			HAL_GPIO_WritePin(LED1B_GPIO_Port, LED1B_Pin, GPIO_PIN_SET);
			STS_Servo_SetGoalPosition(&servo3, STS_GetPositionInUnits(360));
			HAL_Delay(500);
			first_stage_init_phase = FIRST_STAGE_INIT_WAIT_LOCK_STAGE;
			break;
		}
		case FIRST_STAGE_INIT_WAIT_LOCK_STAGE: {
			HAL_Delay(1000);
			HAL_GPIO_TogglePin(LED1R_GPIO_Port, LED1R_Pin);
			HAL_Delay(500);
			HAL_GPIO_TogglePin(LED1R_GPIO_Port, LED1R_Pin);
			first_stage_flight_phase = FIRST_STAGE_FLIGHT_WAIT_LAUNCH_CONFIRMATION;
			break;
		}
	}
}

void first_stage_flight_state_machine(void) {
	// FIRST_STAGE_INITIALISATION,
	// FIRST_STAGE_WAIT_LAUNCH_CONFIRMATION,
	// FIRST_STAGE_WAIT_BURN_END,
	// FIRST_STAGE_SEPARATION,
	// FIRST_STAGE_WAIT_SEPARATION_CONFIRMATION,
	// FIRST_STAGE_WAIT_APOGEE_CONFIRMATION,

	current_stage_phase_transition.stage_phase_type = STAGE_PHASE_FIRST_STAGE_FLIGHT;
	current_stage_phase_transition.phase_variable = (uint8_t *)&first_stage_flight_phase;

	switch (first_stage_flight_phase) {
		case FIRST_STAGE_FLIGHT_INITIALISATION: {
			first_stage_init_state_machine();
			break;
		}
		case FIRST_STAGE_FLIGHT_WAIT_LAUNCH_CONFIRMATION: {
			// Wait for launch confirmation
			if (HAL_GPIO_ReadPin(IN_TRG_N1_GPIO_Port, IN_TRG_N1_Pin) == GPIO_PIN_RESET) {
				t_launch = HAL_GetTick();

				HAL_GPIO_TogglePin(LED1G_GPIO_Port, LED1G_Pin);
				HAL_Delay(100);
				HAL_GPIO_TogglePin(LED1G_GPIO_Port, LED1G_Pin);
				is_launch_confirmed = true;
				change_state_and_notify(FIRST_STAGE_FLIGHT_WAIT_BURN_END);
			}
			break;
		}
		case FIRST_STAGE_FLIGHT_WAIT_BURN_END: {
			if (HAL_GetTick() - t_launch > T_ALPHA_BETA_0) {
				HAL_GPIO_TogglePin(LED1R_GPIO_Port, LED1R_Pin);
				first_stage_flight_phase = FIRST_STAGE_FLIGHT_SEPARATION;
				change_state_and_notify(FIRST_STAGE_FLIGHT_SEPARATION);
			}
			break;
		}
		case FIRST_STAGE_FLIGHT_SEPARATION: {
			if (HAL_GetTick() - t_launch > T_ALPHA_BETA_1) {
				HAL_GPIO_TogglePin(LED1R_GPIO_Port, LED1R_Pin);
				STS_Servo_SetGoalPosition(&servo1, STS_GetPositionInUnits(120));
				HAL_Delay(1);
				STS_Servo_SetGoalPosition(&servo3, STS_GetPositionInUnits(120));
				HAL_Delay(1);
				change_state_and_notify(FIRST_STAGE_FLIGHT_WAIT_SEPARATION_CONFIRMATION);
			}
			break;
		}
		case FIRST_STAGE_FLIGHT_WAIT_SEPARATION_CONFIRMATION: {
			switch (window_time_get_state(&window_time_sepa, HAL_GetTick() - t_launch)) {
				case WINDOW_TIME_STATE_WAITING: {
					// Still waiting for separation confirmation
					break;
				}
				case WINDOW_TIME_STATE_ACTIVE: {
					if (HAL_GPIO_ReadPin(IN_TRG_N2_GPIO_Port, IN_TRG_N2_Pin) == GPIO_PIN_RESET) {

						HAL_GPIO_TogglePin(LED1B_GPIO_Port, LED1B_Pin);
						HAL_Delay(100);
						HAL_GPIO_TogglePin(LED1B_GPIO_Port, LED1B_Pin);

						is_separation_confirmed = true;
						window_time_alpha_apogee = window_time_alpha_beta_apogee_sepa;
						event_uart_producer_add_event(&event_uart_producer, event_uart_msg_format(
							HAL_GetTick(), EVENT_UART_TYPE_WINDOW_TIME, (event_uart_payload_u){ .window_time_payload = {
								.window_id = window_time_sepa.id,
								.result = EVENT_UART_WINDOW_TIME_RESULT_PASS
							}}
						));
						change_state_and_notify(FIRST_STAGE_FLIGHT_WAIT_APOGEE_CONFIRMATION);
					}
					break;
				}
				case WINDOW_TIME_STATE_EXPIRED: {
					is_separation_confirmed = false;
					window_time_alpha_apogee = window_time_alpha_apogee_no_sepa;
					event_uart_producer_add_event(&event_uart_producer, event_uart_msg_format(
						HAL_GetTick(), EVENT_UART_TYPE_WINDOW_TIME, (event_uart_payload_u){ .window_time_payload = {
							.window_id = window_time_sepa.id,
							.result = EVENT_UART_WINDOW_TIME_RESULT_TIMEOUT
						}}
					));
					change_state_and_notify(FIRST_STAGE_FLIGHT_WAIT_APOGEE_CONFIRMATION);
					break;
				}
			}
			break;
		}
		case FIRST_STAGE_FLIGHT_WAIT_APOGEE_CONFIRMATION: {
			switch (window_time_get_state(&window_time_alpha_apogee, HAL_GetTick() - t_launch)) {
				case WINDOW_TIME_STATE_WAITING: {
					// Still waiting for apogee confirmation
					break;
				}
				case WINDOW_TIME_STATE_ACTIVE: {
					if (current_pressure_variation_pa_s < 5.0f) { // If we are near apogee (low pressure variation), confirm apogee
						// Apogee confirmed
						event_uart_producer_add_event(&event_uart_producer, event_uart_msg_format(
							HAL_GetTick(), EVENT_UART_TYPE_WINDOW_TIME, (event_uart_payload_u){ .window_time_payload = {
								.window_id = window_time_alpha_apogee.id,
								.result = EVENT_UART_WINDOW_TIME_RESULT_PASS
							}}
						));
						change_state_and_notify(FIRST_STAGE_FLIGHT_APOGEE);
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
					change_state_and_notify(FIRST_STAGE_FLIGHT_APOGEE);
					break;
				}
			}
			break;
		}
		case FIRST_STAGE_FLIGHT_APOGEE: {
			STS_Servo_SetGoalPosition(&servo2, STS_GetPositionInUnits(120));
			HAL_GPIO_TogglePin(LED1R_GPIO_Port, LED1R_Pin);
			HAL_Delay(500);
			HAL_GPIO_TogglePin(LED1R_GPIO_Port, LED1R_Pin);
			change_state_and_notify(FIRST_STAGE_FLIGHT_IDLE);
			break;
		}
		case FIRST_STAGE_FLIGHT_IDLE: {
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
