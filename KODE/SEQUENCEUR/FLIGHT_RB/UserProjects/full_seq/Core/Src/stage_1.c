#include "stage_1.h"

#include "full_seq_utils.h"
#include "led.h"
#include "STS.h"
#include "WT901B.h"

#include "event_uart.h"
#include "window_time.h"





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

static uint8_t uart1_buffer[WT901B_RX_BUFFER_SIZE];
static uint8_t uart2_buffer[STS_SERIAL_BUFFER_SIZE];
static uint8_t uart3_buffer[STS_SERIAL_BUFFER_SIZE];
static uint8_t uart4_buffer[sizeof(event_uart_msg_t)];

static first_stage_initialisation_phase_t first_stage_init_phase;
static first_stage_initialisation_phase_t first_stage_init_next_phase;
static first_stage_flight_phase_t first_stage_flight_phase;

static waiting_button_t waiting_button;



void setup_stage_1(rocket_state_t *rocket_state) {
	
	LED_RGB_SetColor(rocket_state->led_rgb, (float3_t){ .x = 0.0, .y = 1.0, .z = 0.0});
	HAL_Delay(500);
	LED_RGB_SetColor(rocket_state->led_rgb, FLOAT3_ZERO);
	HAL_Delay(500);

    setup_uart_buffers_stage_1(rocket_state);
    setup_servomotors_stage_1(rocket_state);

	first_stage_init_next_phase = FIRST_STAGE_INIT_AF_ZERO;

    phase_transition_init(&rocket_state->stage_phase_transition, STAGE_PHASE_STAGE_FLIGHT, (uint8_t*)&first_stage_flight_phase);
	change_state_and_notify(&rocket_state->stage_phase_transition, FIRST_STAGE_FLIGHT_INITIALISATION);
	phase_transition_init(&rocket_state->stage_phase_transition, STAGE_PHASE_STAGE_INIT, (uint8_t*)&first_stage_init_phase);
	first_stage_init_phase = FIRST_STAGE_INIT_WAIT_BUTTON;

	waiting_button_init(&waiting_button, PRGM_RUN_GPIO_Port, PRGM_RUN_Pin, rocket_state->led_rgb, false);
}

void loop_stage_1(rocket_state_t *rocket_state) {
    // WT901B_Parse_Frames(&wt901b);

    // on_new_pressure_frame(rocket_state, &pressure_sub);

	first_stage_flight_state_machine(rocket_state);
}


/* ===================================================
   SETUP FUNCTIONS
   =================================================== */


void setup_uart_buffers_stage_1(rocket_state_t *rocket_state) {
	UART_buffer_init(&huart1, uart1_buffer, sizeof(uart1_buffer));
	UART_buffer_init(&huart2, uart2_buffer, sizeof(uart2_buffer));
	UART_buffer_init(&huart3, uart3_buffer, sizeof(uart3_buffer));
	UART_buffer_init(&huart4, uart4_buffer, sizeof(uart4_buffer));
}

void setup_servomotors_stage_1(rocket_state_t *rocket_state) {
	HAL_StatusTypeDef res;

	res = STS_UART_Port_Init(&huart_sts_port1, &huart2);
	if (res != HAL_OK) { goto error; }
	res = STS_UART_Port_Init(&huart_sts_port2, &huart3);
	if (res != HAL_OK) { goto error; }

	res = STS_Servo_Init(&servo1, &huart_sts_port1, 1);
	if (res != HAL_OK) { goto error; }
	res = STS_Servo_Init(&servo2, &huart_sts_port1, 2);
	if (res != HAL_OK) { goto error; }
	res = STS_Servo_Init(&servo3, &huart_sts_port2, 3);
	if (res != HAL_OK) { goto error; }
	HAL_Delay(100);
    return;

error:
    LED_RGB_SetColor(rocket_state->led_rgb, (float3_t){ .x = 1.0f, .y = 0.0f, .z = 0.0f });
    Error_Handler();
}




/* ===================================================
   FIRST STAGE STATE MACHINE
   =================================================== */


void first_stage_init_state_machine(rocket_state_t *rocket_state) {

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
				HAL_Delay(100);
				STS_Servo_PositionCalibration(&servo1, STS_GetPositionInUnits(0));
				STS_Servo_SetOperatingMode(&servo1, STS_OP_MODE_POSITION_CONTROL);
				STS_Servo_SetGoalPosition(&servo1, STS_GetPositionInUnits(30));
				first_stage_init_next_phase = FIRST_STAGE_INIT_PARA_ZERO;
				first_stage_init_phase = FIRST_STAGE_INIT_WAIT_BUTTON;
			}
			break;
		}
		case FIRST_STAGE_INIT_PARA_ZERO: {
			STS_Servo_SetOperatingMode(&servo2, STS_OP_MODE_SPEED_CONTROL);
			STS_Servo_SetGoalSpeed(&servo2, STS_GetSpeedInUnits(-10));
			change_state_and_notify(&rocket_state->stage_phase_transition, FIRST_STAGE_INIT_WAIT_PARA_ZERO);
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
				first_stage_init_next_phase = FIRST_STAGE_INIT_SEPA_ZERO;
				first_stage_init_phase = FIRST_STAGE_INIT_WAIT_BUTTON;
			}
			break;
		}
		case FIRST_STAGE_INIT_SEPA_ZERO: {
			STS_Servo_SetOperatingMode(&servo3, STS_OP_MODE_SPEED_CONTROL);
			STS_Servo_SetGoalSpeed(&servo3, STS_GetSpeedInUnits(10));
			change_state_and_notify(&rocket_state->stage_phase_transition, FIRST_STAGE_INIT_WAIT_SEPA_ZERO);
			break;
		}
		case FIRST_STAGE_INIT_WAIT_SEPA_ZERO: {
			bool in_overload = false;
			STS_Servo_InOverload(&servo3, &in_overload);
			if (in_overload) {
				STS_Servo_SetGoalSpeed(&servo3, 0);
				HAL_Delay(100);
				STS_Servo_PositionCalibration(&servo3, STS_GetPositionInUnits(720));
				STS_Servo_SetOperatingMode(&servo3, STS_OP_MODE_POSITION_CONTROL);
				HAL_Delay(100);
				STS_Servo_SetGoalPosition(&servo3, STS_GetPositionInUnits(360));
				first_stage_init_next_phase = FIRST_STAGE_INIT_WAIT_JACK;
				first_stage_init_phase = FIRST_STAGE_INIT_WAIT_BUTTON;
			}
			break;
		}
		case FIRST_STAGE_INIT_WAIT_JACK: {
			if (HAL_GPIO_ReadPin(IN_TRG_N1_GPIO_Port, IN_TRG_N1_Pin) == GPIO_PIN_SET) {
				change_state_and_notify(&rocket_state->stage_phase_transition, FIRST_STAGE_INIT_WAIT_STAGE_ASSEMBLY_CONFIRMATION);
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
					first_stage_init_next_phase = FIRST_STAGE_INIT_LOCK_STAGE;
					first_stage_init_phase = FIRST_STAGE_INIT_WAIT_BUTTON;
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
			
			if (HAL_GPIO_ReadPin(PRGM_RUN_GPIO_Port, PRGM_RUN_Pin) == GPIO_PIN_SET) {
				phase_transition_init(&rocket_state->stage_phase_transition, STAGE_PHASE_STAGE_FLIGHT, (uint8_t *)&first_stage_flight_phase);
				change_state_and_notify(&rocket_state->stage_phase_transition, FIRST_STAGE_FLIGHT_WAIT_LAUNCH_CONFIRMATION);
			} else {
				LED_RGB_SetColor(rocket_state->led_rgb, (float3_t){ .x = 1.0, .y = 0.0, .z = 0.0});
				HAL_Delay(500);
				LED_RGB_SetColor(rocket_state->led_rgb, FLOAT3_ZERO);
				HAL_Delay(500);
			}
			break;
		}
		case FIRST_STAGE_INIT_WAIT_BUTTON: {
			if (waiting_button_play(&waiting_button)) {
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
			// Wait for launch confirmation
			if (HAL_GPIO_ReadPin(IN_TRG_N1_GPIO_Port, IN_TRG_N1_Pin) == GPIO_PIN_RESET) {
				rocket_state->t_launch = HAL_GetTick();

				HAL_GPIO_TogglePin(LED1G_GPIO_Port, LED1G_Pin);
				HAL_Delay(100);
				HAL_GPIO_TogglePin(LED1G_GPIO_Port, LED1G_Pin);
				rocket_state->is_launch_confirmed = true;
				change_state_and_notify(&rocket_state->stage_phase_transition, FIRST_STAGE_FLIGHT_WAIT_BURN_END);
			}
			break;
		}
		case FIRST_STAGE_FLIGHT_WAIT_BURN_END: {
			if (HAL_GetTick() - rocket_state->t_launch > T_ALPHA_BETA_0) {
				HAL_GPIO_TogglePin(LED1R_GPIO_Port, LED1R_Pin);
				first_stage_flight_phase = FIRST_STAGE_FLIGHT_SEPARATION;
				change_state_and_notify(&rocket_state->stage_phase_transition, FIRST_STAGE_FLIGHT_SEPARATION);
			}
			break;
		}
		case FIRST_STAGE_FLIGHT_SEPARATION: {
			if (HAL_GetTick() - rocket_state->t_launch > T_ALPHA_BETA_1) {
				HAL_GPIO_TogglePin(LED1R_GPIO_Port, LED1R_Pin);
				STS_Servo_SetGoalPosition(&servo1, STS_GetPositionInUnits(120));
				HAL_Delay(1);
				STS_Servo_SetGoalPosition(&servo3, STS_GetPositionInUnits(120));
				HAL_Delay(1);
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
					if (HAL_GPIO_ReadPin(IN_TRG_N2_GPIO_Port, IN_TRG_N2_Pin) == GPIO_PIN_RESET) {

						HAL_GPIO_TogglePin(LED1B_GPIO_Port, LED1B_Pin);
						HAL_Delay(100);
						HAL_GPIO_TogglePin(LED1B_GPIO_Port, LED1B_Pin);

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
					if (rocket_state->dynamics.pressure_variation_pa_s < 5.0f) { // If we are near apogee (low pressure variation), confirm apogee
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
			STS_Servo_SetGoalPosition(&servo2, STS_GetPositionInUnits(120));
			HAL_GPIO_TogglePin(LED1R_GPIO_Port, LED1R_Pin);
			HAL_Delay(500);
			HAL_GPIO_TogglePin(LED1R_GPIO_Port, LED1R_Pin);
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
