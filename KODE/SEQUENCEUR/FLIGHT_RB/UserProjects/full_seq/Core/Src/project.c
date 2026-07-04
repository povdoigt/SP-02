#include "project.h"
#include "float3.h"
#include "led_scheduler.h"
#include "main.h"

#include "stm32f0xx_hal_gpio.h"
#include "usart.h"

#include "full_seq_utils.h"

#include "stage_1.h"
#include "stage_2.h"

#include "led.h"

#include "waveform.h"
#include "event_uart.h"
#include "waveform_def.h"

#include <stdint.h>



static board_func_t board_funcs[16];

static rocket_state_t rocket_state;

static led_rgb_t led_rgb2;

static stage_phase_type_t current_stage_phase_type;

static uint8_t current_board_func_id = BOARD_FUNC_NONE;

void setup() {

	LED_Init(&led_rgb2.red  , &htim3, TIM_CHANNEL_1);
	LED_Init(&led_rgb2.green, &htim3, TIM_CHANNEL_2);
	LED_Init(&led_rgb2.blue , &htim3, TIM_CHANNEL_3);

	LED_RGB_SetColor(&led_rgb2, FLOAT3_ZERO);

	rocket_state_init(&rocket_state, &led_rgb2);
	LedSched_Init();
	led_states_init();
	current_stage_phase_type = STAGE_PHASE_STAGE_BOARD_FUNC;

	event_uart_producer_init(&event_uart_producer, TX_OPTO_N1_GPIO_Port, TX_OPTO_N1_Pin, &huart4);

	// Looking for which stage we are
	if (HAL_GPIO_ReadPin(STAGE1_GPIO_Port, STAGE1_Pin) == GPIO_PIN_RESET) {
		rocket_state.stage = ROCKET_FIRST_STAGE;
		board_funcs[1]  = board_func_1_stage_1;
		board_funcs[2]  = board_func_2_stage_1;
		board_funcs[3]  = board_func_3_stage_1;
		board_funcs[4]  = board_func_4_stage_1;
		board_funcs[5]  = board_func_5_stage_1;
		board_funcs[6]  = board_func_6_stage_1;
		board_funcs[7]  = board_func_7_stage_1;
		board_funcs[8]  = board_func_8_stage_1;
		board_funcs[9]  = board_func_9_stage_1;
		board_funcs[10] = board_func_10_stage_1;
		board_funcs[11] = board_func_11_stage_1;
		board_funcs[12] = board_func_12_stage_1;
		board_funcs[13] = board_func_13_stage_1;
		board_funcs[14] = board_func_14_stage_1;
		board_funcs[15] = board_func_15_stage_1;
	} else if (HAL_GPIO_ReadPin(STAGE2_GPIO_Port, STAGE2_Pin) == GPIO_PIN_RESET) {
		rocket_state.stage = ROCKET_SECOND_STAGE;
		board_funcs[1]  = board_func_1_stage_2;
		board_funcs[2]  = board_func_2_stage_2;
		board_funcs[3]  = board_func_3_stage_2;
		board_funcs[4]  = board_func_4_stage_2;
		board_funcs[5]  = board_func_5_stage_2;
		board_funcs[6]  = board_func_6_stage_2;
		board_funcs[7]  = board_func_7_stage_2;
		board_funcs[8]  = board_func_8_stage_2;
		board_funcs[9]  = board_func_9_stage_2;
		board_funcs[10] = board_func_10_stage_2;
		board_funcs[11] = board_func_11_stage_2;
		board_funcs[12] = board_func_12_stage_2;
		board_funcs[13] = board_func_13_stage_2;
		board_funcs[14] = board_func_14_stage_2;
		board_funcs[15] = board_func_15_stage_2;

	} else {
		LED_RGB_SetColor(&led_rgb2, (float3_t){ .x = 1.0f, .y = 0.0f, .z = 0.0f });
		Error_Handler();
	}

	waiting_button_init(&waiting_button, false);

	switch (rocket_state.stage) {
		case ROCKET_FIRST_STAGE: {
			setup_stage_1(&rocket_state);
			break;
		}
		case ROCKET_SECOND_STAGE: {
			setup_stage_2(&rocket_state);
			break;
		}
		case ROCKET_STAGE_NO_SET: {
			LED_RGB_SetColor(&led_rgb2, (float3_t){ .x = 1.0f, .y = 0.0f, .z = 0.0f });
			Error_Handler();
		}
	}
}

void loop() {

	update_gpio_input_states(
		rocket_state.input_gpio_ports,
		rocket_state.input_gpio_pins,
		rocket_state.input_gpio_states,
		MAX_INPUT_NUMBER
	);


	switch (current_stage_phase_type) {
		case STAGE_PHASE_STAGE_BOARD_FUNC: {
			if (current_board_func_id == BOARD_FUNC_NONE) {
				if (waiting_button_play(&waiting_button)) {
					uint8_t prgm = 0x00 | (
						(HAL_GPIO_ReadPin(PRGM0_GPIO_Port, PRGM0_Pin) == GPIO_PIN_SET ? 1 : 0) << 0 |
						(HAL_GPIO_ReadPin(PRGM1_GPIO_Port, PRGM1_Pin) == GPIO_PIN_SET ? 1 : 0) << 1 |
						(HAL_GPIO_ReadPin(PRGM2_GPIO_Port, PRGM2_Pin) == GPIO_PIN_SET ? 1 : 0) << 2 |
						(HAL_GPIO_ReadPin(PRGM3_GPIO_Port, PRGM3_Pin) == GPIO_PIN_SET ? 1 : 0) << 3
					);
					if (prgm == 0x00) {
						current_stage_phase_type = STAGE_PHASE_STAGE_INIT;
						LedSched_Clear();
						LedSched_Add(&waveform_prgm0_start, 0, false, 1000, LED_SCHED_NO_FORCE);
					} else if (prgm <= BOARD_FUNC_MAX_ID) {
						current_board_func_id = prgm;
					}
				}
			} else {
				board_func_t board_func = board_funcs[current_board_func_id];
				board_func_state_t board_func_state = board_func(&rocket_state);
				if (board_func_state == BOARD_FUNC_STATE_DONE) {
					current_board_func_id = BOARD_FUNC_NONE;
				}
			}
			break;
		}
		case STAGE_PHASE_STAGE_INIT:
		case STAGE_PHASE_STAGE_FLIGHT: {
			switch (rocket_state.stage) {
				case ROCKET_FIRST_STAGE: {
					loop_stage_1(&rocket_state);
					break;
				}
				case ROCKET_SECOND_STAGE: {
					loop_stage_2(&rocket_state);
					break;
				}
				case ROCKET_STAGE_NO_SET: {
					// Wait what ????
				}
			}
			break;
		}
	}



	event_uart_producer_run(&event_uart_producer);

	LED_RGB_SetColor(rocket_state.led_rgb, LedSched_Process(HAL_GetTick()));
}
