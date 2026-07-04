#include "project.h"
#include "float3.h"
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



// waveform_space_t wait_button_waveform

static rocket_state_t rocket_state;

static led_rgb_t led_rgb2;

static stage_phase_type_t current_stage_phase_type;

static waiting_button_t wait;

void setup() {

	HAL_GPIO_WritePin(LED1R_GPIO_Port, LED1R_Pin, GPIO_PIN_SET);
	HAL_GPIO_WritePin(LED1G_GPIO_Port, LED1G_Pin, GPIO_PIN_SET);
	HAL_GPIO_WritePin(LED1B_GPIO_Port, LED1B_Pin, GPIO_PIN_SET);

	LED_Init(&led_rgb2.red  , &htim3, TIM_CHANNEL_1);
	LED_Init(&led_rgb2.green, &htim3, TIM_CHANNEL_2);
	LED_Init(&led_rgb2.blue , &htim3, TIM_CHANNEL_3);

	LED_RGB_SetColor(&led_rgb2, FLOAT3_ZERO);

	rocket_state_init(&rocket_state, &led_rgb2);
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

	waiting_button_init(&wait, PRGM_RUN_GPIO_Port, PRGM_RUN_Pin, &rocket_state.current_waveform_space, false);

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
			uint8_t prgm = 1;
			if (waiting_button_play(&wait)) {
				prgm = 0x00 | (
					(HAL_GPIO_ReadPin(PRGM0_GPIO_Port, PRGM0_Pin) == GPIO_PIN_SET ? 1 : 0) << 0 |
					(HAL_GPIO_ReadPin(PRGM1_GPIO_Port, PRGM1_Pin) == GPIO_PIN_SET ? 1 : 0) << 1 |
					(HAL_GPIO_ReadPin(PRGM2_GPIO_Port, PRGM2_Pin) == GPIO_PIN_SET ? 1 : 0) << 2 |
					(HAL_GPIO_ReadPin(PRGM3_GPIO_Port, PRGM3_Pin) == GPIO_PIN_SET ? 1 : 0) << 3
				);

				if (prgm != 0x00) {
					LED_RGB_SetColor(rocket_state.led_rgb, FLOAT3_ZERO);
					board_funcs[prgm](&rocket_state);
				} else {
					current_stage_phase_type = STAGE_PHASE_STAGE_INIT;
					for (uint8_t i = 1; i < 3; i++) {
						LED_RGB_SetColor(rocket_state.led_rgb, (float3_t){ .x = 1.0f, .y = 0.0f, .z = 0.0f });
						HAL_Delay(100);
						LED_RGB_SetColor(rocket_state.led_rgb, (float3_t){ .x = 0.0f, .y = 1.0f, .z = 0.0f });
						HAL_Delay(100);
						LED_RGB_SetColor(rocket_state.led_rgb, (float3_t){ .x = 0.0f, .y = 0.0f, .z = 1.0f });
						HAL_Delay(100);
					}
					LED_RGB_SetColor(rocket_state.led_rgb, FLOAT3_ZERO);
					HAL_Delay(500);
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

	float3_t color = FLOAT3_ZERO;
	if (rocket_state.current_waveform_space != NULL) {
		color = Waveform_Play_Space(rocket_state.current_waveform_space, HAL_GetTick());	
	}
	LED_RGB_SetColor(rocket_state.led_rgb, color);
}
