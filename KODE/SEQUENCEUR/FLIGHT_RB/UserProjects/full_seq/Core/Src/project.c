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



// waveform_space_t wait_button_waveform

static rocket_state_t rocket_state;

static led_rgb_t led_rgb2;





void setup() {

	HAL_GPIO_WritePin(LED1R_GPIO_Port, LED1R_Pin, GPIO_PIN_SET);
	HAL_GPIO_WritePin(LED1G_GPIO_Port, LED1G_Pin, GPIO_PIN_SET);
	HAL_GPIO_WritePin(LED1B_GPIO_Port, LED1B_Pin, GPIO_PIN_SET);

	LED_Init(&led_rgb2.red  , &htim3, TIM_CHANNEL_1);
	LED_Init(&led_rgb2.green, &htim3, TIM_CHANNEL_2);
	LED_Init(&led_rgb2.blue , &htim3, TIM_CHANNEL_3);

	LED_RGB_SetColor(&led_rgb2, FLOAT3_ZERO);

	init_rocket_state(&rocket_state, &led_rgb2);

	event_uart_producer_init(&event_uart_producer, TX_OPTO_N1_GPIO_Port, TX_OPTO_N1_Pin, &huart4);

	// Looking for which stage we are
	if (HAL_GPIO_ReadPin(STAGE1_GPIO_Port, STAGE1_Pin) == GPIO_PIN_RESET) {
		rocket_state.stage = ROCKET_FIRST_STAGE;
	} else if (HAL_GPIO_ReadPin(STAGE2_GPIO_Port, STAGE2_Pin) == GPIO_PIN_RESET) {
		rocket_state.stage = ROCKET_SECOND_STAGE;
	} else {
		LED_RGB_SetColor(&led_rgb2, (float3_t){ .x = 1.0f, .y = 0.0f, .z = 0.0f });
		Error_Handler();
	}

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

	event_uart_producer_run(&event_uart_producer);

	float3_t color = FLOAT3_ZERO;
	if (rocket_state.current_waveform_space != NULL) {
		color = Waveform_Play_Space(rocket_state.current_waveform_space, HAL_GetTick());	
	}
	LED_RGB_SetColor(&rocket_state.led_rgb, color);
}
