#include "project.h"

#include "main.h"

#include "stm32f0xx_hal.h"

#include "stm32f0xx_hal_flash_ex.h"
#include "stm32f0xx_hal_gpio.h"
#include "tim.h"

#include "led.h"
#include "waveform.h"
#include "waveform_built_in.h"
#include "waveform_def.h"

#include <math.h>
#include <stdbool.h>
#include <stdint.h>

static led_rgb_t led_rgb2;

static waveform_space_t waveform;
static waveform_space_add_t waveform_ctx =  {
    .wave_function_add_1 = waveform_space_mult_const,
    .ctx_add_1 = &(waveform_space_mult_const_t) {
        .wave_function_mult = waveform_sine,
        .ctx_mult = &(waveform_sine_t){ .start = 0.0f, .end = 0.5f },
        .factor = (float3_t){ .x = 1.0f, .y = 0.0f, .z = 0.0f }
    },
    .wave_function_add_2 = waveform_space_add,
    .ctx_add_2 = &(waveform_space_add_t) {
        .wave_function_add_1 = waveform_space_mult_const,
        .ctx_add_1 = &(waveform_space_mult_const_t) {
            .wave_function_mult = waveform_sine,
            .ctx_mult = &(waveform_sine_t){ .start = 0.25f, .end = 0.75f },
            .factor = (float3_t){ .x = 0.0f, .y = 1.0f, .z = 1.0f }
        },
        .wave_function_add_2 = waveform_space_mult_const,
        .ctx_add_2 = &(waveform_space_mult_const_t) {
            .wave_function_mult = waveform_sine,
            .ctx_mult = &(waveform_sine_t){ .start = 0.5f, .end = 1.0f },
            .factor = (float3_t){ .x = 0.0f, .y = 0.0f, .z = 1.0f }
        }
    }
};

static bool prgm_running = false;
static uint8_t prgm_state = 0;

void setup() {
	LED_Init(&led_rgb2.red  , &htim3, TIM_CHANNEL_1);
	LED_Init(&led_rgb2.green, &htim3, TIM_CHANNEL_2);
	LED_Init(&led_rgb2.blue , &htim3, TIM_CHANNEL_3);

	LED_SetBrightness(&led_rgb2.red  , 0.0f);
	LED_SetBrightness(&led_rgb2.green, 0.0f);
	LED_SetBrightness(&led_rgb2.blue , 0.0f);

    Waveform_Init_Space(&waveform, waveform_space_add, &waveform_ctx, 5000, true);
}

void loop() {
    if (Waveform_IsActive((waveform_generic_t *)&waveform)) {
        float3_t color = Waveform_Play_Space(&waveform, HAL_GetTick());
        LED_RGB_SetColor(&led_rgb2, color);
    }

    if (HAL_GPIO_ReadPin(PRGM_RUN_GPIO_Port, PRGM_RUN_Pin)) {
        uint8_t state = (HAL_GPIO_ReadPin(PRGM0_GPIO_Port, PRGM0_Pin) << 0) |
                        (HAL_GPIO_ReadPin(PRGM1_GPIO_Port, PRGM1_Pin) << 1) |
                        (HAL_GPIO_ReadPin(PRGM2_GPIO_Port, PRGM2_Pin) << 2) |
                        (HAL_GPIO_ReadPin(PRGM3_GPIO_Port, PRGM3_Pin) << 3);
        for (int i = 0; i < state; i++) {
            
        }
    }
}
