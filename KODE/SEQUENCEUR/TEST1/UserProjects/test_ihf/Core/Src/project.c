#include "project.h"

#include "main.h"

#include "stm32f0xx_hal.h"

#include "tim.h"


#include "led.h"
#include "waveform.h"

#include <math.h>
#include <stdbool.h>
#include <stdint.h>

static led_rgb_t led_rgb2;

void setup() {
	LED_Init(&led_rgb2.red  , &htim3, TIM_CHANNEL_1);
	LED_Init(&led_rgb2.green, &htim3, TIM_CHANNEL_2);
	LED_Init(&led_rgb2.blue , &htim3, TIM_CHANNEL_3);

	LED_SetBrightness(&led_rgb2.red  , 0.0f);
	LED_SetBrightness(&led_rgb2.green, 0.0f);
	LED_SetBrightness(&led_rgb2.blue , 0.0f);
}

void loop() {
	for (int i = 0; i < 256; i++) {
		LED_SetBrightness(&led_rgb2.blue, (float_t)i / 255.0f);
        HAL_Delay(10);
	}
}