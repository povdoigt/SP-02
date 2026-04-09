#include "led.h"

#include <stdbool.h>
#include <stdint.h>



void LED_Init(led_t *led, TIM_HandleTypeDef *timer, uint32_t channel) {
    led->timer = timer;
    led->channel = channel;

    TIM_set_frequency(timer, LED_TIMER_FREQUENCY); // Set the timer frequency for PWM
    HAL_TIM_PWM_Start(timer, channel);
}

void LED_SetBrightness(led_t *led, float brightness) {
    brightness = brightness < 0.0f ? 0.0f : (brightness > 1.0f ? 1.0f : brightness); // Clamp to [0, 1]
    // from [0.0 1.0] to [0 tim->Period]
    uint32_t period = __HAL_TIM_GET_AUTORELOAD(led->timer);
    uint32_t brightness_32 = brightness * period;
    __HAL_TIM_SET_COMPARE(led->timer, led->channel, period - brightness_32);
}



void LED_RGB_SetColor(led_rgb_t *led_rgb, float3_t color) {
    LED_SetBrightness(&led_rgb->red, color.x);
    LED_SetBrightness(&led_rgb->green, color.y);
    LED_SetBrightness(&led_rgb->blue, color.z);
}
