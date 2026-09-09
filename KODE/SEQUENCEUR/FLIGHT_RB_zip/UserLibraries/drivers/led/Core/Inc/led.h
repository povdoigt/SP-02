#ifndef LED_H
#define LED_H

#include "tim.h"
#include "float3.h"

#include <stdbool.h>
#include <stdint.h>

#define LED_TIMER_FREQUENCY 10000 // Frequency for PWM (10 kHz)

typedef struct led_t {
    TIM_HandleTypeDef *timer; // Timer handle for PWM
    uint32_t channel; // PWM channel
} led_t;

void LED_Init(led_t *led, TIM_HandleTypeDef *timer, uint32_t channel);
void LED_SetBrightness(led_t *led, float brightness);

typedef struct led_rgb_t {
    led_t red;
    led_t green;
    led_t blue;
} led_rgb_t;

void LED_RGB_SetColor(led_rgb_t *led_rgb, float3_t color);


#endif /* LED_H */