#include "project.h"

#include "float3.h"
#include "main.h"
#include "quaternion.h"
#include "stm32f0xx_hal.h"
#include "stm32f0xx_hal_gpio.h"
#include "usart.h"

#include "usbd_cdc_if.h"

#include "led.h"
#include "WT901B.h"

#include "iir_filter.h"
#include "quaternion_dynamics.h"

#include <stddef.h>
#include <stdint.h>
#include <stdio.h>
#include <string.h>
#include <math.h>

static uint8_t uart1_buffer[WT901B_RX_BUFFER_SIZE];

static data_sub_t pressure_sub = { 0 };

static float current_altitude_m = 0.0f;
static float variation_altitude_m_s;

void on_new_pressure_frame(data_sub_t *sub);

void setup() {

	UART_buffer_init(&huart1, uart1_buffer, sizeof(uart1_buffer));

	WT901B_status_t wt_res = WT901B_Init(&wt901b, &huart1);
	if (wt_res != WT901B_OK) { Error_Handler(); }

	WT901B_Write(&wt901b, WT901B_REG_RRATE, WT901B_RRATE_20HZ);

	data_sub_attach(&pressure_sub, &wt901b.data_topic, DATA_ATTACH_FROM_OLDEST);
}

void loop() {
	uint8_t buffer[256];

	WT901B_Parse_Frames(&wt901b);

	on_new_pressure_frame(&pressure_sub);

	snprintf((char*)buffer, sizeof(buffer), "Baro : %-05.2f %05.2f\r\n", variation_altitude_m_s, current_altitude_m);

	CDC_Transmit_FS(buffer, strlen((char *)buffer) + 1);

	HAL_Delay(1);

}

void on_new_pressure_frame(data_sub_t *sub) {
	if (data_sub_num_to_read(sub) > 0) {
		WT901B_Frame_t frame;
		data_sub_read(sub, &frame);
		if (frame.type == WT901B_FRAME_PRESSURE) {
			// Assuming we receive pressure in Pa, we can compute the variation if we store the previous pressure
			static float previous_altitude_m = 0.0f;
			current_altitude_m = (float)frame.data.pressure.altitude_cm / 100.0f; // convert to meters
			if (previous_altitude_m > 0.0f) {
				variation_altitude_m_s = (current_altitude_m - previous_altitude_m) / 0.05f; // variation en m/s
			}
			previous_altitude_m = current_altitude_m;
		}
	}
}