#include "project.h"

#include "float3.h"
#include "main.h"
#include "stm32f0xx_hal.h"
#include "usart.h"

// #include "usbd_cdc_if.h"

#include "WT901B.h"

#include <stdint.h>
#include <stdio.h>
#include <string.h>


static data_sub_t accel_sub = { 0 };
static data_sub_t gyro_sub = { 0 };
static data_sub_t pressure_sub = { 0 };

static float3_t last_acc;
static float3_t last_gyro;
static float	last_baro;

static uint8_t buffer[256];

void on_new_accel_frame(data_sub_t *sub);
void on_new_gyro_frame(data_sub_t *sub);
void on_new_pressure_frame(data_sub_t *sub);



void setup() {
	WT901B_status_t wt_res = WT901B_Init(&wt901b, &huart1);
	if (wt_res != WT901B_OK) { Error_Handler(); }

	data_sub_attach(&accel_sub, &wt901b.data_topic, DATA_ATTACH_FROM_OLDEST);
	data_sub_attach(&gyro_sub, &wt901b.data_topic, DATA_ATTACH_FROM_OLDEST);
	data_sub_attach(&pressure_sub, &wt901b.data_topic, DATA_ATTACH_FROM_OLDEST);
}

void loop() {
	WT901B_Parse_Frames(&wt901b);

	on_new_accel_frame(&accel_sub);
	on_new_gyro_frame(&gyro_sub);
	on_new_pressure_frame(&pressure_sub);

	snprintf((char *)buffer, sizeof(buffer), "Acc : %3.2f, %3.2f, %3.2f ; Gyr : %3.2f, %3.2f, %3.2f ; Baro : %6.2f\r\n",
		last_acc.x, last_acc.y, last_acc.z, last_gyro.x, last_gyro.y, last_gyro.z, last_baro
	);

	// CDC_Transmit_FS(buffer, strlen((char *)buffer) + 1);

	HAL_Delay(1);

}


void on_new_accel_frame(data_sub_t *sub) {
	if (data_sub_num_to_read(sub) > 0) {
		WT901B_Frame_t frame;
		data_sub_read(sub, &frame);
		if (frame.type == WT901B_FRAME_ACCEL) {
			last_acc.x = frame.data.accel.ax_g;
			last_acc.y = frame.data.accel.ay_g;
			last_acc.z = frame.data.accel.az_g;
		}
	}
}

void on_new_gyro_frame(data_sub_t *sub) {
	if (data_sub_num_to_read(sub) > 0) {
		WT901B_Frame_t frame;
		data_sub_read(sub, &frame);
		if (frame.type == WT901B_FRAME_GYRO) {
			last_gyro.x = frame.data.gyro.gx_dps;
			last_gyro.y = frame.data.gyro.gy_dps;
			last_gyro.z = frame.data.gyro.gz_dps;
		}
	}
}

void on_new_pressure_frame(data_sub_t *sub) {
	if (data_sub_num_to_read(sub) > 0) {
		WT901B_Frame_t frame;
		data_sub_read(sub, &frame);
		if (frame.type == WT901B_FRAME_PRESSURE) {
			last_baro = (float)frame.data.pressure.pressure_pa;
		}
	}
}