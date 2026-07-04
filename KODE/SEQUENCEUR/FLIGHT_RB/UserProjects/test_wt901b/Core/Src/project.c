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

#define DEG2RAD (M_PI / 180.0f)
#define RAD2DEG (180.0f / M_PI)
#define INIT_ELEVATION_DEG 0.0f

static data_sub_t accel_sub = { 0 };
static data_sub_t gyro_sub = { 0 };
static data_sub_t pressure_sub = { 0 };

static float3_t last_acc;
static float3_t last_gyro;
static float	last_baro;

static uint8_t buffer[256];

static rocket_attitude_dynamics_t attitude;

static float b_coef[5]; // Moving average filter coefficients
// static const float a_coef[] = {1,-3.7509592220591514, 6.148473086730176, -5.562037987327878, 2.9088047777759973, -0.8299899993964646, 0.10061670488299389};
// static const float b_coef[] = {1, 6, 15, 20, 15, 6, 1};
// static const float a_coef[] = {-1, +0.4476263137046185, -1.221579158227121, +0.7267857842672348, -0.49111348727170195, +0.23969269502475835};
// static const float b_coef[] = {1 / 24.642152580081603, 5, 10, 10, 5, 1};
static float wx_inp_storage[sizeof(b_coef) / sizeof(float) - 1];
static float wy_inp_storage[sizeof(b_coef) / sizeof(float) - 1];
static float wz_inp_storage[sizeof(b_coef) / sizeof(float) - 1];
static float elev_inp_storage[sizeof(b_coef) / sizeof(float) - 1];
// static float wx_out_storage[sizeof(a_coef) / sizeof(float) - 1];
// static float wy_out_storage[sizeof(a_coef) / sizeof(float) - 1];
// static float wz_out_storage[sizeof(a_coef) / sizeof(float) - 1];	
static iir_filter_t wx_iir_filter;
static iir_filter_t wy_iir_filter;
static iir_filter_t wz_iir_filter;
static iir_filter_t elev_iir_filter;

static bool is_launch_confirmed = false;

static led_rgb_t led_rgb2;

static bool button_pressed = false;

static uint8_t uart1_buffer[WT901B_RX_BUFFER_SIZE];

// static float3_t v_up_body_earth;

void on_new_accel_frame(data_sub_t *sub);
void on_new_gyro_frame(data_sub_t *sub);
void on_new_pressure_frame(data_sub_t *sub);

void setup_attitude(void) {
	// attitude.x_body = FLOAT3_UNIT_X;
	// attitude.y_body = FLOAT3_UNIT_Y;
	// attitude.z_body = FLOAT3_UNIT_Z;

	// attitude.v_up_body = FLOAT3_UNIT_Y; /* axe “up” de la fusée aligné avec l’axe Y du body */
	// // attitude.v_left_body = FLOAT3_UNIT_X; /* axe “left” de la fusée aligné avec l’axe X du body */
	// // attitude.initial_elevation_deg = INIT_ELEVATION_DEG * DEG2RAD; /* angle d’élévation initiale */
	// // attitude.q = quatf_from_axis_angle(FLOAT3_UNIT_Z, - M_PI / 2.0f); // rotation initiale pour aligner l’axe forward du repère Terre avec l’axe Y du body
	// // attitude.q = quatf_mul(attitude.q, quatf_from_axis_angle(FLOAT3_UNIT_X, attitude.initial_elevation_deg)); // rotation initiale d’élévation
	// attitude.elevation_deg = 0.0f;
	// attitude.azimuth_deg = 0.0f;
}

void compute_elevation_azimut(void) {

	if (!is_launch_confirmed) {
		attitude.q_init = attitude.q;
	}

	float3_t v_up_body_earth_init = float3_normalized(quatf_rotate_vector(attitude.q_init, FLOAT3_UNIT_Y));
	float3_t v_up_body_earth = float3_normalized(quatf_rotate_vector(attitude.q, FLOAT3_UNIT_Y));
	// attitude.elevation_deg = iir_process(&elev_iir_filter, 90 - acosf(v_up_body_earth.z) * RAD2DEG);
	attitude.elevation_deg = 90 - acosf(v_up_body_earth.z) * RAD2DEG;

	float3_t v_up_body_earth_init_proj = float3_normalized(float3_sub(v_up_body_earth_init, float3_scale(FLOAT3_UNIT_Z, v_up_body_earth_init.z)));
	float3_t v_up_body_earth_proj = float3_normalized(float3_sub(v_up_body_earth, float3_scale(FLOAT3_UNIT_Z, v_up_body_earth.z)));
	attitude.azimuth_deg = acosf(float3_dot(v_up_body_earth_init_proj, v_up_body_earth_proj)) * RAD2DEG;
}

void check_elevation_azimuth(void) {
	float delta_elev = fabsf(attitude.elevation_deg - 72.0f);
	float delta_azim = fabsf(attitude.azimuth_deg - 0.0f);

	bool is_elev_ok = delta_elev <= 10.0f;
	bool is_azim_ok = delta_azim <= 45.0f;

	if (is_launch_confirmed) {
		if (is_elev_ok && is_azim_ok) {
			LED_RGB_SetColor(&led_rgb2, (float3_t){ .x = 0.0f, .y = 1.0f, .z = 0.0f });
		} else if (is_elev_ok || is_azim_ok) {
			LED_RGB_SetColor(&led_rgb2, (float3_t){ .x = 0.0f, .y = 0.0f, .z = 1.0f });
		} else {
			LED_RGB_SetColor(&led_rgb2, (float3_t){ .x = 1.0f, .y = 0.0f, .z = 0.0f });
		}
	}
}


void setup() {

	UART_buffer_init(&huart1, uart1_buffer, sizeof(uart1_buffer));

	WT901B_status_t wt_res = WT901B_Init(&wt901b, &huart1);
	if (wt_res != WT901B_OK) { Error_Handler(); }

	WT901B_Write(&wt901b, WT901B_REG_RRATE, WT901B_RRATE_20HZ);

	LED_Init(&led_rgb2.red  , &htim3, TIM_CHANNEL_1);
	LED_Init(&led_rgb2.green, &htim3, TIM_CHANNEL_2);
	LED_Init(&led_rgb2.blue , &htim3, TIM_CHANNEL_3);
	LED_RGB_SetColor(&led_rgb2, FLOAT3_ZERO);

	HAL_GPIO_WritePin(LED1R_GPIO_Port, LED1R_Pin, GPIO_PIN_SET);
	HAL_GPIO_WritePin(LED1G_GPIO_Port, LED1G_Pin, GPIO_PIN_SET);
	HAL_GPIO_WritePin(LED1B_GPIO_Port, LED1B_Pin, GPIO_PIN_SET);

	size_t b_order = sizeof(b_coef) / sizeof(float) - 1;
	// size_t a_order = sizeof(a_coef) / sizeof(float) - 1;

	for (int i = 0; i < b_order + 1; i++) {
		b_coef[i] = 1.0f / (b_order + 1); // Moving average filter coefficients
	}

	iir_init(&wx_iir_filter, 0, b_order, NULL, b_coef, wx_inp_storage, NULL);
	iir_init(&wy_iir_filter, 0, b_order, NULL, b_coef, wy_inp_storage, NULL);
	iir_init(&wz_iir_filter, 0, b_order, NULL, b_coef, wz_inp_storage, NULL);
	// iir_init(&elev_iir_filter, 0, b_order, NULL, b_coef, elev_inp_storage, NULL);
	// iir_init(&wx_iir_filter, a_order, b_order, a_coef, b_coef, wx_inp_storage, wx_out_storage);
	// iir_init(&wy_iir_filter, a_order, b_order, a_coef, b_coef, wy_inp_storage, wy_out_storage);
	// iir_init(&wz_iir_filter, a_order, b_order, a_coef, b_coef, wz_inp_storage, wz_out_storage);

	setup_attitude();

	data_sub_attach(&accel_sub, &wt901b.data_topic, DATA_ATTACH_FROM_OLDEST);
	data_sub_attach(&gyro_sub, &wt901b.data_topic, DATA_ATTACH_FROM_OLDEST);
	data_sub_attach(&pressure_sub, &wt901b.data_topic, DATA_ATTACH_FROM_OLDEST);
}

void loop() {
	WT901B_Parse_Frames(&wt901b);

	on_new_accel_frame(&accel_sub);
	on_new_gyro_frame(&gyro_sub);
	on_new_pressure_frame(&pressure_sub);

	compute_elevation_azimut();
	check_elevation_azimuth();

	
	if (!button_pressed && HAL_GPIO_ReadPin(PRGM_RUN_GPIO_Port, PRGM_RUN_Pin) == GPIO_PIN_SET) {
		button_pressed = true;
	}
	if (button_pressed && HAL_GPIO_ReadPin(PRGM_RUN_GPIO_Port, PRGM_RUN_Pin) == GPIO_PIN_RESET) {
		button_pressed = false;
		is_launch_confirmed = !is_launch_confirmed;
	}

	if (!is_launch_confirmed) {
		attitude.q = quatf_from_2_vec3(last_acc, FLOAT3_UNIT_Z);
		LED_RGB_SetColor(&led_rgb2, (float3_t){ .x = 1.0f, .y = 1.0f, .z = 0.0f });
	}

	// snprintf((char *)buffer, sizeof(buffer), "Acc : %3.2f, %3.2f, %3.2f ; Gyr : %3.2f, %3.2f, %3.2f ; Baro : %6.2f\r\n",
	// 	last_acc.x, last_acc.y, last_acc.z, last_gyro.x, last_gyro.y, last_gyro.z, last_baro
	// );
	// snprintf((char *)buffer, sizeof(buffer), "Elev : %3.2f ; Azim : %3.2f\r\n", attitude.elevation_deg, attitude.azimuth_deg);
	// snprintf((char*)buffer, sizeof(buffer), "Q: %3.2f, %3.2f, %3.2f, %3.2f\r\n", attitude.q.w, attitude.q.x, attitude.q.y, attitude.q.z);
	// snprintf((char*)buffer, sizeof(buffer), "%03.2f,%03.2f,%03.2f\n", v_up_body_earth.x, v_up_body_earth.y, v_up_body_earth.z);

	float3_t x_basis = quatf_rotate_vector(attitude.q, FLOAT3_UNIT_X);
	float3_t y_basis = quatf_rotate_vector(attitude.q, FLOAT3_UNIT_Y);
	float3_t z_basis = quatf_rotate_vector(attitude.q, FLOAT3_UNIT_Z);	

	snprintf((char*)buffer, sizeof(buffer), "%03.2f,%03.2f,%03.2f %03.2f,%03.2f,%03.2f %03.2f,%03.2f,%03.2f\r\n",
		z_basis.x, z_basis.y, z_basis.z,
		x_basis.x, x_basis.y, x_basis.z,
		y_basis.x, y_basis.y, y_basis.z
	);


	CDC_Transmit_FS(buffer, strlen((char *)buffer) + 1);

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
	if (is_launch_confirmed && data_sub_num_to_read(sub) > 0) {
		WT901B_Frame_t frame;
		data_sub_read(sub, &frame);
		if (frame.type == WT901B_FRAME_GYRO) {
			quatdyn_propagate_ip_3(
				&attitude.q,
				(float3_t){
					.x = iir_process(&wx_iir_filter, frame.data.gyro.gx_dps * DEG2RAD), /* convert to radians/s */
					.y = iir_process(&wy_iir_filter, frame.data.gyro.gy_dps * DEG2RAD),
					.z = iir_process(&wz_iir_filter, frame.data.gyro.gz_dps * DEG2RAD)
					// .x = frame.data.gyro.gx_dps * DEG2RAD, /* convert to radians/s */
					// .y = frame.data.gyro.gy_dps * DEG2RAD,
					// .z = frame.data.gyro.gz_dps * DEG2RAD
				},
				0.050f /* 50 ms, since we set the update rate to 20 Hz */
			);
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