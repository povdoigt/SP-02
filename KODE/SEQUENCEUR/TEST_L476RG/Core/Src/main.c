/* USER CODE BEGIN Header */
/**
 ******************************************************************************
 * @file           : main.c
 * @brief          : Main program body
 ******************************************************************************
 * @attention
 *
 * Copyright (c) 2026 STMicroelectronics.
 * All rights reserved.
 *
 * This software is licensed under terms that can be found in the LICENSE file
 * in the root directory of this software component.
 * If no LICENSE file comes with this software, it is provided AS-IS.
 *
 ******************************************************************************
 */
/* USER CODE END Header */
/* Includes ------------------------------------------------------------------*/
#include "main.h"

/* Private includes ----------------------------------------------------------*/
/* USER CODE BEGIN Includes */

#include "drivers/STS.h"
#include "drivers/WT901B.h"

#include "peripherals/gpio.h"
#include "peripherals/usart.h"

#include "stm32l4xx_hal.h"
#include "stm32l4xx_hal_gpio.h"
#include "utils/data_topic.h"
#include "utils/float3.h"
#include "utils/quaternion.h"
#include "utils/quaternion_dynamics.h"

#include <stdint.h>
#include <stdio.h>
#include <string.h>
#include <math.h>
#include <sys/types.h>

/* USER CODE END Includes */

/* Private typedef -----------------------------------------------------------*/
/* USER CODE BEGIN PTD */

#define PI (3.14159265f)
#define DEG_TO_RAD (PI / 180.0f)
#define RAD_TO_DEG (180.0f / PI)

#define WT901B_FREQUENCY_HZ 100 /* Fréquence de lecture des données du WT901B, à ajuster selon la configuration du capteur */
#define WT901B_PERIOD_S (1.0f / WT901B_FREQUENCY_HZ)

#define ATTITUDE_ELEVATION_GOAL_DEG 72 /* Objectif d'angle d'élévation pour le second étage, à ajuster selon la trajectoire souhaitée */
#define ATTITUDE_AZIMUTH_GOAL_DEG 0 /* Objectif d'angle d'azimut pour le second étage, à ajuster selon la trajectoire souhaitée */

typedef struct rocket_attitude_dynamics_t {
	float3_t v_up_body; /* direction of the rocket’s “up” axis in body frame */
	float initial_elevation_deg; /* initial elevation angle in degrees (0 = horizontal, 90 = vertical) */
	
	quatf_t q; /* attitude quaternion (Body -> Earth) */
	float elevation_deg; /* elevation angle in degrees (0 = horizontal, 90 = vertical) */
	float azimuth_deg; /* azimuth angle in degrees (0 = forward, 90 = right) */
} rocket_attitude_dynamics_t;

typedef enum rocket_stage_t {
	ROCKET_FIRST_STAGE,
	ROCKET_SECOND_STAGE,
} rocket_stage_t;

typedef enum first_stage_initialisation_phase_t {
	FIRST_STAGE_AF_ZERO,
	FIRST_STAGE_WAIT_AF_ZERO,
	FIRST_STAGE_SEPA_ZERO,
	FIRST_STAGE_WAIT_SEPA_ZERO,
	FIRST_STAGE_WAIT_STAGE_ASSEMBLY_CONFIRMATION,
} first_stage_initialisation_phase_t;

typedef enum flight_phase_first_stage_t {
	FIRST_STAGE_INITIALISATION,
	FIRST_STAGE_WAIT_LAUNCH_CONFIRMATION,
	FIRST_STAGE_WAIT_BURN_END,
	FIRST_STAGE_SEPARATION,
	FIRST_STAGE_WAIT_SEPARATION_CONFIRMATION,
	FIRST_STAGE_WAIT_APOGEE_CONFIRMATION,
} flight_phase_first_stage_t;

typedef enum flight_phase_second_stage_t {
	SECOND_STAGE_WAIT_STAGE_ASSEMBLY_CONFIRMATION,
	SECOND_STAGE_WAIT_LAUNCH_CONFIRMATION,
	SECOND_STAGE_WAIT_SEPARATION_CONFIRMATION,
	SECOND_STAGE_WAIT_ATTITUDE_CONFIRMATION,
	SECOND_STAGE_BURN_SECOND_BURN_COMMAND,
	SECOND_STAGE_WAIT_SECOND_BURN_CONFIRMATION,
	SECOND_STAGE_WAIT_APOGEE_CONFIRMATION,
} flight_phase_second_stage_t;

rocket_stage_t current_stage;
first_stage_initialisation_phase_t first_stage_init_phase;
flight_phase_first_stage_t first_stage_flight_phase;
flight_phase_second_stage_t second_stage_flight_phase;

bool is_launch_confirmed;
bool is_separation_confirmed;
bool is_second_burn_confirmed;

uint32_t t0;

uint32_t t_launch; // Time refered to the launch, in seconds since boot

float current_accel_g; // Latest acceleration norm in g, used for state machine decisions
float current_pressure_variation_pa_s; // Latest pressure variation in Pa/s, used for state machine decisions



// Les temps sont en millisecondes par rapport au T0 (lancement)

#define T_ALPHA_BETA_0 5000 /* Fin de propulsion du premier étage */
#define T_ALPHA_BETA_1 5300 /* Séparation du premier étage */
#define T_ALPHA_BETA_2 0 /* Apogée de l'ensemble de la fusée (1er et 2e étage sans séparation) */
#define T_ALPHA_BETA_3 0 /* Atterrisage de l'ensemble de la fusée (1er et 2e étage sans séparation) */

#define T_ALPHA_0 0 /* Apogée du premier étage */
#define T_ALPHA_1 0 /* Atterrissage du premier étage */

#define T_BETA_0 0 /* Allumage du second étage */
#define T_BETA_1 0 /* Confirmation de l'allumage du second étage */
#define T_BETA_2 0 /* Apogée du second étage si séparation et actif */
#define T_BETA_3 0 /* Atterrissage du second étage si séparation et actif */
#define T_BETA_4 0 /* Apogée du second étage si séparation mais passif */
#define T_BETA_5 0 /* Atterrissage du second étage si séparation mais passif */

/* USER CODE END PTD */

/* Private define ------------------------------------------------------------*/
/* USER CODE BEGIN PD */

/* USER CODE END PD */

/* Private macro -------------------------------------------------------------*/
/* USER CODE BEGIN PM */

/* USER CODE END PM */

/* Private variables ---------------------------------------------------------*/

/* USER CODE BEGIN PV */

STS_Servo_t servo1, servo2, servo3, servo4;

rocket_attitude_dynamics_t attitude;
bool is_init;

/* USER CODE END PV */

/* Private function prototypes -----------------------------------------------*/
void SystemClock_Config(void);
/* USER CODE BEGIN PFP */
void on_new_accel_frame(data_sub_t *sub);
void on_new_gyro_frame(data_sub_t *sub);
void on_new_pressure_frame(data_sub_t *sub);

void first_stage_state_machine(void);
void first_stage_initialisation(void);
void second_stage_state_machine(void);
/* USER CODE END PFP */

/* Private user code ---------------------------------------------------------*/
/* USER CODE BEGIN 0 */

/* USER CODE END 0 */

/**
  * @brief  The application entry point.
  * @retval int
  */
int main(void)
{

  /* USER CODE BEGIN 1 */

  /* USER CODE END 1 */

  /* MCU Configuration--------------------------------------------------------*/

  /* Reset of all peripherals, Initializes the Flash interface and the Systick. */
  HAL_Init();

  /* USER CODE BEGIN Init */

  /* USER CODE END Init */

  /* Configure the system clock */
  SystemClock_Config();

  /* USER CODE BEGIN SysInit */

  /* USER CODE END SysInit */

  /* Initialize all configured peripherals */
  MX_GPIO_Init();
  MX_USART1_UART_Init();
  MX_USART2_UART_Init();
  MX_USART3_UART_Init();
  MX_UART4_Init();
  /* USER CODE BEGIN 2 */
	HAL_StatusTypeDef res;

	res = STS_UART_Port_Init(&huart_sts_port1, &huart1);
	if (res != HAL_OK) { Error_Handler(); }
	res = STS_UART_Port_Init(&huart_sts_port2, &huart4);
	if (res != HAL_OK) { Error_Handler(); }

	// res = STS_Servo_Init(&servo1, &huart_sts_port1, 1);
	// // if (res != HAL_OK) { Error_Handler(); }
	// HAL_Delay(1);
	// res = STS_Servo_Init(&servo2, &huart_sts_port1, 2);
	// if (res != HAL_OK) { Error_Handler(); }
	// HAL_Delay(1);
	// res = STS_Servo_Init(&servo3, &huart_sts_port2, 3);
	// if (res != HAL_OK) { Error_Handler(); }
	// HAL_Delay(1);
	// res = STS_Servo_Init(&servo4, &huart_sts_port2, 4);
	// if (res != HAL_OK) { Error_Handler(); }
	// HAL_Delay(1);

	WT901B_status_t wt_res = WT901B_Init(&wt901b, &huart3);
	if (wt_res != WT901B_OK) { Error_Handler(); }

	data_sub_t accel_sub = { 0 };
	data_sub_attach(&accel_sub, &wt901b.data_topic, DATA_ATTACH_FROM_OLDEST);

	data_sub_t gyro_sub = { 0 };
	data_sub_attach(&gyro_sub, &wt901b.data_topic, DATA_ATTACH_FROM_OLDEST);

	data_sub_t pressure_sub = { 0 };
	data_sub_attach(&pressure_sub, &wt901b.data_topic, DATA_ATTACH_FROM_OLDEST);

	// Set up initial attitude surveyor state
	attitude.v_up_body = FLOAT3_UNIT_Y; /* axe “up” de la fusée aligné avec l’axe Y du body */
	attitude.initial_elevation_deg = 0.0f * DEG_TO_RAD; /* fusée initialement à 80° */
	attitude.q = quatf_from_axis_angle(FLOAT3_UNIT_Z, - PI / 2); // rotation initiale pour aligner l’axe forward du repère Terre avec l’axe Y du body
	attitude.q = quatf_mul(attitude.q, quatf_from_axis_angle(FLOAT3_UNIT_X, attitude.initial_elevation_deg)); // rotation initiale d’élévation
	attitude.elevation_deg = 0.0f;
	attitude.azimuth_deg = 0.0f;

	// Set up initial state machine state
	// first_stage_init_phase = FIRST_STAGE_AF_ZERO;
	first_stage_init_phase = FIRST_STAGE_WAIT_SEPA_ZERO;
	first_stage_flight_phase = FIRST_STAGE_INITIALISATION;
	second_stage_flight_phase = SECOND_STAGE_WAIT_STAGE_ASSEMBLY_CONFIRMATION;

	// Set up initial confirmations
	is_launch_confirmed = true;
	// is_launch_confirmed = false;
	is_separation_confirmed = false;
	is_second_burn_confirmed = false;

	// Looking for which stage we are
	if (HAL_GPIO_ReadPin(SET_STAGE_GPIO_Port, SET_STAGE_Pin) == GPIO_PIN_RESET) {
		current_stage = ROCKET_SECOND_STAGE;
		for (size_t i = 0; i < 10; i++) {
			HAL_GPIO_TogglePin(LED_BLUE_GPIO_Port, LED_BLUE_Pin);
			HAL_Delay(500);
		}
	} else {
		current_stage = ROCKET_FIRST_STAGE;
		for (size_t i = 0; i < 10; i++) {
			HAL_GPIO_TogglePin(LED_GREEN_GPIO_Port, LED_GREEN_Pin);
			HAL_Delay(500);
		}
	}


	uint8_t rx_buffer[256] = { 0 };
	t0 = HAL_GetTick();

  /* USER CODE END 2 */

  /* Infinite loop */
  /* USER CODE BEGIN WHILE */
	while (1) {
    /* USER CODE END WHILE */

    /* USER CODE BEGIN 3 */
		
		WT901B_Parse_Frames(&wt901b);
		on_new_accel_frame(&accel_sub);
		on_new_pressure_frame(&pressure_sub);
		on_new_gyro_frame(&gyro_sub);

		// switch (current_stage) {
		// 	case ROCKET_FIRST_STAGE: {
		// 		first_stage_state_machine();
		// 		break;
		// 	}
		// 	case ROCKET_SECOND_STAGE: {
		// 		on_new_gyro_frame(&gyro_sub);
		// 		second_stage_state_machine();
		// 		break;
		// 	}
		// }
		// NEED TO CHANGE THIS TO PRINT ALL USEFULL DATA
		if (HAL_GetTick() - t0 > 10) {
			t0 = HAL_GetTick();
			sprintf((char*)rx_buffer, "Elevation: %.2f deg, Azimuth: %.2f deg\r\n", attitude.elevation_deg, attitude.azimuth_deg);
			HAL_UART_Transmit(&huart2, (uint8_t*)rx_buffer, strlen((char*)rx_buffer), HAL_MAX_DELAY);
		}
	
	}

  /* USER CODE END 3 */
}

/**
  * @brief System Clock Configuration
  * @retval None
  */
void SystemClock_Config(void)
{
  RCC_OscInitTypeDef RCC_OscInitStruct = {0};
  RCC_ClkInitTypeDef RCC_ClkInitStruct = {0};

  /** Configure the main internal regulator output voltage
  */
  if (HAL_PWREx_ControlVoltageScaling(PWR_REGULATOR_VOLTAGE_SCALE1) != HAL_OK)
  {
    Error_Handler();
  }

  /** Initializes the RCC Oscillators according to the specified parameters
  * in the RCC_OscInitTypeDef structure.
  */
  RCC_OscInitStruct.OscillatorType = RCC_OSCILLATORTYPE_HSI;
  RCC_OscInitStruct.HSIState = RCC_HSI_ON;
  RCC_OscInitStruct.HSICalibrationValue = RCC_HSICALIBRATION_DEFAULT;
  RCC_OscInitStruct.PLL.PLLState = RCC_PLL_ON;
  RCC_OscInitStruct.PLL.PLLSource = RCC_PLLSOURCE_HSI;
  RCC_OscInitStruct.PLL.PLLM = 1;
  RCC_OscInitStruct.PLL.PLLN = 10;
  RCC_OscInitStruct.PLL.PLLP = RCC_PLLP_DIV7;
  RCC_OscInitStruct.PLL.PLLQ = RCC_PLLQ_DIV2;
  RCC_OscInitStruct.PLL.PLLR = RCC_PLLR_DIV2;
  if (HAL_RCC_OscConfig(&RCC_OscInitStruct) != HAL_OK)
  {
    Error_Handler();
  }

  /** Initializes the CPU, AHB and APB buses clocks
  */
  RCC_ClkInitStruct.ClockType = RCC_CLOCKTYPE_HCLK|RCC_CLOCKTYPE_SYSCLK
                              |RCC_CLOCKTYPE_PCLK1|RCC_CLOCKTYPE_PCLK2;
  RCC_ClkInitStruct.SYSCLKSource = RCC_SYSCLKSOURCE_PLLCLK;
  RCC_ClkInitStruct.AHBCLKDivider = RCC_SYSCLK_DIV1;
  RCC_ClkInitStruct.APB1CLKDivider = RCC_HCLK_DIV1;
  RCC_ClkInitStruct.APB2CLKDivider = RCC_HCLK_DIV1;

  if (HAL_RCC_ClockConfig(&RCC_ClkInitStruct, FLASH_LATENCY_4) != HAL_OK)
  {
    Error_Handler();
  }
}

/* USER CODE BEGIN 4 */

void on_new_accel_frame(data_sub_t *sub) {
	if (data_sub_num_to_read(sub) > 0) {
		WT901B_Frame_t frame;
		data_sub_read(sub, &frame);
		if (frame.type == WT901B_FRAME_ACCEL) {
			current_accel_g = sqrtf(
				frame.data.accel.ax_g * frame.data.accel.ax_g +
				frame.data.accel.ay_g * frame.data.accel.ay_g +
				frame.data.accel.az_g * frame.data.accel.az_g
			);
		}
	}
}

void on_new_gyro_frame(data_sub_t *sub) {
	if (is_launch_confirmed && data_sub_num_to_read(sub) > 0) {
		WT901B_Frame_t frame;
		data_sub_read(sub, &frame);
		if (frame.type == WT901B_FRAME_GYRO) {
			quatdyn_propagate_ip(
				&attitude.q,
				(float3_t){
					.x = frame.data.gyro.gx_dps * DEG_TO_RAD, /* convert to radians/s */
					.y = frame.data.gyro.gy_dps * DEG_TO_RAD,
					.z = frame.data.gyro.gz_dps * DEG_TO_RAD
				},
				WT901B_PERIOD_S
			);
			float3_t v_body_earth = quatf_rotate_vector(attitude.q, attitude.v_up_body);
			attitude.elevation_deg = 90 - acosf(v_body_earth.z) * RAD_TO_DEG;
			float3_t v_body_earth_proj = float3_sub(v_body_earth, float3_scale(FLOAT3_UNIT_Z, v_body_earth.z));
			attitude.azimuth_deg = atan2f(v_body_earth_proj.y, v_body_earth_proj.x) * RAD_TO_DEG;
		}
	}
}

void on_new_pressure_frame(data_sub_t *sub) {
	if (is_launch_confirmed && data_sub_num_to_read(sub) > 0) {
		WT901B_Frame_t frame;
		data_sub_read(sub, &frame);
		if (frame.type == WT901B_FRAME_PRESSURE) {
			// Assuming we receive pressure in Pa, we can compute the variation if we store the previous pressure
			static float previous_pressure_pa = 0.0f;
			float current_pressure_pa = (float)frame.data.pressure.pressure_pa;
			if (previous_pressure_pa > 0.0f) {
				current_pressure_variation_pa_s = (current_pressure_pa - previous_pressure_pa) / WT901B_PERIOD_S; // variation en Pa/s
			}
			previous_pressure_pa = current_pressure_pa;
		}
	}
}


void first_stage_initialisation(void) {
	switch (first_stage_init_phase) {
		case FIRST_STAGE_AF_ZERO: {
			STS_Servo_SetOperatingMode(&servo1, STS_OP_MODE_SPEED_CONTROL);
			HAL_Delay(1);
			STS_Servo_SetGoalSpeed(&servo1, STS_Servo_GetSpeedInUnits(-200));
			HAL_Delay(1);
			first_stage_init_phase = FIRST_STAGE_WAIT_AF_ZERO;
			break;
		}
		case FIRST_STAGE_WAIT_AF_ZERO: {
			bool in_overload;
			HAL_Delay(1);
			STS_Servo_InOverload(&servo1, &in_overload);
			if (in_overload) {
				HAL_Delay(1);
				STS_Servo_SetGoalSpeed(&servo1, 0);
				HAL_Delay(1);
				STS_Servo_SetOperatingMode(&servo1, STS_OP_MODE_POSITION_CONTROL);
				HAL_Delay(1);
				STS_Servo_PositionCalibration(&servo1, 0);
				HAL_Delay(1);
				STS_Servo_SetGoalPosition(&servo1, STS_Servo_GetPositionInUnits(30));
				HAL_Delay(1);
				first_stage_init_phase = FIRST_STAGE_SEPA_ZERO;
			}
			break;
		}
		case FIRST_STAGE_SEPA_ZERO: {
			STS_Servo_SetOperatingMode(&servo3, STS_OP_MODE_SPEED_CONTROL);
			HAL_Delay(1);
			STS_Servo_SetGoalSpeed(&servo3, STS_Servo_GetSpeedInUnits(-200));
			HAL_Delay(1);
			first_stage_init_phase = FIRST_STAGE_WAIT_SEPA_ZERO;
			break;
		}
		case FIRST_STAGE_WAIT_SEPA_ZERO: {
			bool in_overload;
			HAL_Delay(1);
			STS_Servo_InOverload(&servo3, &in_overload);
			if (true) {
				HAL_Delay(1);
				STS_Servo_SetGoalSpeed(&servo3, 0);
				HAL_Delay(1);
				STS_Servo_SetOperatingMode(&servo3, STS_OP_MODE_POSITION_CONTROL);
				HAL_Delay(1);
				STS_Servo_PositionCalibration(&servo3, 0);
				HAL_Delay(1);
				STS_Servo_SetGoalPosition(&servo3, STS_Servo_GetPositionInUnits(120));
				HAL_Delay(1);
				first_stage_init_phase = FIRST_STAGE_WAIT_STAGE_ASSEMBLY_CONFIRMATION;
			}
			break;
		}
		case FIRST_STAGE_WAIT_STAGE_ASSEMBLY_CONFIRMATION: {
			if (HAL_GPIO_ReadPin(SEPA_GPIO_Port, SEPA_Pin) == GPIO_PIN_RESET) {
				HAL_GPIO_TogglePin(LED_RED_GPIO_Port, LED_RED_Pin);
				HAL_Delay(100);
				HAL_GPIO_TogglePin(LED_RED_GPIO_Port, LED_RED_Pin);
				first_stage_flight_phase = FIRST_STAGE_WAIT_LAUNCH_CONFIRMATION;
			}
			break;
		}
	}
}

void first_stage_state_machine(void) {
	// FIRST_STAGE_INITIALISATION,
	// FIRST_STAGE_WAIT_LAUNCH_CONFIRMATION,
	// FIRST_STAGE_WAIT_BURN_END,
	// FIRST_STAGE_SEPARATION,
	// FIRST_STAGE_WAIT_SEPARATION_CONFIRMATION,
	// FIRST_STAGE_WAIT_APOGEE_CONFIRMATION,

	switch (first_stage_flight_phase) {
		case FIRST_STAGE_INITIALISATION: {
			first_stage_initialisation();
			break;
		}
		case FIRST_STAGE_WAIT_LAUNCH_CONFIRMATION: {
			// Wait for launch confirmation
			if (HAL_GPIO_ReadPin(JACK_GPIO_Port, JACK_Pin) == GPIO_PIN_SET) {
				t_launch = HAL_GetTick();
				HAL_GPIO_TogglePin(LED_RED_GPIO_Port, LED_RED_Pin);
				HAL_Delay(100);
				HAL_GPIO_TogglePin(LED_RED_GPIO_Port, LED_RED_Pin);
				is_launch_confirmed = true;
				first_stage_flight_phase = FIRST_STAGE_WAIT_BURN_END;
			}
			break;
		}
		case FIRST_STAGE_WAIT_BURN_END: {
			if (HAL_GetTick() - t_launch > T_ALPHA_BETA_0) {
				HAL_GPIO_TogglePin(LED_RED_GPIO_Port, LED_RED_Pin);
				first_stage_flight_phase = FIRST_STAGE_SEPARATION;
			}
			break;
		}
		case FIRST_STAGE_SEPARATION: {
			if (HAL_GetTick() - t_launch > T_ALPHA_BETA_1) {
				HAL_GPIO_TogglePin(LED_RED_GPIO_Port, LED_RED_Pin);
				STS_Servo_SetGoalPosition(&servo1, STS_Servo_GetPositionInUnits(120));
				HAL_Delay(1);
				STS_Servo_SetGoalPosition(&servo3, STS_Servo_GetPositionInUnits(120));
				HAL_Delay(1);
				first_stage_flight_phase = FIRST_STAGE_WAIT_SEPARATION_CONFIRMATION;
			}
			break;
		}
		case FIRST_STAGE_WAIT_SEPARATION_CONFIRMATION: {
			if (HAL_GPIO_ReadPin(SEPA_GPIO_Port, SEPA_Pin) == GPIO_PIN_RESET) {
				is_separation_confirmed = true;
				first_stage_flight_phase = FIRST_STAGE_WAIT_APOGEE_CONFIRMATION;
			} else if (HAL_GetTick() - t_launch > T_ALPHA_BETA_1 + 5000) { // 5 seconds after expected separation time, if no confirmation received
				is_separation_confirmed = false;
				first_stage_flight_phase = FIRST_STAGE_WAIT_APOGEE_CONFIRMATION; 
			}
			break;
		}
		case FIRST_STAGE_WAIT_APOGEE_CONFIRMATION: {
			uint32_t t_apogee_estimated = is_separation_confirmed ? T_ALPHA_0 : T_ALPHA_BETA_2;
			if (HAL_GetTick() - t_launch < t_apogee_estimated) {
				// Wait to reach estimated apogee time
			} else if (current_pressure_variation_pa_s < 5.0f) { // If we are near apogee (low pressure variation), confirm apogee
				// Apogee confirmed
				// ...
			} else if (HAL_GetTick() - t_launch > t_apogee_estimated + 3000) { // 3 seconds after estimated apogee time, if no confirmation received
				// Apogee not confirmed, but we can consider it passed
				// ...
			}
			break;
		}
	}
}

void second_stage_state_machine(void) {
	// SECOND_STAGE_WAIT_STAGE_ASSEMBLY_CONFIRMATION,
	// SECOND_STAGE_WAIT_LAUNCH_CONFIRMATION,
	// SECOND_STAGE_WAIT_SEPARATION_CONFIRMATION,
	// SECOND_STAGE_WAIT_ATTITUDE_CONFIRMATION,
	// SECOND_STAGE_BURN_SECOND_BURN_COMMAND,
	// SECOND_STAGE_WAIT_SECOND_BURN_CONFIRMATION,
	// SECOND_STAGE_WAIT_APOGEE_CONFIRMATION,

	switch (second_stage_flight_phase) {
		case SECOND_STAGE_WAIT_STAGE_ASSEMBLY_CONFIRMATION: {
			if (false) {
				second_stage_flight_phase = SECOND_STAGE_WAIT_LAUNCH_CONFIRMATION;
			}
			break;
		}
		case SECOND_STAGE_WAIT_LAUNCH_CONFIRMATION: {
			if (false) {
				is_launch_confirmed = true;
				second_stage_flight_phase = SECOND_STAGE_WAIT_SEPARATION_CONFIRMATION;
			}
			break;
		}
		case SECOND_STAGE_WAIT_SEPARATION_CONFIRMATION: {
			if (false) {
				is_separation_confirmed = true;
				second_stage_flight_phase = SECOND_STAGE_WAIT_ATTITUDE_CONFIRMATION;
			} else if (HAL_GetTick() - t_launch > T_BETA_0 + 5000) { // 5 seconds after expected separation time, if no confirmation received
				is_separation_confirmed = false;
				second_stage_flight_phase = SECOND_STAGE_WAIT_ATTITUDE_CONFIRMATION; 
			}
			break;
		}
		case SECOND_STAGE_WAIT_ATTITUDE_CONFIRMATION: {
			if (fabsf(attitude.elevation_deg - ATTITUDE_ELEVATION_GOAL_DEG	) < 10.0f) {
				if (fabsf(attitude.azimuth_deg - ATTITUDE_AZIMUTH_GOAL_DEG) < 45.0f) {
					// Attitude confirmed
					second_stage_flight_phase = SECOND_STAGE_BURN_SECOND_BURN_COMMAND;
				}
			} else if (HAL_GetTick() - t_launch > T_BETA_1 + 3000) { // 3 seconds after expected attitude confirmation time, if no confirmation received
				// Attitude not confirmed, but we can still command the burn
				second_stage_flight_phase = SECOND_STAGE_WAIT_APOGEE_CONFIRMATION; 
			}
			break;
		}
		case SECOND_STAGE_BURN_SECOND_BURN_COMMAND: {
			// Command second burn (e.g., by sending a signal to the engine)
			// ...
			second_stage_flight_phase = SECOND_STAGE_WAIT_SECOND_BURN_CONFIRMATION;
			break;
		}
		case SECOND_STAGE_WAIT_SECOND_BURN_CONFIRMATION: {
			if (HAL_GetTick() - t_launch < T_BETA_1) {
				// Wait to reach expected burn confirmation time
			} else if (current_accel_g	> 1.5f) { // If we detect a significant acceleration increase, confirm second burn
				is_second_burn_confirmed = true;
				second_stage_flight_phase = SECOND_STAGE_WAIT_APOGEE_CONFIRMATION;
			} else if (HAL_GetTick() - t_launch > T_BETA_1 + 100) { // 0.1 seconds after expected burn confirmation time, if no confirmation received
				is_second_burn_confirmed = false;
				second_stage_flight_phase = SECOND_STAGE_WAIT_APOGEE_CONFIRMATION; 
			}
			break;
		}
		case SECOND_STAGE_WAIT_APOGEE_CONFIRMATION: {
			uint32_t t_apogee_estimated = is_second_burn_confirmed ? T_BETA_2 : T_BETA_4;
			if (HAL_GetTick() - t_launch < t_apogee_estimated) {
				// Wait to reach estimated apogee time
			} else if (current_pressure_variation_pa_s < 5.0f) { // If we are near apogee (low pressure variation), confirm apogee
				// Apogee confirmed
				// ...
			} else if (HAL_GetTick() - t_launch > t_apogee_estimated + 3000) { // 3 seconds after estimated apogee time, if no confirmation received
				// Apogee not confirmed, but we can consider it passed
				// ...
			}
			break;
		}
	}
}

/* USER CODE END 4 */

/**
  * @brief  This function is executed in case of error occurrence.
  * @retval None
  */
void Error_Handler(void)
{
  /* USER CODE BEGIN Error_Handler_Debug */
	/* User can add his own implementation to report the HAL error return state */
	__disable_irq();
	while (1) {
	}
  /* USER CODE END Error_Handler_Debug */
}
#ifdef USE_FULL_ASSERT
/**
  * @brief  Reports the name of the source file and the source line number
  *         where the assert_param error has occurred.
  * @param  file: pointer to the source file name
  * @param  line: assert_param error line source number
  * @retval None
  */
void assert_failed(uint8_t *file, uint32_t line)
{
  /* USER CODE BEGIN 6 */
  /* User can add his own implementation to report the file name and line
     number, ex: printf("Wrong parameters value: file %s on line %d\r\n", file,
     line) */
  /* USER CODE END 6 */
}
#endif /* USE_FULL_ASSERT */
