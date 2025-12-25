/* USER CODE BEGIN Header */
/**
  ******************************************************************************
  * @file           : main.c
  * @brief          : Main program body
  ******************************************************************************
  * @attention
  *
  * Copyright (c) 2025 STMicroelectronics.
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
#include "stm32l4xx_hal.h"
#include "usart.h"
#include "gpio.h"

/* Private includes ----------------------------------------------------------*/
/* USER CODE BEGIN Includes */
#include "STS.h"
#include "WT901B.h"
#include "tools.h"
#include <stdint.h>
#include <string.h>
#include <stdio.h>
/* USER CODE END Includes */

/* Private typedef -----------------------------------------------------------*/
/* USER CODE BEGIN PTD */

/* USER CODE END PTD */

/* Private define ------------------------------------------------------------*/
/* USER CODE BEGIN PD */

/* USER CODE END PD */

/* Private macro -------------------------------------------------------------*/
/* USER CODE BEGIN PM */

/* USER CODE END PM */

/* Private variables ---------------------------------------------------------*/

/* USER CODE BEGIN PV */

/* USER CODE END PV */

/* Private function prototypes -----------------------------------------------*/
void SystemClock_Config(void);
/* USER CODE BEGIN PFP */

/* USER CODE END PFP */

/* Private user code ---------------------------------------------------------*/
/* USER CODE BEGIN 0 */
int _write(int file, char *ptr, int len) {
	int DataIdx;
	for (DataIdx = 0; DataIdx < len; DataIdx++) {
		ITM_SendChar(*ptr++);
	}
	return len;
}


uint8_t compute_checksum(uint8_t *data, uint8_t length) {
  uint8_t sum = 0;
  for (uint8_t i = 0; i < length; i++) {
    sum += data[i];
  }
  return ~sum;
}


typedef enum AeroBrakeState_t {
  AERO_BRAKE_STATE_WAIT,
  AERO_BRAKE_STATE_START,
  AERO_BRAKE_STATE_CALIBRATION_0_START,
  AERO_BRAKE_STATE_CALIBRATION_0_WAIT_0,
  AERO_BRAKE_STATE_CALIBRATION_0_WAIT_1,
  AERO_BRAKE_STATE_CALIBRATION_1_START,
  AERO_BRAKE_STATE_CALIBRATION_1_WAIT_0,
  AERO_BRAKE_STATE_CALIBRATION_1_WAIT_1,
  AERO_BRAKE_STATE_RETRACT_START,
  AERO_BRAKE_STATE_RETRACT_WAIT,
  AERO_BRAKE_STATE_DEPLOY_START,
  AERO_BRAKE_STATE_DEPLOY_WAIT,
} AeroBrakeState_t;



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
  MX_USART2_UART_Init();
  MX_USART1_UART_Init();
  MX_USART3_UART_Init();
  /* USER CODE BEGIN 2 */

  HAL_StatusTypeDef res;

  // WT901B_status_t wt_res = WT901B_Init(&wt901b, &huart1);

  res = STS_UART_Port_Init(&huart_sts_port1, &huart3);
  if (res != HAL_OK) {
      // Initialization failed
      Error_Handler();
  }

  STS_Servo_t servo1;
  res = STS_Servo_Init(&servo1, &huart_sts_port1, 1);
  if (res != HAL_OK) {
      // Initialization failed
      Error_Handler();
  }

  STS_Servo_Current_raw_t current_raw_status;
  STS_Servo_Current_t current_status;
  uint8_t is_moving = 0;
  bool in_overload = 0;

  uint8_t usb_buff[768];
  char position_str[10];
  char speed_str[10];
  char load_str[10];
  char voltage_str[10];
  char temperature_str[10];
  char current_str[10];

  
  AeroBrakeState_t current_state = AERO_BRAKE_STATE_START;
  AeroBrakeState_t next_state = AERO_BRAKE_STATE_START;
  uint32_t t0;
  uint32_t t1 = HAL_GetTick();

  // HAL_Delay(1);
  // res = STS_Servo_SetOperatingMode(&servo1, STS_OP_MODE_SPEED_CONTROL);
  // HAL_Delay(1);
  // res = STS_Servo_SetGoalSpeed(&servo1, STS_Servo_GetSpeedInUnits(-200));
  // // STS_Servo_SetGoalPosition(&servo1, 0);
  // HAL_Delay(10);


  /* USER CODE END 2 */

  /* Infinite loop */
  /* USER CODE BEGIN WHILE */
	while (1) {
  
    // Get current status
    res = STS_Servo_GetCurrentStatus(&servo1, &current_raw_status);
    HAL_Delay(1);
    res = STS_Servo_IsMoving(&servo1, &is_moving);
    HAL_Delay(1);
    res = STS_Servo_InOverload(&servo1, &in_overload);
    HAL_Delay(1);
  
    STS_Servo_raw_to_physical(&current_raw_status, &current_status);

    // Print status over USB UART
    if ((HAL_GetTick() - t1 >= 10) && current_state != AERO_BRAKE_STATE_WAIT && current_state != AERO_BRAKE_STATE_START) {
      float_format(position_str   , current_status.position   , 4, 10);
      float_format(speed_str      , current_status.speed      , 4, 10);
      float_format(load_str       , current_status.load       , 4, 10);
      float_format(voltage_str    , current_status.voltage    , 2,  6);
      float_format(temperature_str, current_status.temperature, 2,  6);
      float_format(current_str    , current_status.current    , 2,  6);

      sprintf((char*)usb_buff, "Move: %1d, Overload: %1d, Pos: %10s deg, Speed: %10s RPM, Load: %10s, Volt: %6s V, Temp: %6s C, Curr: %6s mA\r\n",
              is_moving, in_overload, position_str, speed_str, load_str, voltage_str, temperature_str, current_str);
      HAL_UART_Transmit(&huart2, usb_buff, strlen((char*)usb_buff), HAL_MAX_DELAY);

      t1 = HAL_GetTick();
    }

    // State machine
    switch (current_state) {

    case AERO_BRAKE_STATE_WAIT: {
      // Toggle LED to indicate waiting
      if (HAL_GetTick() - t0 > 500) {
        HAL_GPIO_TogglePin(LD2_GPIO_Port, LD2_Pin);
        t0 = HAL_GetTick();
      }
      // Waiting for button press to proceed
      if (HAL_GPIO_ReadPin(B1_GPIO_Port, B1_Pin) == GPIO_PIN_RESET) {
        HAL_GPIO_WritePin(LD2_GPIO_Port, LD2_Pin, GPIO_PIN_SET); // Turn on LED
        current_state = next_state;
      }
      break; }

    case AERO_BRAKE_STATE_START: {
      // Print start message
      sprintf((char*)usb_buff, "================ Aero-brake test code started! ================\r\n" \
              "The test is made of several steps:\r\n" \
              "1. Calibration of the servo at 0 degrees\r\n" \
              "2. Calibration of the servo at the maximum angle\r\n" \
              "3. Retraction of the aero-brake\r\n" \
              "4. Deployment of the aero-brake\r\n" \
              "After step 4, the test go back to step 3 until the board is reset.\r\n" \
              "The led LD2 blinks while waiting for user input to start each step and stay on when proceed.\r\n" \
              "While running, the servo status is printed over USB UART.\r\n" \
              "Press the user button to start the test...\r\n");
      HAL_UART_Transmit(&huart2, usb_buff, strlen((char*)usb_buff), HAL_MAX_DELAY);
      next_state = AERO_BRAKE_STATE_CALIBRATION_0_START;
      current_state = AERO_BRAKE_STATE_WAIT;
      break; }

    case AERO_BRAKE_STATE_CALIBRATION_0_START: {
      // Start calibration at 0 degrees
      res = STS_Servo_SetOperatingMode(&servo1, STS_OP_MODE_SPEED_CONTROL);
      HAL_Delay(1);
      res = STS_Servo_SetGoalSpeed(&servo1, STS_Servo_GetSpeedInUnits(-200));
      HAL_Delay(1);
      t0 = HAL_GetTick();
      current_state = AERO_BRAKE_STATE_CALIBRATION_0_WAIT_0;
      break; }
    case AERO_BRAKE_STATE_CALIBRATION_0_WAIT_0: {
      // Wait until servo stops moving and overload is detected
      if ((HAL_GetTick() - t0 >= 100) && !is_moving && in_overload) {
        res = STS_Servo_SetGoalSpeed(&servo1, 200);
        HAL_Delay(1);
        current_state = AERO_BRAKE_STATE_CALIBRATION_0_WAIT_1;
      }
      break; }
    case AERO_BRAKE_STATE_CALIBRATION_0_WAIT_1: {
      // Wait until servo is not in overload anymore
      if (!in_overload) {
        res = STS_Servo_SetGoalSpeed(&servo1, 0);
        HAL_Delay(1);
        res = STS_Servo_PositionCalibration(&servo1, 0);
        HAL_Delay(1);
        sprintf((char*)usb_buff, "Calibration at 0 degrees done.\r\nPress the user button to proceed to next step...\r\n");
        HAL_UART_Transmit(&huart2, usb_buff, strlen((char*)usb_buff), HAL_MAX_DELAY);
        next_state = AERO_BRAKE_STATE_CALIBRATION_1_START;
        current_state = AERO_BRAKE_STATE_WAIT;
      }
      break; }

    case AERO_BRAKE_STATE_CALIBRATION_1_START: {
      // Start calibration at maximum angle
      res = STS_Servo_SetOperatingMode(&servo1, STS_OP_MODE_SPEED_CONTROL);
      HAL_Delay(1);
      res = STS_Servo_SetGoalSpeed(&servo1, STS_Servo_GetSpeedInUnits(200));
      HAL_Delay(1);
      t0 = HAL_GetTick();
      current_state = AERO_BRAKE_STATE_CALIBRATION_1_WAIT_0;
      break; }
    case AERO_BRAKE_STATE_CALIBRATION_1_WAIT_0: {
      // Wait until servo stops moving and overload is detected
      if ((HAL_GetTick() - t0 >= 100) && !is_moving && in_overload) {
        res = STS_Servo_SetGoalSpeed(&servo1, -200);
        HAL_Delay(1);
        current_state = AERO_BRAKE_STATE_CALIBRATION_1_WAIT_1;
      }
      break; }
    case AERO_BRAKE_STATE_CALIBRATION_1_WAIT_1: {
      // Wait until servo is not in overload anymore
      if (!in_overload) {
        res = STS_Servo_SetGoalSpeed(&servo1, 0);
        HAL_Delay(1);
        res = STS_Servo_SetOperatingMode(&servo1, STS_OP_MODE_POSITION_CONTROL);
        HAL_Delay(1);
        float_format(position_str, current_status.position, 4, 10);
        sprintf((char*)usb_buff, "Calibration at maximum angle done. Measured max angle: %10s degrees.\r\nPress the user button to proceed to next step...\r\n",
                position_str);
        HAL_UART_Transmit(&huart2, usb_buff, strlen((char*)usb_buff), HAL_MAX_DELAY);
        next_state = AERO_BRAKE_STATE_RETRACT_START;
        current_state = AERO_BRAKE_STATE_WAIT;
      }
      break; }

    case AERO_BRAKE_STATE_RETRACT_START: {
      // Retract aero-brake
      res = STS_Servo_SetGoalPosition(&servo1, 0);
      HAL_Delay(1);
      t0 = HAL_GetTick();
      current_state = AERO_BRAKE_STATE_RETRACT_WAIT;
      break; }
    case AERO_BRAKE_STATE_RETRACT_WAIT: {
      // Wait until retraction is done
      if ((HAL_GetTick() - t0 >= 100) && !is_moving) {
        sprintf((char*)usb_buff, "Aero-brake retracted.\r\nPress the user button to deploy the aero-brake...\r\n");
        HAL_UART_Transmit(&huart2, usb_buff, strlen((char*)usb_buff), HAL_MAX_DELAY);
        next_state = AERO_BRAKE_STATE_DEPLOY_START;
        current_state = AERO_BRAKE_STATE_WAIT;
      }
      break; }

    case AERO_BRAKE_STATE_DEPLOY_START: {
      // Deploy aero-brake
      res = STS_Servo_SetGoalPosition(&servo1, STS_Servo_GetPositionInUnits(120));
      HAL_Delay(1);
      t0 = HAL_GetTick();
      current_state = AERO_BRAKE_STATE_DEPLOY_WAIT;
      break; }
    case AERO_BRAKE_STATE_DEPLOY_WAIT: {
      // Wait until deployment is done
      if ((HAL_GetTick() - t0 >= 100) && !is_moving) {
        sprintf((char*)usb_buff, "Aero-brake deployed.\r\nPress the user button to retract the aero-brake...\r\n");
        HAL_UART_Transmit(&huart2, usb_buff, strlen((char*)usb_buff), HAL_MAX_DELAY);
        next_state = AERO_BRAKE_STATE_RETRACT_START;
        current_state = AERO_BRAKE_STATE_WAIT;
      }
      break; }
    }

    // HAL_Delay(1);
    // if (in_overload) {
    //   res = STS_Servo_SetGoalSpeed(&servo1, 0);
    //   HAL_Delay(1);
    //   res = STS_Servo_SetOperatingMode(&servo1, STS_OP_MODE_POSITION_CONTROL);
    //   HAL_Delay(1);
    //   res = STS_Servo_PositionCalibration(&servo1, 0);
    //   HAL_Delay(1);
    //   STS_Servo_SetGoalPosition(&servo1, 0);
    //   HAL_Delay(1);
    //   STS_Servo_SetGoalPosition(&servo1, STS_Servo_GetPositionInUnits(120));
    // }
    // HAL_Delay(10);

    /* USER CODE END WHILE */

    /* USER CODE BEGIN 3 */
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
  while (1)
  {
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
  /* User can add his own implementation to report the file name and line number,
     ex: printf("Wrong parameters value: file %s on line %d\r\n", file, line) */
  /* USER CODE END 6 */
}
#endif /* USE_FULL_ASSERT */
