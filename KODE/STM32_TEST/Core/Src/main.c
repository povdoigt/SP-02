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
#include "stm32l4xx_hal_uart.h"
#include "usart.h"
#include "gpio.h"

/* Private includes ----------------------------------------------------------*/
/* USER CODE BEGIN Includes */
#include <stdint.h>
#include <string.h>
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
void Wit_UART_Write(uint8_t *p_ucData, uint32_t uiLen);
void Wit_Reg_Update_Cb(uint32_t uiReg, uint32_t uiRegNum);
void Wit_Delayms(uint16_t ucMs);
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

void Wit_UART_Write(uint8_t *p_ucData, uint32_t uiLen) {
  HAL_UART_Transmit(&huart1, p_ucData, uiLen, HAL_MAX_DELAY);
}

void Wit_Reg_Update_Cb(uint32_t uiReg, uint32_t uiRegNum) {
  // Callback function implementation
}

void Wit_Delayms(uint16_t ucMs) {
  HAL_Delay(ucMs);
}

void ftUart_Send(uint8_t *nDat, int nLen) {
  HAL_UART_Transmit(&huart3, nDat, nLen, HAL_MAX_DELAY);
}
int ftUart_Read(uint8_t *nDat, int nLen) {
  uint16_t rx_len;
  HAL_UARTEx_ReceiveToIdle(&huart3, nDat, nLen, &rx_len, HAL_MAX_DELAY);
  return (int)rx_len;
}
void ftBus_Delay(void) {
  HAL_Delay(1);
}




uint8_t compute_checksum(uint8_t *data, uint8_t length) {
  uint8_t sum = 0;
  for (uint8_t i = 0; i < length; i++) {
    sum += data[i];
  }
  return ~sum;
}





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
  // uint8_t tx_data[6] = {
  //   0xFF, 0xFF, 0x01, 0x02, 0x01, 0xFB
  // };
  // uint8_t tx_cmd1[13] = {
  //   0xFF, 0xFF, 0x01, 0x09, 0x03, 0x2A, 0x00, 0x08, 0x00, 0x00, 0xE8, 0x03, 0xD5,
  // };

  // // Compute and set checksum for tx_cmd1
  // tx_cmd1[12] = compute_checksum(&tx_cmd1[2], 10);
  
  // // Change position to 0 (0x0000)
  // tx_cmd1[6] = 0x00;
  // tx_cmd1[7] = 0x00;
  // tx_cmd1[12] = compute_checksum(&tx_cmd1[2], 10);

  // // Change speed to 2000 (0x07D0)
  // tx_cmd1[10] = 0xD0;
  // tx_cmd1[11] = 0x07;
  // tx_cmd1[12] = compute_checksum(&tx_cmd1[2], 10);

  // HAL_HalfDuplex_EnableTransmitter(&huart3);
  // HAL_UART_Transmit(&huart3, tx_cmd1, 13, HAL_MAX_DELAY);
  // HAL_HalfDuplex_EnableReceiver(&huart3);

  // HAL_Delay(2000);

  // __NOP();

  // WitSerialWriteRegister(Wit_UART_Write);
  // WitRegisterCallBack(Wit_Reg_Update_Cb);
  // WitDelayMsRegister(Wit_Delayms);

  // WitInit(WIT_PROTOCOL_NORMAL, 0x50);

  // setEnd(0);
  // WritePosEx(7, 4095, 2400, 50);

  /* USER CODE END 2 */

  /* Infinite loop */
  /* USER CODE BEGIN WHILE */
  uint8_t tx_buff[256] = "Coucou !";
	while (1) {
    // // Go to position 4095 (0x0FFF)
    // tx_cmd1[6] = 0xFF;
    // tx_cmd1[7] = 0x0F;
    // tx_cmd1[12] = compute_checksum(&tx_cmd1[2], 10);
  
    // HAL_HalfDuplex_EnableTransmitter(&huart3);
    // HAL_UART_Transmit(&huart3, tx_cmd1, 13, HAL_MAX_DELAY);
    // HAL_HalfDuplex_EnableReceiver(&huart3);
  
    // HAL_Delay(2000);
  
    // // Go to position 0 (0x0000)
    // tx_cmd1[6] = 0x00;
    // tx_cmd1[7] = 0x00;
    // tx_cmd1[12] = compute_checksum(&tx_cmd1[2], 10);
  
    // HAL_HalfDuplex_EnableTransmitter(&huart3);
    // HAL_UART_Transmit(&huart3, tx_cmd1, 13, HAL_MAX_DELAY);
    // HAL_HalfDuplex_EnableReceiver(&huart3);
  
    // HAL_Delay(2000);

    HAL_UART_Transmit(&huart2, tx_buff, strlen((char *)tx_buff), HAL_MAX_DELAY);
    HAL_Delay(1000);

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
