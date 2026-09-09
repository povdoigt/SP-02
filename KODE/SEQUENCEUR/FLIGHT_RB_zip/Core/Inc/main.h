/* USER CODE BEGIN Header */
/**
  ******************************************************************************
  * @file           : main.h
  * @brief          : Header for main.c file.
  *                   This file contains the common defines of the application.
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

/* Define to prevent recursive inclusion -------------------------------------*/
#ifndef __MAIN_H
#define __MAIN_H

#ifdef __cplusplus
extern "C" {
#endif

/* Includes ------------------------------------------------------------------*/
#include "stm32f0xx_hal.h"

/* Private includes ----------------------------------------------------------*/
/* USER CODE BEGIN Includes */

/* USER CODE END Includes */

/* Exported types ------------------------------------------------------------*/
/* USER CODE BEGIN ET */

/* USER CODE END ET */

/* Exported constants --------------------------------------------------------*/
/* USER CODE BEGIN EC */

/* USER CODE END EC */

/* Exported macro ------------------------------------------------------------*/
/* USER CODE BEGIN EM */

/* USER CODE END EM */

/* Exported functions prototypes ---------------------------------------------*/
void Error_Handler(void);

/* USER CODE BEGIN EFP */

/* USER CODE END EFP */

/* Private defines -----------------------------------------------------------*/
#define TX_OPTO_N1_Pin GPIO_PIN_0
#define TX_OPTO_N1_GPIO_Port GPIOC
#define TX_OPTO_N2_Pin GPIO_PIN_1
#define TX_OPTO_N2_GPIO_Port GPIOC
#define RX_OPTO_N1_Pin GPIO_PIN_2
#define RX_OPTO_N1_GPIO_Port GPIOC
#define RX_OPTO_N2_Pin GPIO_PIN_3
#define RX_OPTO_N2_GPIO_Port GPIOC
#define TX_OUT_Pin GPIO_PIN_0
#define TX_OUT_GPIO_Port GPIOA
#define RX_OUT_Pin GPIO_PIN_1
#define RX_OUT_GPIO_Port GPIOA
#define DATA_SM2_Pin GPIO_PIN_2
#define DATA_SM2_GPIO_Port GPIOA
#define BUZZER_Pin GPIO_PIN_6
#define BUZZER_GPIO_Port GPIOA
#define PWM_N3_Pin GPIO_PIN_7
#define PWM_N3_GPIO_Port GPIOA
#define DATA_SM1_Pin GPIO_PIN_4
#define DATA_SM1_GPIO_Port GPIOC
#define LED1R_Pin GPIO_PIN_0
#define LED1R_GPIO_Port GPIOB
#define LED1B_Pin GPIO_PIN_1
#define LED1B_GPIO_Port GPIOB
#define LED1G_Pin GPIO_PIN_2
#define LED1G_GPIO_Port GPIOB
#define PRGM0_Pin GPIO_PIN_13
#define PRGM0_GPIO_Port GPIOB
#define PRGM1_Pin GPIO_PIN_14
#define PRGM1_GPIO_Port GPIOB
#define PRGM2_Pin GPIO_PIN_15
#define PRGM2_GPIO_Port GPIOB
#define LED2R_Pin GPIO_PIN_6
#define LED2R_GPIO_Port GPIOC
#define LED2G_Pin GPIO_PIN_7
#define LED2G_GPIO_Port GPIOC
#define LED2B_Pin GPIO_PIN_8
#define LED2B_GPIO_Port GPIOC
#define STAGE1_Pin GPIO_PIN_9
#define STAGE1_GPIO_Port GPIOC
#define STAGE2_Pin GPIO_PIN_8
#define STAGE2_GPIO_Port GPIOA
#define PRGM3_Pin GPIO_PIN_9
#define PRGM3_GPIO_Port GPIOA
#define PRGM_RUN_Pin GPIO_PIN_10
#define PRGM_RUN_GPIO_Port GPIOA
#define IN_TRG_N1_Pin GPIO_PIN_15
#define IN_TRG_N1_GPIO_Port GPIOA
#define IN_TRG_N2_Pin GPIO_PIN_10
#define IN_TRG_N2_GPIO_Port GPIOC
#define IN_TRG_N3_Pin GPIO_PIN_11
#define IN_TRG_N3_GPIO_Port GPIOC
#define IN_TRG_N4_Pin GPIO_PIN_12
#define IN_TRG_N4_GPIO_Port GPIOC
#define OUT_N1_Pin GPIO_PIN_2
#define OUT_N1_GPIO_Port GPIOD
#define OUT_N2_Pin GPIO_PIN_3
#define OUT_N2_GPIO_Port GPIOB
#define OUT_N3_Pin GPIO_PIN_4
#define OUT_N3_GPIO_Port GPIOB
#define RX_WT_Pin GPIO_PIN_6
#define RX_WT_GPIO_Port GPIOB
#define TX_WT_Pin GPIO_PIN_7
#define TX_WT_GPIO_Port GPIOB

/* USER CODE BEGIN Private defines */

/* USER CODE END Private defines */

#ifdef __cplusplus
}
#endif

#endif /* __MAIN_H */
