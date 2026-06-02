/* USER CODE BEGIN Header */
/**
  ******************************************************************************
  * @file    usart.h
  * @brief   This file contains all the function prototypes for
  *          the usart.c file
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
#ifndef __USART_H__
#define __USART_H__

#ifdef __cplusplus
extern "C" {
#endif

/* Includes ------------------------------------------------------------------*/
#include "main.h"

/* USER CODE BEGIN Includes */

/* USER CODE END Includes */

extern UART_HandleTypeDef huart1;

extern UART_HandleTypeDef huart2;

extern UART_HandleTypeDef huart3;

extern UART_HandleTypeDef huart4;

/* USER CODE BEGIN Private defines */

/* USER CODE END Private defines */

void MX_USART1_UART_Init(void);
void MX_USART2_UART_Init(void);
void MX_USART3_UART_Init(void);
void MX_USART4_UART_Init(void);

/* USER CODE BEGIN Prototypes */

typedef struct UART_buffer_t {
  uint8_t *rx_buffer;
  size_t rx_length;
} UART_buffer_t;


#ifdef USART1
  extern UART_buffer_t uart_buffer_1;
#endif
#ifdef USART2
  extern UART_buffer_t uart_buffer_2;
#endif
#ifdef USART3
  extern UART_buffer_t uart_buffer_3;
#endif
#ifdef USART4
  extern UART_buffer_t uart_buffer_4;
#endif
#ifdef USART5
  extern UART_buffer_t uart_buffer_5;
#endif

void UART_get_buffer(UART_HandleTypeDef *huart, UART_buffer_t **buffer_obj_ptr);

/* USER CODE END Prototypes */

#ifdef __cplusplus
}
#endif

#endif /* __USART_H__ */

