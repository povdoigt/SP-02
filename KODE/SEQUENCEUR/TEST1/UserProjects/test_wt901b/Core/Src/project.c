#include "project.h"

#include "main.h"
#include "usart.h"

#include "WT901B.h"

#include <string.h>


void setup() {
	WT901B_status_t wt_res = WT901B_Init(&wt901b, &huart3);
	if (wt_res != WT901B_OK) { Error_Handler(); }
}

void loop() {

}
