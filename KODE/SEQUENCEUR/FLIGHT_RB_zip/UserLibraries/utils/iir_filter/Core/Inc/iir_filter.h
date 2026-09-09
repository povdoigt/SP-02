#include "circular_buffer.h"
// #include "stm32f4xx_hal.h"

#ifndef IIR_FILTER_H
#define IIR_FILTER_H

typedef struct iir_filter_t {
    circular_buffer_t x_cb; // Input buffer
    circular_buffer_t y_cb; // Output buffer
    const float *a_storage; // Coefficients for the feedback loop (should contain a_order + 1 elements)
    const float *b_storage; // Coefficients for the feedforward loop (should contain b_order + 1 elements)
} iir_filter_t;

int iir_init(iir_filter_t *filter,
             size_t a_order, size_t b_order,
             const float *a_storage, const float *b_storage,
             float *x_storage, float *y_storage);

float iir_process(iir_filter_t *filter, float input);

#endif // IIR_FILTER_H
