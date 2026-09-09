#include "iir_filter.h"
#include "circular_buffer.h"
#include <string.h>


int iir_init(iir_filter_t *filter,
             size_t a_order, size_t b_order,
             const float *a_storage, const float *b_storage,
             float *x_storage, float *y_storage) {

    // a_storage should have a_order + 1 elements
    // b_storage should have b_order + 1 elements
    // x_storage should have b_order elements
    // y_storage should have a_order elements

    if (!filter) { return -1; } // Null pointer check
    if (a_order > 1 && !a_storage) { return -1; } // Null pointer check for a coefficients
    if (b_order > 0 && !b_storage) { return -1; } // Null pointer check for b coefficients
    if (a_order > 1) {
        if (!y_storage) { return -1; } // Null pointer check for y storage
        cb_init(&filter->y_cb, y_storage, sizeof(float), a_order, CB_OVERWRITE_OLDEST);
    } else {
        a_storage = NULL; // No feedback coefficients
    }
    if (b_order > 0) {
        if (!x_storage) { return -1; } // Null pointer check for x storage
        cb_init(&filter->x_cb, x_storage, sizeof(float), b_order, CB_OVERWRITE_OLDEST);
    } else {
        b_storage = NULL; // No feedforward coefficients
    }

    filter->a_storage = a_storage;
    filter->b_storage = b_storage;

    float zero = 0.0f;
    if (b_storage) {
        for (size_t i = 0; i < filter->x_cb.capacity; i++) {
            cb_push(&filter->x_cb, (uint8_t *)&zero);
        }
    }
    if (a_storage) {
        for (size_t i = 0; i < filter->y_cb.capacity; i++) {
            cb_push(&filter->y_cb, (uint8_t *)&zero);
        }
    }

    return 0; // Success
}

float iir_process(iir_filter_t *filter, float input) {

    // Structure of cb currently
    //                      tail                                     head
    //                       v                                         v
    // x (input) buffer     : [x[0  ], x[1  ], ..., x[B-2], x[B-1]]  (inp)   (B elements)
    // y (output) buffer    : [y[0  ], y[1  ], ..., y[A-2], y[A-1]]  (out)   (A elements)
    // a (feedback) coef    : [a[A-1], a[A-2], ..., a[2  ], a[1  ], a[0  ]]  (A + 1 elements)
    // b (feedforward) coef : [b[B-1], b[B-2], ..., b[2  ], b[1  ], b[0  ]]  (B + 1 elements)

    float output = 0.0f;
    
    if (filter->b_storage) {
        output += filter->b_storage[0] * input; // Start with the current input multiplied by the first feedforward coefficient
        for (size_t i = 1; i < filter->x_cb.capacity + 1; i++) {
            float x_val = *(float*)cb_peek_relative_ptr(&filter->x_cb, filter->x_cb.head, -(int)i);
            output += filter->b_storage[i] * x_val;
        }
    }
    if (filter->a_storage) {
        for (size_t i = 1; i < filter->y_cb.capacity + 1; i++) {
            float y_val = *(float*)cb_peek_relative_ptr(&filter->y_cb, filter->y_cb.head, -(int)i);
            output += filter->a_storage[i] * y_val;
        }
    }
    
    cb_push(&filter->x_cb, &input);
    cb_push(&filter->y_cb, &output);

    return output;
}
