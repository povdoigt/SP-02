#ifndef WAVEFORM_SCALAR_ARITHMETIC_H
#define WAVEFORM_SCALAR_ARITHMETIC_H

#include "waveform_def.h"

/**
 * Add a constant to the scalar waveform.
 * wave_function_add: the waveform function to add
 * ctx_add: context of the waveform function to add
 * offset: the constant to add
 */
typedef struct waveform_scale_add_const_t {
    wave_func_scalar_t   wave_function_add;
    const void          *ctx_add;
    float                offset;
} waveform_scale_add_const_t; 
float waveform_scale_add_const(float t, const void *ctx);

/**
 * Add two scalar waveforms.
 * wave_function_add_1: the first waveform function to add
 * ctx_add_1: context of the first waveform function to add
 * wave_function_add_2: the second waveform function to add
 * ctx_add_2: context of the second waveform function to add
 */
typedef struct waveform_scale_add_t {
    wave_func_scalar_t   wave_function_add_1;
    const void          *ctx_add_1;
    wave_func_scalar_t   wave_function_add_2;
    const void          *ctx_add_2;
} waveform_scale_add_t;
float waveform_scale_add(float t, const void *ctx);

/**
 * Multiply the scalar waveform by a constant.
 * wave_function_mult: the waveform function to multiply by
 * ctx_mult: context of the waveform function to multiply by
 * factor: the constant to multiply by
 */
typedef struct waveform_scale_mult_const_t {
    wave_func_scalar_t   wave_function_mult;
    const void          *ctx_mult;
    float                factor;
} waveform_scale_mult_const_t;
float waveform_scale_mult_const(float t, const void *ctx);

/**
 * Multiply two scalar waveforms.
 * wave_function_mult_1: the first waveform function to multiply by
 * ctx_mult_1: context of the first waveform function to multiply by
 * wave_function_mult_2: the second waveform function to multiply by
 * ctx_mult_2: context of the second waveform function to multiply by
 */
typedef struct waveform_scale_mult_t {
    wave_func_scalar_t   wave_function_mult_1;
    const void          *ctx_mult_1;
    wave_func_scalar_t   wave_function_mult_2;
    const void          *ctx_mult_2;
} waveform_scale_mult_t;
float waveform_scale_mult(float t, const void *ctx);

#endif /* WAVEFORM_SCALAR_ARITHMETIC_H */