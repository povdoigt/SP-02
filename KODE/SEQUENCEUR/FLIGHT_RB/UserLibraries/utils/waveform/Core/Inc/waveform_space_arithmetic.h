#ifndef WAVEFORM_SPACE_ARITHMETIC_H
#define WAVEFORM_SPACE_ARITHMETIC_H

#include "waveform_def.h"

/**
 * Add a constant to the space waveform.
 * wave_function_add: the waveform function to add
 * ctx_add: context of the waveform function to add
 * offset: the constant to add
 */
typedef struct waveform_space_add_const_t {
    wave_func_space_t   wave_function_add;
    const void          *ctx_add;
    float3_t             offset;
} waveform_space_add_const_t; 
float3_t waveform_space_add_const(float t, const void *ctx);

/**
 * Add two space waveforms.
 * wave_function_add_1: the first waveform function to add
 * ctx_add_1: context of the first waveform function to add
 * wave_function_add_2: the second waveform function to add
 * ctx_add_2: context of the second waveform function to add
 */
typedef struct waveform_space_add_t {
    wave_func_space_t   wave_function_add_1;
    const void          *ctx_add_1;
    wave_func_space_t   wave_function_add_2;
    const void          *ctx_add_2;
} waveform_space_add_t;
float3_t waveform_space_add(float t, const void *ctx);

/**
 * Multiply the scalar waveform by a 3D constant.
 * wave_function_mult: the waveform function to multiply by
 * ctx_mult: context of the waveform function to multiply by
 * factor: the constant to multiply by
 */
typedef struct waveform_space_mult_const_t {
    wave_func_scalar_t   wave_function_mult;
    const void          *ctx_mult;
    float3_t             factor;
} waveform_space_mult_const_t;
float3_t waveform_space_mult_const(float t, const void *ctx);

#endif /* WAVEFORM_SPACE_ARITHMETIC_H */