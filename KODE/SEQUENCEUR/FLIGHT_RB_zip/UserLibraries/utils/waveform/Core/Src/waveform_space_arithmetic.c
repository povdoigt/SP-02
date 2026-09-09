#include "waveform_space_arithmetic.h"
#include <stddef.h>

float3_t waveform_space_add_const(float t, const void *ctx) {
    const waveform_space_add_const_t *contx = (const waveform_space_add_const_t *)ctx;
    if (contx == NULL) {
        return FLOAT3_ZERO;
    }

    float3_t add_val = contx->wave_function_add(t, contx->ctx_add);
    return float3_add(add_val, contx->offset);
}

float3_t waveform_space_add(float t, const void *ctx) {
    const waveform_space_add_t *contx = (const waveform_space_add_t *)ctx;
    if (contx == NULL) {
        return FLOAT3_ZERO;
    }

    float3_t add_val_1 = contx->wave_function_add_1(t, contx->ctx_add_1);
    float3_t add_val_2 = contx->wave_function_add_2(t, contx->ctx_add_2);
    return float3_add(add_val_1, add_val_2);
}

float3_t waveform_space_mult_const(float t, const void *ctx) {
    const waveform_space_mult_const_t *contx = (const waveform_space_mult_const_t *)ctx;
    if (contx == NULL) {
        return FLOAT3_ZERO;
    }

    float intensity = contx->wave_function_mult(t, contx->ctx_mult);
    return float3_scale(contx->factor, intensity);
}