#include "waveform_scalar_arithmetic.h"
#include <stddef.h>

float waveform_scale_add_const(float t, const void *ctx) {
    const waveform_scale_add_const_t *contx = (const waveform_scale_add_const_t *)ctx;
    if (contx == NULL) {
        return 0.0f;
    }

    return contx->offset + contx->wave_function_add(t, contx->ctx_add);
}

float waveform_scale_add(float t, const void *ctx) {
    const waveform_scale_add_t *contx = (const waveform_scale_add_t *)ctx;
    if (contx == NULL) {
        return 0.0f;
    }

    return contx->wave_function_add_1(t, contx->ctx_add_1) + contx->wave_function_add_2(t, contx->ctx_add_2);
}

float waveform_scale_mult_const(float t, const void *ctx) {
    const waveform_scale_mult_const_t *contx = (const waveform_scale_mult_const_t *)ctx;
    if (contx == NULL) {
        return 0.0f;
    }

    return contx->factor * contx->wave_function_mult(t, contx->ctx_mult);
}

float waveform_scale_mult(float t, const void *ctx) {
    const waveform_scale_mult_t *contx = (const waveform_scale_mult_t *)ctx;
    if (contx == NULL) {
        return 0.0f;
    }

    return contx->wave_function_mult_1(t, contx->ctx_mult_1) * contx->wave_function_mult_2(t, contx->ctx_mult_2);
}
