#include "waveform.h"



static inline void __Waveform_Init(waveform_t *waveform, waveform_kind_t kind, wave_func_u wave_function,
                                   const void *ctx, uint32_t duration, bool periodic) {
    
    waveform->kind = kind;
    waveform->wave_function = wave_function;
    waveform->ctx = ctx;
    waveform->time = duration;
    waveform->periodic = periodic;
    Waveform_Restart(waveform);
}

void Waveform_Init_Scalar(waveform_t *waveform, wave_func_scalar_t wave_function,
                          const void *ctx, uint32_t duration, bool periodic) {
    wave_func_u func_union = { .scalar_wave_function = wave_function };
    __Waveform_Init(waveform, WAVEFORM_SCALAR, func_union, ctx, duration, periodic);
}

void Waveform_Init_Space(waveform_t *waveform, wave_func_space_t wave_function,
                         const void *ctx, uint32_t duration, bool periodic) {
    wave_func_u func_union = { .space_wave_function = wave_function };
    __Waveform_Init(waveform, WAVEFORM_SPACE, func_union, ctx, duration, periodic);
}

void Waveform_Restart(waveform_t *waveform) {
    waveform->t0 = 0;
    waveform->active = true;
}

bool Waveform_IsActive(waveform_t *waveform) {
    return waveform->active;
}

float Waveform_Play_1D(waveform_t *waveform, uint32_t current_time) {
    if (!Waveform_IsActive(waveform)) {
        return 0.0f;
    }
    if (waveform->t0 == 0) {
        waveform->t0 = current_time;
    }
    uint32_t elapsed_time = current_time - waveform->t0;
    if (elapsed_time >= waveform->time) {
        if (waveform->periodic) {
            waveform->t0 = current_time;
        } else {
            waveform->active = false;
            return 0.0f;
        }
    }
    float t = (float)elapsed_time / (float)waveform->time; // Phase in [0, 1]
    return waveform->wave_function.scalar_wave_function(t, waveform->ctx);
}

float3_t Waveform_Play_3D(waveform_t *waveform, uint32_t current_time) {
    if (!Waveform_IsActive(waveform)) {
        return FLOAT3_ZERO;
    }
    if (waveform->t0 == 0) {
        waveform->t0 = current_time;
    }
    uint32_t elapsed_time = current_time - waveform->t0;
    if (elapsed_time >= waveform->time) {
        if (waveform->periodic) {
            waveform->t0 = current_time;
        } else {
            waveform->active = false;
            return FLOAT3_ZERO;
        }
    }
    float t = (float)elapsed_time / (float)waveform->time; // Phase in [0, 1]
    return waveform->wave_function.space_wave_function(t, waveform->ctx);
}
