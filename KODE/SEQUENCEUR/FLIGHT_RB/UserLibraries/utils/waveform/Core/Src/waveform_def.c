#include "waveform_def.h"



static inline void __Waveform_Init(waveform_generic_t *waveform, waveform_kind_t kind, wave_func_u wave_function,
                                   const void *ctx, uint32_t duration, bool periodic) {
    
    waveform->common.kind = kind;
    waveform->wave_function = wave_function;
    waveform->common.ctx = ctx;
    waveform->common.time = duration;
    waveform->common.periodic = periodic;
    Waveform_Restart(waveform);
}

void Waveform_Init_Scalar(waveform_scalar_t *waveform, wave_func_scalar_t wave_function,
                          const void *ctx, uint32_t duration, bool periodic) {
    wave_func_u func_union = { .scalar_wave_function = wave_function };
    __Waveform_Init((waveform_generic_t *)waveform, WAVEFORM_SCALAR, func_union, ctx, duration, periodic);
}

void Waveform_Init_Space(waveform_space_t *waveform, wave_func_space_t wave_function,
                         const void *ctx, uint32_t duration, bool periodic) {
    wave_func_u func_union = { .space_wave_function = wave_function };
    __Waveform_Init((waveform_generic_t *)waveform, WAVEFORM_SPACE, func_union, ctx, duration, periodic);
}

void Waveform_Restart(waveform_generic_t *waveform) {
    waveform->common.t0 = 0;
    waveform->common.active = true;
    waveform->common.paused = false;
    waveform->common.pause_t0 = 0;
}

bool Waveform_IsActive(waveform_generic_t *waveform) {
    return waveform->common.active;
}

void Waveform_Pause(waveform_generic_t *waveform, bool pause, uint32_t current_time) {
    if (pause == waveform->common.paused) {
        return; // idempotent
    }
    if (pause) {
        waveform->common.paused = true;
        waveform->common.pause_t0 = current_time;
    } else {
        // Décale t0 du temps total passé en pause, pour que la phase
        // reprenne exactement où elle s'était arrêtée (pas de saut).
        // Exception : si la waveform n'a jamais été jouée (t0 encore au
        // sentinel 0, cf. Waveform_Play), il n'y a rien à décaler ; le
        // lazy-init au prochain Play fera démarrer la phase à zéro,
        // ce qui est le comportement voulu pour un event jamais visible.
        if (waveform->common.t0 != 0) {
            uint32_t paused_duration = current_time - waveform->common.pause_t0;
            waveform->common.t0 += paused_duration;
        }
        waveform->common.paused = false;
    }
}

bool Waveform_IsPaused(waveform_generic_t *waveform) {
    return waveform->common.paused;
}

static inline float3_t Waveform_Play(waveform_generic_t *waveform, uint32_t current_time) {
    if (!Waveform_IsActive(waveform)) {
        return FLOAT3_ZERO;
    }
    if (waveform->common.paused) {
        return FLOAT3_ZERO;
    }
    if (waveform->common.t0 == 0) {
        waveform->common.t0 = current_time;
    }
    uint32_t elapsed_time = current_time - waveform->common.t0;
    if (elapsed_time >= waveform->common.time) {
        if (waveform->common.periodic) {
            waveform->common.t0 = current_time;
        } else {
            waveform->common.active = false;
            return FLOAT3_ZERO;
        }
    }
    float t = (float)elapsed_time / (float)waveform->common.time; // Phase in [0, 1]
    if (waveform->common.kind == WAVEFORM_SCALAR) {
        float intensity = waveform->wave_function.scalar_wave_function(t, waveform->common.ctx);
        return (float3_t){ .x = intensity, .y = 0, .z = 0 };
    } else { // WAVEFORM_SPACE
        return waveform->wave_function.space_wave_function(t, waveform->common.ctx);
    }
}

float Waveform_Play_Scalar(waveform_scalar_t *waveform, uint32_t current_time) {
    if (waveform->common.kind != WAVEFORM_SCALAR) {
        return 0.0f; // or some error value
    }
    float3_t result = Waveform_Play((waveform_generic_t *)waveform, current_time);
    return result.x; // The scalar intensity is stored in the x component
}

float3_t Waveform_Play_Space(waveform_space_t *waveform, uint32_t current_time) {
    if (waveform->common.kind != WAVEFORM_SPACE) {
        return FLOAT3_ZERO;
    }
    return Waveform_Play((waveform_generic_t *)waveform, current_time);
}
