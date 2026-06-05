#include "waveform_built_in.h"
#include "float3.h"

#include <math.h>
#include <stdbool.h>


static float clamp01(float x) {
    if (x < 0.0f) {
        return 0.0f;
    }
    if (x > 1.0f) {
        return 1.0f;
    }
    return x;
}

static bool phase_in_interval(float t, float start, float end) {
    t = fmodf(t, 1.0f);
    if (start >= end) {
        return false;   // empty interval by convention
    }
    return (t >= start) && (t < end);
}

static float phase_local(float t, float start, float end) {
    float len = end - start;

    t = fmodf(t, 1.0f);

    float u = t - start;
    if (u < 0.0f) {
        u += 1.0f;
    }

    return clamp01(u / len);
}

float waveform_gate(float t, const void *ctx) {
    const waveform_gate_t *contx = (const waveform_gate_t *)ctx;
    if (contx == NULL) {
        return 0.0f;
    }

    return phase_in_interval(t, contx->start, contx->end) ? 1.0f : 0.0f;
}

float waveform_ramp(float t, const void *ctx) {
    const waveform_ramp_t *contx = (const waveform_ramp_t *)ctx;
    if (contx == NULL) {
        return 0.0f;
    }


    if (!phase_in_interval(t, contx->start, contx->end)) {
        return 0.0f;
    }

    return phase_local(t, contx->start, contx->end);
}

float waveform_triangle(float t, const void *ctx) {
    const waveform_triangle_t *contx = (const waveform_triangle_t *)ctx;
    if (contx == NULL) {
        return 0.0f;
    }


    if (!phase_in_interval(t, contx->start, contx->end)) {
        return 0.0f;
    }

    float u = phase_local(t, contx->start, contx->end);

    if (u < 0.5f) {
        return 2.0f * u;
    } else {
        return 2.0f * (1.0f - u);
    }
}

float waveform_sine(float t, const void *ctx) {
    const waveform_sine_t *contx = (const waveform_sine_t *)ctx;
    if (contx == NULL) {
        return 0.0f;
    }

    if (!phase_in_interval(t, contx->start, contx->end)) {
        return 0.0f;
    }

    float u = phase_local(t, contx->start, contx->end);
    return 0.5f - 0.5f * cosf(2.0f * M_PI * u);
}

float waveform_impulse(float t, const void *ctx) {
    const waveform_impulse_t *contx = (const waveform_impulse_t *)ctx;
    if (contx == NULL) {
        return 0.0f;
    }

    if (!phase_in_interval(t, contx->start, contx->end)) {
        return 0.0f;
    }

    float u = phase_local(t, contx->start, contx->end);
    float k = (contx->k > 0.0f) ? contx->k : 1.0f;

    return expf(-k * u);
}