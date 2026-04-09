#include "waveform_built_in.h"

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
    if (ctx == NULL) {
        return 0.0f;
    }

    const waveform_gate_t *gate = (const waveform_gate_t *)ctx;
    return phase_in_interval(t, gate->start, gate->end) ? 1.0f : 0.0f;
}

float waveform_ramp(float t, const void *ctx) {
    if (ctx == NULL) {
        return 0.0f;
    }

    const waveform_ramp_t *ramp = (const waveform_ramp_t *)ctx;

    if (!phase_in_interval(t, ramp->start, ramp->end)) {
        return 0.0f;
    }

    return phase_local(t, ramp->start, ramp->end);
}

float waveform_triangle(float t, const void *ctx) {
    if (ctx == NULL) {
        return 0.0f;
    }

    const waveform_triangle_t *triangle = (const waveform_triangle_t *)ctx;

    if (!phase_in_interval(t, triangle->start, triangle->end)) {
        return 0.0f;
    }

    float u = phase_local(t, triangle->start, triangle->end);

    if (u < 0.5f) {
        return 2.0f * u;
    } else {
        return 2.0f * (1.0f - u);
    }
}

float waveform_sine(float t, const void *ctx) {
    if (ctx == NULL) {
        return 0.0f;
    }

    const waveform_sine_t *sine = (const waveform_sine_t *)ctx;

    if (!phase_in_interval(t, sine->start, sine->end)) {
        return 0.0f;
    }

    float u = phase_local(t, sine->start, sine->end);
    return 0.5f - 0.5f * cosf(2.0f * M_PI * u);
}

float waveform_impulse(float t, const void *ctx) {
    if (ctx == NULL) {
        return 0.0f;
    }

    const waveform_impulse_t *impulse = (const waveform_impulse_t *)ctx;

    if (!phase_in_interval(t, impulse->start, impulse->end)) {
        return 0.0f;
    }

    float u = phase_local(t, impulse->start, impulse->end);
    float k = (impulse->k > 0.0f) ? impulse->k : 1.0f;

    return expf(-k * u);
}