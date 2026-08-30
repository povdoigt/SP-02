#ifndef WAVEFORM_BUILT_IN_H
#define WAVEFORM_BUILT_IN_H

#include "waveform_def.h"


/**
 * Active interval [start, end) on the normalized phase circle.
 * Handles wrap-around if start > end.
 * start, end in [0, 1)
 */
typedef struct waveform_gate_t {
    float start;
    float end;
} waveform_gate_t;
float waveform_gate(float t, const void *ctx);

/**
 * Linear ramp over the active interval [start, end):
 * - before start / after end: 0
 * - inside interval: local phase u in [0,1), returns u
 */
typedef struct waveform_ramp_t {
    float start;
    float end;
} waveform_ramp_t;
float waveform_ramp(float t, const void *ctx);

/**
 * Triangle over the active interval [start, end):
 * - outside interval: 0
 * - inside interval: local phase u in [0,1), returns triangular envelope
 */
typedef struct waveform_triangle_t {
    float start;
    float end;
} waveform_triangle_t;
float waveform_triangle(float t, const void *ctx);

/**
 * Sine pulse over the active interval [start, end):
 * - outside interval: 0
 * - inside interval: local phase u in [0,1), returns 0.5 - 0.5*cos(2*pi*u)
 */
typedef struct waveform_sine_t {
    float start;
    float end;
} waveform_sine_t;
float waveform_sine(float t, const void *ctx);

/**
 * Exponential impulse over the active interval [start, end):
 * - outside interval: 0
 * - inside interval: local phase u in [0,1), returns exp(-k*u)
 * k > 0
 */
typedef struct waveform_impulse_t {
    float start;
    float end;
    float k;
} waveform_impulse_t;
float waveform_impulse(float t, const void *ctx);

#endif /* WAVEFORM_BUILT_IN_H */