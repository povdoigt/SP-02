/* float3.h
 * 3D vector math library (C99) — API only.
 *
 * Rule:
 *  - op(a, b)      -> returns result
 *  - op_ip(&a, b) -> modifies a in-place
 */

#ifndef FLOAT3_H
#define FLOAT3_H

#ifdef __cplusplus
extern "C" {
#endif

#include <stdbool.h>

/* 3D vector */
typedef struct {
    float x;
    float y;
    float z;
} float3_t;

#define FLOAT3_ZERO (float3_t){ .x = 0.0f, .y = 0.0f, .z = 0.0f }
#define FLOAT3_UNIT_X (float3_t){ .x = 1.0f, .y = 0.0f, .z = 0.0f }
#define FLOAT3_UNIT_Y (float3_t){ .x = 0.0f, .y = 1.0f, .z = 0.0f }
#define FLOAT3_UNIT_Z (float3_t){ .x = 0.0f, .y = 0.0f, .z = 1.0f }

/* ---- Constructors ---- */
float3_t float3_make(float x, float y, float z);

/* ---- Basic ops ---- */
float3_t float3_add(float3_t a, float3_t b);
void     float3_add_ip(float3_t *a, float3_t b);

float3_t float3_sub(float3_t a, float3_t b);
void     float3_sub_ip(float3_t *a, float3_t b);

float3_t float3_scale(float3_t v, float s);
void     float3_scale_ip(float3_t *v, float s);

/* ---- Products ---- */
float    float3_dot(float3_t a, float3_t b);

float3_t float3_cross(float3_t a, float3_t b);
void     float3_cross_ip(float3_t *a, float3_t b);

/* ---- Norms ---- */
float    float3_norm2(float3_t v);
float    float3_norm(float3_t v);

/* ---- Normalization ---- */
float3_t float3_normalized(float3_t v);
bool     float3_normalized_ip(float3_t *v);

/* ---- Utilities ---- */
bool     float3_isfinite(float3_t v);
bool     float3_equal_eps(float3_t a, float3_t b, float eps);

#ifdef __cplusplus
}
#endif

#endif /* FLOAT3_H */
