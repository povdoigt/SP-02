/* float3.c
 * 3D vector math library (C99) — implementation.
 */

#include "float3.h"
#include <math.h>
#include <stdbool.h>

#define FLOAT3_EPS_NORM (1e-20f)

float3_t float3_make(float x, float y, float z) {
    float3_t v = { x, y, z };
    return v;
}

/* ---- Basic ops ---- */

float3_t float3_add(float3_t a, float3_t b) {
    return float3_make(a.x + b.x, a.y + b.y, a.z + b.z);
}

void float3_add_ip(float3_t *a, float3_t b) {
    if (!a) return;
    a->x += b.x;
    a->y += b.y;
    a->z += b.z;
}

float3_t float3_sub(float3_t a, float3_t b) {
    return float3_make(a.x - b.x, a.y - b.y, a.z - b.z);
}

void float3_sub_ip(float3_t *a, float3_t b) {
    if (!a) return;
    a->x -= b.x;
    a->y -= b.y;
    a->z -= b.z;
}

float3_t float3_scale(float3_t v, float s) {
    return float3_make(v.x * s, v.y * s, v.z * s);
}

void float3_scale_ip(float3_t *v, float s) {
    if (!v) return;
    v->x *= s;
    v->y *= s;
    v->z *= s;
}

/* ---- Products ---- */

float float3_dot(float3_t a, float3_t b) {
    return a.x*b.x + a.y*b.y + a.z*b.z;
}

float3_t float3_cross(float3_t a, float3_t b) {
    return float3_make(
        a.y*b.z - a.z*b.y,
        a.z*b.x - a.x*b.z,
        a.x*b.y - a.y*b.x
    );
}

void float3_cross_ip(float3_t *a, float3_t b) {
    if (!a) return;
    float3_t tmp = *a;
    a->x = tmp.y*b.z - tmp.z*b.y;
    a->y = tmp.z*b.x - tmp.x*b.z;
    a->z = tmp.x*b.y - tmp.y*b.x;
}

/* ---- Norms ---- */

float float3_norm2(float3_t v) {
    return float3_dot(v, v);
}

float float3_norm(float3_t v) {
    return sqrtf(float3_norm2(v));
}

/* ---- Normalization ---- */

float3_t float3_normalized(float3_t v) {
    float n2 = float3_norm2(v);
    if (!isfinite(n2) || n2 < FLOAT3_EPS_NORM) {
        return FLOAT3_ZERO;
    }
    float inv_n = 1.0f / sqrtf(n2);
    return float3_scale(v, inv_n);
}

bool float3_normalized_ip(float3_t *v) {
    if (!v) return false;
    float n2 = float3_norm2(*v);
    if (!isfinite(n2) || n2 < FLOAT3_EPS_NORM) {
        *v = FLOAT3_ZERO;
        return false;
    }
    float inv_n = 1.0f / sqrtf(n2);
    float3_scale_ip(v, inv_n);
    return true;
}

/* ---- Utilities ---- */

bool float3_isfinite(float3_t v) {
    return isfinite(v.x) && isfinite(v.y) && isfinite(v.z);
}

bool float3_equal_eps(float3_t a, float3_t b, float eps) {
    return (fabsf(a.x - b.x) <= eps) &&
           (fabsf(a.y - b.y) <= eps) &&
           (fabsf(a.z - b.z) <= eps);
}
