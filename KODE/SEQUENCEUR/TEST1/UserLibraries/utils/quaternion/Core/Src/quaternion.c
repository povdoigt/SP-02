/* quaternion.c
 * Quaternion math library (C99) — implementation.
 */

#include "quaternion.h"
#include <math.h>   /* sqrtf, isfinite, fabsf */

#define QUATF_EPS_NORM (1e-20f)

quatf_t quatf_make(float w, float x, float y, float z) {
    return (quatf_t){ .w = w, .x = x, .y = y, .z = z };
}

quatf_t quatf_from_vec3(float3_t v) {
    return quatf_make(0.0f, v.x, v.y, v.z);
}

quatf_t quatf_from_axis_angle(float3_t axis, float angle_rad) {
    float half_angle = 0.5f * angle_rad;
    float sin_half = sinf(half_angle);
    return quatf_make(
        cosf(half_angle),
        axis.x * sin_half,
        axis.y * sin_half,
        axis.z * sin_half
    );
}

float3_t quatf_to_vec3(quatf_t q) {
    return (float3_t){ .x = q.x, .y = q.y, .z = q.z };
}

/* --------------------------- Basic ops ---------------------------------- */

quatf_t quatf_add(quatf_t a, quatf_t b) {
    return quatf_make(a.w + b.w, a.x + b.x, a.y + b.y, a.z + b.z);
}

void quatf_add_ip(quatf_t *a, quatf_t b) {
    if (!a) return;
    a->w += b.w; a->x += b.x; a->y += b.y; a->z += b.z;
}

quatf_t quatf_sub(quatf_t a, quatf_t b) {
    return quatf_make(a.w - b.w, a.x - b.x, a.y - b.y, a.z - b.z);
}

void quatf_sub_ip(quatf_t *a, quatf_t b) {
    if (!a) return;
    a->w -= b.w; a->x -= b.x; a->y -= b.y; a->z -= b.z;
}

quatf_t quatf_scale(quatf_t q, float s) {
    return quatf_make(q.w * s, q.x * s, q.y * s, q.z * s);
}

void quatf_scale_ip(quatf_t *q, float s) {
    if (!q) return;
    q->w *= s; q->x *= s; q->y *= s; q->z *= s;
}

/* Hamilton product */
quatf_t quatf_mul(quatf_t a, quatf_t b) {
    return quatf_make(
        a.w*b.w - a.x*b.x - a.y*b.y - a.z*b.z,
        a.w*b.x + a.x*b.w + a.y*b.z - a.z*b.y,
        a.w*b.y - a.x*b.z + a.y*b.w + a.z*b.x,
        a.w*b.z + a.x*b.y - a.y*b.x + a.z*b.w
    );
}

void quatf_mul_ip(quatf_t *a, quatf_t b) {
    if (!a) return;
    quatf_t temp = *a;
    a->w = temp.w*b.w - temp.x*b.x - temp.y*b.y - temp.z*b.z;
    a->x = temp.w*b.x + temp.x*b.w + temp.y*b.z - temp.z*b.y;
    a->y = temp.w*b.y - temp.x*b.z + temp.y*b.w + temp.z*b.x;
    a->z = temp.w*b.z + temp.x*b.y - temp.y*b.x + temp.z*b.w;
}

/* ---------------------------- Metrics ----------------------------------- */

float quatf_dot(quatf_t a, quatf_t b) {
    return a.w*b.w + a.x*b.x + a.y*b.y + a.z*b.z;
}

float quatf_norm2(quatf_t q) {
    return quatf_dot(q, q);
}

float quatf_norm(quatf_t q) {
    return sqrtf(quatf_norm2(q));
}

/* ----------------------- Conjugate / inverse ---------------------------- */

quatf_t quatf_conj(quatf_t q) {
    return quatf_make(q.w, -q.x, -q.y, -q.z);
}

void quatf_conj_ip(quatf_t *q) {
    if (!q) return;
    q->x = -q->x; q->y = -q->y; q->z = -q->z;
}

bool quatf_inv_get(quatf_t q, quatf_t *out_qinv) {
    if (!out_qinv) return false;

    float n2 = quatf_norm2(q);
    if (!isfinite(n2) || n2 < QUATF_EPS_NORM) return false;

    quatf_t qc = quatf_conj(q);
    *out_qinv = quatf_scale(qc, 1.0f / n2);
    return quatf_isfinite(*out_qinv);
}

bool quatf_inv_ip(quatf_t *q) {
    if (!q) return false;
    quatf_t inv;
    if (!quatf_inv_get(*q, &inv)) return false;
    *q = inv;
    return true;
}

/* --------------------------- Normalize ---------------------------------- */

quatf_t quatf_normalized(quatf_t q) {
    float n2 = quatf_norm2(q);
    if (!isfinite(n2) || n2 < QUATF_EPS_NORM) {
        return QUATF_IDENTITY; /* choix “safe” pour une version pure */
    }
    float inv_n = 1.0f / sqrtf(n2);
    return quatf_scale(q, inv_n);
}

bool quatf_normalized_ip(quatf_t *q) {
    if (!q) return false;

    float n2 = quatf_norm2(*q);
    if (!isfinite(n2) || n2 < QUATF_EPS_NORM) return false;

    float inv_n = 1.0f / sqrtf(n2);
    quatf_scale_ip(q, inv_n);
    return true;
}

/* ------------------------ Vector rotation ------------------------------- */

float3_t quatf_rotate_vector(quatf_t q, float3_t v) {
    quatf_t vq = quatf_from_vec3(v);
    quatf_t qc = quatf_conj(q);

    quatf_t tmp = quatf_mul(q, vq);
    quatf_t res = quatf_mul(tmp, qc);

    return quatf_to_vec3(res);
}

/* ----------------------- Validity / compare ------------------------------ */

bool quatf_isfinite(quatf_t q) {
    return isfinite(q.w) && isfinite(q.x) && isfinite(q.y) && isfinite(q.z);
}

bool quatf_equal_eps(quatf_t a, quatf_t b, float eps) {
    return (fabsf(a.w - b.w) <= eps) &&
           (fabsf(a.x - b.x) <= eps) &&
           (fabsf(a.y - b.y) <= eps) &&
           (fabsf(a.z - b.z) <= eps);
}

float quatf_default_eps(void) {
    return QUATF_EPS_NORM;
}
