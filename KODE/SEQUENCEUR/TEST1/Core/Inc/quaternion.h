/* quaternion.h
 * Quaternion math library (C99) — header (API only).
 *
 * Rule:
 *  - op(a, b)        -> returns result
 *  - op_ip(&a, b)    -> modifies a in-place
 *
 * Quaternion q = w + xi + yj + zk stored as (w, x, y, z).
 */

#ifndef QUATERNION_H
#define QUATERNION_H

#ifdef __cplusplus
extern "C" {
#endif

#include <stdbool.h>
#include <stdint.h>
#include "float3.h"

/* ----------------------------- Types ----------------------------------- */

typedef struct {
    float w;
    float x;
    float y;
    float z;
} quatf_t;

#define QUATF_IDENTITY (quatf_t){ .w = 1.0f, .x = 0.0f, .y = 0.0f, .z = 0.0f }
#define QUATF_ZERO     (quatf_t){ .w = 0.0f, .x = 0.0f, .y = 0.0f, .z = 0.0f }

/* -------------------------- Constructors -------------------------------- */

quatf_t   quatf_make(float w, float x, float y, float z);

quatf_t   quatf_from_vec3(float3_t v); /* pure quaternion (0, v) */
quatf_t   quatf_from_axis_angle(float3_t axis, float angle_rad); /* axis must be normalized */
float3_t  quatf_to_vec3(quatf_t q); /* returns vector part (x, y, z) */

/* --------------------------- Basic ops ---------------------------------- */

quatf_t   quatf_add(quatf_t a, quatf_t b);
void      quatf_add_ip(quatf_t *a, quatf_t b);

quatf_t   quatf_sub(quatf_t a, quatf_t b);
void      quatf_sub_ip(quatf_t *a, quatf_t b);

quatf_t   quatf_scale(quatf_t q, float s);
void      quatf_scale_ip(quatf_t *q, float s);

/* Hamilton product: a x b */
quatf_t   quatf_mul(quatf_t a, quatf_t b);
void      quatf_mul_ip(quatf_t *a, quatf_t b); /* a <- a x b */

/* ---------------------------- Metrics ----------------------------------- */

float     quatf_dot(quatf_t a, quatf_t b);
float     quatf_norm2(quatf_t q);
float     quatf_norm(quatf_t q);

/* ----------------------- Conjugate / inverse ---------------------------- */

quatf_t   quatf_conj(quatf_t q);
void      quatf_conj_ip(quatf_t *q);

/* out_qinv = conj(q) / ||q||^2 ; returns false if ||q|| too small or invalid */
bool      quatf_inv_get(quatf_t q, quatf_t *out_qinv);

/* Optional true in-place inverse: q <- q^{-1} */
bool      quatf_inv_ip(quatf_t *q);

/* --------------------------- Normalize ---------------------------------- */

/* Pure: returns normalized copy (returns identity if invalid) */
quatf_t   quatf_normalized(quatf_t q);

/* In-place: q <- q / ||q|| ; returns false if invalid (q unchanged) */
bool      quatf_normalized_ip(quatf_t *q);

/* ------------------------ Vector rotation -------------------------------- */

/* v_rot = q x (0,v) x conj(q)  (no assumption about q normalization here) */
float3_t  quatf_rotate_vector(quatf_t q, float3_t v);

/* ----------------------- Validity / compare ------------------------------ */

bool      quatf_isfinite(quatf_t q);
bool      quatf_equal_eps(quatf_t a, quatf_t b, float eps);

/* Default epsilon used internally */
float     quatf_default_eps(void);

#ifdef __cplusplus
}
#endif

#endif /* QUATERNION_H */
