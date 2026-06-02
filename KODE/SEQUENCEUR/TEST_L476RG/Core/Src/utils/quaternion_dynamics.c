/* quaternion_dynamic.c
 * Quaternion dynamics (C99) — implementation.
 */

#include "utils/quaternion_dynamics.h"
#include "utils/float3.h"
#include "utils/quaternion.h"
#include <math.h> /* sqrtf, sinf, cosf, isfinite */

bool quatdyn_delta_from_omega(float3_t omega_body_rad_s, float dt_s, quatf_t *out_dq) {
    if (!out_dq) return false;

    /* Compute norm of omega */
    float omega_norm = float3_norm(omega_body_rad_s);
    if (!isfinite(omega_norm) || omega_norm < 1e-10f) {
        return false;
    }

    /* Compute half-angle */
    float half_theta = 0.5f * omega_norm * dt_s;
    float sin_half_theta = sinf(half_theta);
    float cos_half_theta = cosf(half_theta);

    /* Compute unit rotation axis */
    float3_t axis = float3_scale(omega_body_rad_s, 1.0f / omega_norm);

    /* Build delta quaternion */
    *out_dq = quatf_make(
        cos_half_theta,
        axis.x * sin_half_theta,
        axis.y * sin_half_theta,
        axis.z * sin_half_theta
    );

    return true;
}

bool quatdyn_delta_from_omega_2(float3_t omega_body_rad_s, float dt_s, quatf_t *out_dq, quatf_t q) {
	if (!out_dq) return false;

	/* Build pure quaternion (0, ω) */
	quatf_t omega_quat = quatf_from_vec3(omega_body_rad_s);

	/* dq = 0.5 * dt * q ⊗ Omega(ω) */
	
	*out_dq = quatf_normalized(
		quatf_scale(
			quatf_mul(q, omega_quat),
			0.5f * dt_s
		)
	);

	return true;
}

quatf_t quatdyn_propagate(quatf_t q_body_to_earth, float3_t omega_body_rad_s, float dt_s) {
    quatf_t dq;
    if (!quatdyn_delta_from_omega(omega_body_rad_s, dt_s, &dq)) {
        return QUATF_IDENTITY;
    }
    return quatf_mul(q_body_to_earth, dq);
}

quatf_t quatdyn_propagate_2(quatf_t q_body_to_earth, float3_t omega_body_rad_s, float dt_s) {
	quatf_t dq;
	if (!quatdyn_delta_from_omega_2(omega_body_rad_s, dt_s, &dq, q_body_to_earth)) {
		return QUATF_IDENTITY;
	}
	return quatf_mul(q_body_to_earth, dq);
}

quatf_t quatdyn_propagate_3(quatf_t q_body_to_earth, float3_t omega_body_rad_s, float dt_s) {
	// q_k+1 = (I + (T/2) S(w)) q_k
	// where S(w) is the skew-symmetric matrix of w
	quatf_t q = q_body_to_earth;
	float3_t w = omega_body_rad_s;
	return quatf_normalized((quatf_t){
		.w = q.w + 0.5f * dt_s * (- w.x * q.x - w.y * q.y - w.z * q.z),
		.x = q.x + 0.5f * dt_s * (+ w.x * q.w + w.z * q.y - w.y * q.z),
		.y = q.y + 0.5f * dt_s * (+ w.y * q.w - w.z * q.x + w.x * q.z),
		.z = q.z + 0.5f * dt_s * (+ w.z * q.w + w.y * q.x - w.x * q.y)
	});	
}

bool quatdyn_propagate_ip(quatf_t *q_body_to_earth, float3_t omega_body_rad_s, float dt_s) {
    if (!q_body_to_earth) return false;

    quatf_t dq;
    if (!quatdyn_delta_from_omega(omega_body_rad_s, dt_s, &dq)) {
        return false;
    }

    quatf_mul_ip(q_body_to_earth, dq);
    return true;
}

bool quatdyn_propagate_ip_2(quatf_t *q_body_to_earth, float3_t omega_body_rad_s, float dt_s) {
	if (!q_body_to_earth) return false;

	quatf_t dq;
	if (!quatdyn_delta_from_omega_2(omega_body_rad_s, dt_s, &dq, *q_body_to_earth)) {
		return false;
	}

	quatf_mul_ip(q_body_to_earth, dq);
	return true;
}

bool quatdyn_propagate_ip_3(quatf_t *q_body_to_earth, float3_t omega_body_rad_s, float dt_s) {
	if (!q_body_to_earth) return false;

	quatf_t q = *q_body_to_earth;
	float3_t w = omega_body_rad_s;
	q_body_to_earth->w = q.w + 0.5f * dt_s * (- w.x * q.x - w.y * q.y - w.z * q.z);
	q_body_to_earth->x = q.x + 0.5f * dt_s * (+ w.x * q.w + w.z * q.y - w.y * q.z);
	q_body_to_earth->y = q.y + 0.5f * dt_s * (+ w.y * q.w - w.z * q.x + w.x * q.z);
	q_body_to_earth->z = q.z + 0.5f * dt_s * (+ w.z * q.w + w.y * q.x - w.x * q.y);
	quatf_normalized_ip(q_body_to_earth);
	return true;
}