/* quaternion_dynamic.h
 * Quaternion dynamics (C99) — API only.
 *
 * Purpose:
 *  - Build delta quaternion from gyro ω (rad/s) and dt (s) using exponential map
 *  - Propagate attitude quaternion: q <- q ⊗ dq   (Body -> Earth convention)
 *
 * Assumptions:
 *  - dt is constant (caller provides dt)
 *  - no small-angle approximation (always uses sin/cos form)
 *  - ω is expressed in BODY frame
 *  - q represents Body -> Earth rotation
 */

#ifndef QUATERNION_DYNAMIC_H
#define QUATERNION_DYNAMIC_H

#ifdef __cplusplus
extern "C" {
#endif

#include <stdbool.h>
#include "utils/quaternion.h"
#include "utils/float3.h"

/* Build delta quaternion dq from body angular rate ω (rad/s) and dt (s)
 * dq = exp( 0.5*dt*Omega(ω) )
 * Returns false if ω is near zero or inputs not finite.
 */
bool    quatdyn_delta_from_omega(float3_t omega_body_rad_s, float dt_s, quatf_t *out_dq);

/*  Alternative version using the formula with dq/dt = 0.5 * q ⊗ Omega(ω) and a simple Euler step:
 * dq = 0.5 * dt * q ⊗ Omega(ω)
 */
bool    quatdyn_delta_from_omega_2(float3_t omega_body_rad_s, float dt_s, quatf_t *out_dq, quatf_t q);

/* Pure propagate: returns q_next = q x dq
 * Returns identity if it fails (you can choose to not use this if you prefer strict error handling).
 */
quatf_t quatdyn_propagate(quatf_t q_body_to_earth, float3_t omega_body_rad_s, float dt_s);

/* Same as quatdyn_propagate but uses the alternative delta quaternion calculation */
quatf_t quatdyn_propagate_2(quatf_t q_body_to_earth, float3_t omega_body_rad_s, float dt_s);

/* Same as quatdyn_propagate but uses the alternative delta quaternion calculation */
quatf_t quatdyn_propagate_3(quatf_t q_body_to_earth, float3_t omega_body_rad_s, float dt_s);

/* In-place propagate: q <- q x dq ; returns false on failure (q unchanged) */
bool    quatdyn_propagate_ip(quatf_t *q_body_to_earth, float3_t omega_body_rad_s, float dt_s);

/* Same as quatdyn_propagate_ip but uses the alternative delta quaternion calculation */ 
bool    quatdyn_propagate_ip_2(quatf_t *q_body_to_earth, float3_t omega_body_rad_s, float dt_s);

/* Same as quatdyn_propagate_ip but uses the alternative delta quaternion calculation */ 
bool    quatdyn_propagate_ip_3(quatf_t *q_body_to_earth, float3_t omega_body_rad_s, float dt_s);

#ifdef __cplusplus
}
#endif

#endif /* QUATERNION_DYNAMIC_H */
