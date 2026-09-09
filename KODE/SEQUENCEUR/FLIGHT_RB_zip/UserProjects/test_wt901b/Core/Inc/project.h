#include <stdbool.h>

#include "float3.h"
#include "quaternion.h"

typedef struct rocket_attitude_dynamics_t {
	quatf_t q_init; /* quaternion attitude before launch (Body -> Earth) */
	quatf_t q; /* attitude quaternion (Body -> Earth) */
	float elevation_deg; /* elevation angle in degrees (0 = horizontal, 90 = vertical) */
	float azimuth_deg; /* azimuth angle in degrees (0 = forward, 90 = right) */
} rocket_attitude_dynamics_t;

void uart_apex_callback(void);

void setup();

void loop();