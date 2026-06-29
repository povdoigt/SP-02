#ifndef WAVEFORM_DEF_H
#define WAVEFORM_DEF_H

#include "float3.h"
#include <stdint.h>
#include <stdbool.h>

// Function that takes a phase t in [0, 1] and returns an intensity in [0, 1]
typedef float (*wave_func_scalar_t)(float t, const void *ctx);
// Function that takes a phase t in [0, 1] and returns a color in [0, 1]^3
typedef float3_t (*wave_func_space_t)(float t, const void *ctx);

typedef enum {
	WAVEFORM_SCALAR,
	WAVEFORM_SPACE
} waveform_kind_t;

typedef union wave_func_u {
	wave_func_scalar_t	scalar_wave_function;
	wave_func_space_t	space_wave_function;
} wave_func_u;

typedef struct waveform_common_t {
	waveform_kind_t		 kind;		// Whether it's a scalar or space waveform
	const void			*ctx;		// Context for the waveform function
	uint32_t			 time;		// Duration of the waveform in milliseconds
	uint32_t			 t0;		// Start time of the waveform
	bool				 periodic;	// Whether the waveform should repeat periodically
	bool				 active;	// Whether the waveform is currently active
} waveform_common_t;

typedef struct waveform_generic_t {
	waveform_common_t	common;
	wave_func_u			wave_function;
} waveform_generic_t;

typedef struct waveform_scalar_t {
	waveform_common_t	common;			// Common fields for both scalar and space waveforms
	wave_func_scalar_t	wave_function;	// Function pointer for the waveform
} waveform_scalar_t;

typedef struct waveform_space_t {
	waveform_common_t	common;			// Common fields for both scalar and space waveforms
	wave_func_space_t	wave_function;	// Function pointer for the waveform
} waveform_space_t;

void Waveform_Init_Scalar(waveform_scalar_t *waveform, wave_func_scalar_t wave_function,
						  const void *ctx, uint32_t duration, bool periodic);
void Waveform_Init_Space(waveform_space_t *waveform, wave_func_space_t wave_function,
						 const void *ctx, uint32_t duration, bool periodic);
void Waveform_Restart(waveform_generic_t *waveform);
bool Waveform_IsActive(waveform_generic_t *waveform);
float Waveform_Play_Scalar(waveform_scalar_t *waveform, uint32_t current_time);
float3_t Waveform_Play_Space(waveform_space_t *waveform, uint32_t current_time);

#endif /* WAVEFORM_DEF_H */