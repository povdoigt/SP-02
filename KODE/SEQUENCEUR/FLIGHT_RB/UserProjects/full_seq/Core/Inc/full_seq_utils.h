#ifndef FULL_SEQ_UTILS_H
#define FULL_SEQ_UTILS_H

#include "main.h"

#include "led.h"

#include "float3.h"
#include "quaternion.h"
#include "waveform.h"
#include "waveform_def.h"


/* ===================================================
   CONSTANTS
   =================================================== */

// Les temps sont en millisecondes par rapport au T0 (lancement)

#define T_ALPHA_BETA_0 3600		/* Fin de propulsion du premier étage */
#define T_ALPHA_BETA_1 3700		/* Séparation du premier étage */
#define T_ALPHA_BETA_2 16200	/* Apogée de l'ensemble de la fusée (1er et 2e étage sans séparation) */
#define T_ALPHA_BETA_3 34500	/* Atterrisage de l'ensemble de la fusée (1er et 2e étage sans séparation) */

#define T_ALPHA_0 15100	/* Apogée du premier étage */
#define T_ALPHA_1 33400	/* Atterrissage du premier étage */

#define T_BETA_0 4700	/* Allumage du second étage */
#define T_BETA_1 4700	/* Confirmation de l'allumage du second étage */
#define T_BETA_2 15900	/* Apogée du second étage si séparation et actif */
#define T_BETA_3 35400	/* Atterrissage du second étage si séparation et actif */
#define T_BETA_4 15000	/* Apogée du second étage si séparation mais passif */
#define T_BETA_5 33500	/* Atterrissage du second étage si séparation mais passif */

#define DEG_TO_RAD (M_PI / 180.0f)
#define RAD_TO_DEG (180.0f / M_PI)

#define WT901B_FREQUENCY_HZ 100 /* Fréquence de lecture des données du WT901B, à ajuster selon la configuration du capteur */
#define WT901B_PERIOD_S (1.0f / WT901B_FREQUENCY_HZ)

#define ATTITUDE_ELEVATION_GOAL_DEG 72 /* Objectif d'angle d'élévation pour le second étage, à ajuster selon la trajectoire souhaitée */
#define ATTITUDE_AZIMUTH_GOAL_DEG 0 /* Objectif d'angle d'azimut pour le second étage, à ajuster selon la trajectoire souhaitée */



/* ===================================================
   PHASE TRANSITION
   =================================================== */

typedef enum stage_phase_type_t {
	STAGE_PHASE_STAGE_INIT,
	STAGE_PHASE_STAGE_FLIGHT,
} stage_phase_type_t;

typedef struct stage_phase_transition_t {
	stage_phase_type_t stage_phase_type;
	uint8_t *phase_variable;
} stage_phase_transition_t;

void phase_transition_init(stage_phase_transition_t *transition, stage_phase_type_t stage_phase_type, uint8_t *phase_variable);
void change_state_and_notify(stage_phase_transition_t *transition, uint8_t new_state);




/* ===================================================
   ROCKET
   =================================================== */

typedef struct rocket_dynamics_t {
	float3_t v_up_body; /* direction of the rocket’s “up” axis in body frame */
	float initial_elevation_deg; /* initial elevation angle in degrees (0 = horizontal, 90 = vertical) */
	
	quatf_t q; /* attitude quaternion (Body -> Earth) */
	float elevation_deg; /* elevation angle in degrees (0 = horizontal, 90 = vertical) */
	float azimuth_deg; /* azimuth angle in degrees (0 = forward, 90 = right) */

	float3_t accel_g; /* acceleration vector in body frame, in g */
	float pressure_variation_pa_s; /* pressure variation in Pa/s */
} rocket_dynamics_t;

typedef enum rocket_stage_t {
	ROCKET_FIRST_STAGE,
	ROCKET_SECOND_STAGE,
	ROCKET_STAGE_NO_SET,
} rocket_stage_t;

typedef struct rocket_state_t {
	rocket_stage_t stage;
	rocket_dynamics_t dynamics;
	stage_phase_transition_t stage_phase_transition;
	led_rgb_t *led_rgb;
	waveform_space_t *current_waveform_space;
	uint32_t t_launch;
	bool is_launch_confirmed;
	bool is_separation_confirmed;
	bool is_second_burn_confirmed;
} rocket_state_t;

void init_rocket_state(rocket_state_t *rocket_state, led_rgb_t *led_rgb);




/* ===================================================
   WAITING BUTTON STATE
   =================================================== */

typedef enum waiting_button_state_t {
	WAITING_BUTTON_STATE_WAITING_PRESS,
	WAITING_BUTTON_STATE_WAITING_RELEASE,
} waiting_button_state_t;

typedef struct waiting_button_t {
	waiting_button_state_t state;

	led_rgb_t *led_rgb;
	
	GPIO_TypeDef *button_gpio_port;
	uint32_t t_led;
	uint16_t button_gpio_pin;
	bool inverted_logic;
	bool led_on;
} waiting_button_t;

void waiting_button_init(waiting_button_t *waiting_button, GPIO_TypeDef *button_gpio_port, uint16_t button_gpio_pin, led_rgb_t *led_rgb, bool inverted_logic);
void waiting_button_set_next_phase(waiting_button_t *waiting_button);
bool waiting_button_play(waiting_button_t *waiting_button);

#endif // FULL_SEQ_UTILS_H
