#ifndef FULL_SEQ_UTILS_H
#define FULL_SEQ_UTILS_H

#include "main.h"

#include "led.h"

#include "float3.h"
#include "quaternion.h"
#include "stm32f072xb.h"
#include "stm32f0xx_hal_gpio.h"
#include "waveform.h"
#include "waveform_built_in.h"
#include "waveform_def.h"
#include "waveform_scalar_arithmetic.h"
#include <stdint.h>




/* ===================================================
   GPIO MAPPING
   =================================================== */

#define MAX_INPUT_NUMBER				6

#define JACK_LAUNCH_GPIO_Port			IN_TRG_N1_GPIO_Port
#define JACK_LAUNCH_Pin					IN_TRG_N1_Pin

#define SEPARATION_GPIO_Port			IN_TRG_N2_GPIO_Port
#define SEPARATION_Pin					IN_TRG_N2_Pin

#define JACK_READY_GPIO_Port			IN_TRG_N3_GPIO_Port
#define JACK_READY_Pin					IN_TRG_N3_Pin

#define STAGE2_IGNITION_GPIO_Port		OUT_N2_GPIO_Port
#define STAGE2_IGNITION_Pin				OUT_N2_Pin

#define EVENT_UART_GPIO_Port			TX_OPTO_N1_GPIO_Port
#define EVENT_UART_Pin					TX_OPTO_N1_Pin

#define CAMERA_GPIO_Port				TX_OPTO_N2_GPIO_Port
#define CAMERA_Pin						TX_OPTO_N2_Pin


void update_gpio_input_states(GPIO_TypeDef *input_gpio_port[], uint16_t input_gpio_pin[], GPIO_PinState input_gpio_state[], uint8_t max_input_number);





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
	STAGE_PHASE_STAGE_BOARD_FUNC,
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
	GPIO_TypeDef				*input_gpio_ports[MAX_INPUT_NUMBER];
	uint16_t		 			input_gpio_pins[MAX_INPUT_NUMBER];
	GPIO_PinState	 			input_gpio_states[MAX_INPUT_NUMBER];
	rocket_stage_t	 			stage;
	rocket_dynamics_t			dynamics;
	stage_phase_transition_t	stage_phase_transition;
	led_rgb_t 					*led_rgb;
	waveform_space_t 			*current_waveform_space;
	uint32_t 					t_launch;
	bool 						is_launch_confirmed;
	bool 						is_separation_confirmed;
	bool 						is_second_burn_confirmed;	
} rocket_state_t;

void rocket_state_init(rocket_state_t *rocket_state, led_rgb_t *led_rgb);
void rocket_state_setup_gpio(rocket_state_t *rocket_state, GPIO_TypeDef *input_gpio_port[], uint16_t input_gpio_pin[]);





/* ===================================================
   LED STATE
   =================================================== */

extern waveform_space_t waveform_wait_button;
extern waveform_space_t waveform_wait_jack_ready;
extern waveform_space_t waveform_wait_jack_launch;
extern waveform_space_t waveform_wait_sepa;
extern waveform_space_t waveform_wait_jack_launch_sepa;
extern waveform_space_t waveform_perform_sepa;

extern const waveform_space_mult_const_t 	waveform_wait_button_ctx;
extern const waveform_space_mult_const_t 	waveform_wait_jack_ready_ctx;
extern const waveform_space_mult_const_t 	waveform_wait_jack_launch_ctx;
extern const waveform_space_mult_const_t 	waveform_wait_sepa_ctx;
extern const waveform_space_add_t			waveform_wait_jack_launch_sepa_ctx;
extern const waveform_space_mult_const_t 	waveform_perform_sepa_ctx;



/* ===================================================
   WAITING BUTTON STATE
   =================================================== */

typedef enum waiting_button_state_t {
	WAITING_BUTTON_STATE_WAITING_START,
	WAITING_BUTTON_STATE_WAITING_PRESS,
	WAITING_BUTTON_STATE_WAITING_RELEASE,
} waiting_button_state_t;

typedef struct waiting_button_t {
	waveform_space_t wait_waveform_space;
	GPIO_TypeDef *button_gpio_port;
	waveform_space_t **rocket_waveform_space;
	waiting_button_state_t state;
	uint16_t button_gpio_pin;
	bool inverted_logic;
	bool led_on;
} waiting_button_t;

void waiting_button_init(waiting_button_t *waiting_button, GPIO_TypeDef *button_gpio_port, uint16_t button_gpio_pin, waveform_space_t **rocket_waveform_space, bool inverted_logic);
// void waiting_button_set_next_phase(waiting_button_t *waiting_button);
bool waiting_button_play(waiting_button_t *waiting_button);



/* ===================================================
   BOARD FUNCTIONS
   =================================================== */

typedef void(*board_func_t)(rocket_state_t *rocket_state);


#endif // FULL_SEQ_UTILS_H
