#include "STS.h"
#include "WT901B.h"

#include "quaternion.h"
#include "data_topic.h"

#include <math.h>




/* ===================================================
   CONSTANTS
   =================================================== */

// Les temps sont en millisecondes par rapport au T0 (lancement)

#define T_ALPHA_BETA_0 5000 /* Fin de propulsion du premier étage */
#define T_ALPHA_BETA_1 5300 /* Séparation du premier étage */
#define T_ALPHA_BETA_2 0 /* Apogée de l'ensemble de la fusée (1er et 2e étage sans séparation) */
#define T_ALPHA_BETA_3 0 /* Atterrisage de l'ensemble de la fusée (1er et 2e étage sans séparation) */

#define T_ALPHA_0 0 /* Apogée du premier étage */
#define T_ALPHA_1 0 /* Atterrissage du premier étage */

#define T_BETA_0 0 /* Allumage du second étage */
#define T_BETA_1 0 /* Confirmation de l'allumage du second étage */
#define T_BETA_2 0 /* Apogée du second étage si séparation et actif */
#define T_BETA_3 0 /* Atterrissage du second étage si séparation et actif */
#define T_BETA_4 0 /* Apogée du second étage si séparation mais passif */
#define T_BETA_5 0 /* Atterrissage du second étage si séparation mais passif */

#define DEG_TO_RAD (M_PI / 180.0f)
#define RAD_TO_DEG (180.0f / M_PI)

#define WT901B_FREQUENCY_HZ 100 /* Fréquence de lecture des données du WT901B, à ajuster selon la configuration du capteur */
#define WT901B_PERIOD_S (1.0f / WT901B_FREQUENCY_HZ)

#define ATTITUDE_ELEVATION_GOAL_DEG 72 /* Objectif d'angle d'élévation pour le second étage, à ajuster selon la trajectoire souhaitée */
#define ATTITUDE_AZIMUTH_GOAL_DEG 0 /* Objectif d'angle d'azimut pour le second étage, à ajuster selon la trajectoire souhaitée */




/* ===================================================
   ROCKET ATTITUDE DYNAMICS
   =================================================== */

typedef struct rocket_attitude_dynamics_t {
	float3_t v_up_body; /* direction of the rocket’s “up” axis in body frame */
	float initial_elevation_deg; /* initial elevation angle in degrees (0 = horizontal, 90 = vertical) */
	
	quatf_t q; /* attitude quaternion (Body -> Earth) */
	float elevation_deg; /* elevation angle in degrees (0 = horizontal, 90 = vertical) */
	float azimuth_deg; /* azimuth angle in degrees (0 = forward, 90 = right) */
} rocket_attitude_dynamics_t;

typedef enum rocket_stage_t {
	ROCKET_FIRST_STAGE,
	ROCKET_SECOND_STAGE,
} rocket_stage_t;




/* ===================================================
   MAIN FUNCTIONS
   =================================================== */

void setup(void);

void loop(void);



/* ===================================================
   SETUP FUNCTIONS
   =================================================== */

void setup_servomotors(void);
void setup_data_acquisition(void);
void setup_attitude(void);




/* ===================================================
   SENSOR DATA CALLBACKS
   =================================================== */


void on_new_accel_frame(data_sub_t *sub);
void on_new_gyro_frame(data_sub_t *sub);
void on_new_pressure_frame(data_sub_t *sub);




/* ===================================================
   FIRST STAGE STATE MACHINE
   =================================================== */


typedef enum first_stage_initialisation_phase_t {
	FIRST_STAGE_AF_ZERO,
	FIRST_STAGE_WAIT_AF_ZERO,
	FIRST_STAGE_SEPA_ZERO,
	FIRST_STAGE_WAIT_SEPA_ZERO,
	FIRST_STAGE_WAIT_STAGE_ASSEMBLY_CONFIRMATION,
} first_stage_initialisation_phase_t;

typedef enum first_stage_flight_phase_t {
	FIRST_STAGE_INITIALISATION,
	FIRST_STAGE_WAIT_LAUNCH_CONFIRMATION,
	FIRST_STAGE_WAIT_BURN_END,
	FIRST_STAGE_SEPARATION,
	FIRST_STAGE_WAIT_SEPARATION_CONFIRMATION,
	FIRST_STAGE_WAIT_APOGEE_CONFIRMATION,
} first_stage_flight_phase_t;

void first_stage_initialisation(void);
void first_stage_state_machine(void);




/* ===================================================
   SECOND STAGE STATE MACHINE
   =================================================== */

typedef enum second_stage_initialisation_phase_t {
	SECOND_STAGE_NOP, // TODO: fill this in with actual initialisation phases
} second_stage_initialisation_phase_t;

typedef enum second_stage_flight_phase_t {
	SECOND_STAGE_WAIT_STAGE_ASSEMBLY_CONFIRMATION,
	SECOND_STAGE_WAIT_LAUNCH_CONFIRMATION,
	SECOND_STAGE_WAIT_SEPARATION_CONFIRMATION,
	SECOND_STAGE_WAIT_ATTITUDE_CONFIRMATION,
	SECOND_STAGE_BURN_SECOND_BURN_COMMAND,
	SECOND_STAGE_WAIT_SECOND_BURN_CONFIRMATION,
	SECOND_STAGE_WAIT_APOGEE_CONFIRMATION,
} second_stage_flight_phase_t;

void second_stage_initialisation(void);
void second_stage_state_machine(void);



