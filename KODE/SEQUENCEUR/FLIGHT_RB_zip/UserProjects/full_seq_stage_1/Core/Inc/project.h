#include "STS.h"
// #include "WT901B.h"

// #include "quaternion.h"
#include "data_topic.h"

// #include <math.h>
#include <stdint.h>




/* ===================================================
   CONSTANTS
   =================================================== */

// Les temps sont en millisecondes par rapport au T0 (lancement)

#define T_ALPHA_BETA_0 3600		/* Fin de propulsion du premier étage */
#define T_ALPHA_BETA_1 3700		/* Séparation du premier étage */
#define T_ALPHA_BETA_2 16200	/* Apogée de l'ensemble de la fusée (1er et 2e étage sans séparation) */
// #define T_ALPHA_BETA_3 34500	/* Atterrisage de l'ensemble de la fusée (1er et 2e étage sans séparation) */

#define T_ALPHA_0 15100			/* Apogée du premier étage */
// #define T_ALPHA_1 33400 /* Atterrissage du premier étage */

#define WT901B_FREQUENCY_HZ 100 /* Fréquence de lecture des données du WT901B, à ajuster selon la configuration du capteur */
#define WT901B_PERIOD_S (1.0f / WT901B_FREQUENCY_HZ)




/* ===================================================
   ROCKET ATTITUDE DYNAMICS
   =================================================== */

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

void setup_servomotors_stage_1(void);
void setup_data_acquisition_stage_1(void);
void setup_uart_buffers_stage_1(void);
void setup_event_uart(void);



/* ===================================================
   SENSOR DATA CALLBACKS
   =================================================== */

void on_new_pressure_frame(data_sub_t *sub);


typedef enum stage_phase_type_t {
	STAGE_PHASE_FIRST_STAGE_INIT,
	STAGE_PHASE_FIRST_STAGE_FLIGHT,
} stage_phase_type_t;

/* ===================================================
   FIRST STAGE STATE MACHINE
   =================================================== */

typedef enum first_stage_initialisation_phase_t {
	FIRST_STAGE_INIT_AF_ZERO,
	FIRST_STAGE_INIT_WAIT_AF_ZERO,
	FIRST_STAGE_INIT_PARA_ZERO,
	FIRST_STAGE_INIT_WAIT_PARA_ZERO,
	FIRST_STAGE_INIT_SEPA_ZERO,
	FIRST_STAGE_INIT_WAIT_SEPA_ZERO,
	FIRST_STAGE_INIT_WAIT_JACK,
	FIRST_STAGE_INIT_WAIT_STAGE_ASSEMBLY_CONFIRMATION,
	FIRST_STAGE_INIT_LOCK_STAGE,
	FIRST_STAGE_INIT_WAIT_LOCK_STAGE,
	FIRST_STAGE_INIT_WAIT_BUTTON,
} first_stage_initialisation_phase_t;

typedef enum first_stage_flight_phase_t {
	FIRST_STAGE_FLIGHT_INITIALISATION,
	FIRST_STAGE_FLIGHT_WAIT_LAUNCH_CONFIRMATION,
	FIRST_STAGE_FLIGHT_WAIT_BURN_END,
	FIRST_STAGE_FLIGHT_SEPARATION,
	FIRST_STAGE_FLIGHT_WAIT_SEPARATION_CONFIRMATION,
	FIRST_STAGE_FLIGHT_WAIT_APOGEE_CONFIRMATION,
	FIRST_STAGE_FLIGHT_APOGEE,
	FIRST_STAGE_FLIGHT_WAIT_DROGUE_CONFIRMATION,
	FIRST_STAGE_FLIGHT_WAIT_LANDING_CONFIRMATION,
	FIRST_STAGE_FLIGHT_IDLE,
} first_stage_flight_phase_t;

void first_stage_init_state_machine(void);
void first_stage_flight_state_machine(void);


/* ===================================================
   UTILS
   =================================================== */

typedef struct stage_phase_transition_t {
	stage_phase_type_t stage_phase_type;
	uint8_t *phase_variable;
} stage_phase_transition_t;

void change_state_and_notify(uint8_t new_state);
