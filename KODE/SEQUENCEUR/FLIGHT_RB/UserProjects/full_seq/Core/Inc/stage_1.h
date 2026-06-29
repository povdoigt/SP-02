#ifndef STAGE_1_H
#define STAGE_1_H

#include "full_seq_utils.h"




/* ===================================================
   SETUP FUNCTIONS
   =================================================== */

void setup_stage_1(rocket_state_t *rocket_state);

void setup_servomotors_stage_1(rocket_state_t *rocket_state);




/* ===================================================
   LOOP FUNCTIONS
   =================================================== */

void loop_stage_1(rocket_state_t *rocket_state);

void setup_uart_buffers_stage_1(rocket_state_t *rocket_state);
void setup_servomotors_stage_1(rocket_state_t *rocket_state);




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

void first_stage_init_state_machine(rocket_state_t *rocket_state);
void first_stage_flight_state_machine(rocket_state_t *rocket_state);

#endif // STAGE_1_H
