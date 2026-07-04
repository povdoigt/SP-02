#ifndef STAGE_1_H
#define STAGE_1_H

#include "full_seq_utils.h"


/* ===================================================
   BOARD FUNCTIONS
   =================================================== */

board_func_state_t board_func_1_stage_1(rocket_state_t *rocket_state);
board_func_state_t board_func_2_stage_1(rocket_state_t *rocket_state);
board_func_state_t board_func_3_stage_1(rocket_state_t *rocket_state);
board_func_state_t board_func_4_stage_1(rocket_state_t *rocket_state);
board_func_state_t board_func_5_stage_1(rocket_state_t *rocket_state);
board_func_state_t board_func_6_stage_1(rocket_state_t *rocket_state);
board_func_state_t board_func_7_stage_1(rocket_state_t *rocket_state);
board_func_state_t board_func_8_stage_1(rocket_state_t *rocket_state);
board_func_state_t board_func_9_stage_1(rocket_state_t *rocket_state);
board_func_state_t board_func_10_stage_1(rocket_state_t *rocket_state);
board_func_state_t board_func_11_stage_1(rocket_state_t *rocket_state);
board_func_state_t board_func_12_stage_1(rocket_state_t *rocket_state);
board_func_state_t board_func_13_stage_1(rocket_state_t *rocket_state);
board_func_state_t board_func_14_stage_1(rocket_state_t *rocket_state);
board_func_state_t board_func_15_stage_1(rocket_state_t *rocket_state);


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

typedef enum sepa_assembly_phase_t {
	SEPA_ASSEMBLY_UNLOCK_STAGE,
	SEPA_ASSEMBLY_WAIT_STAGE_CONNECTION,
	SEPA_ASSEMBLY_WAIT_STAGE_CONNECTION_STABLE,
	SEPA_ASSEMBLY_WAIT_LOCK_STAGE,
	SEPA_ASSEMBLY_WAIT_LOCK_STAGE_STABLE,
} sepa_assembly_phase_t;

typedef enum first_stage_initialisation_phase_t {
	FIRST_STAGE_INIT_WAIT_JACK_READY,

	FIRST_STAGE_INIT_AF_ZERO,
	FIRST_STAGE_INIT_PARA_ZERO,
	FIRST_STAGE_INIT_SEPA_ZERO,

	FIRST_STAGE_INIT_WAIT_ALL_GOOD,
	FIRST_STAGE_INIT_WAIT_ALL_GOOD_STABLE,

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
