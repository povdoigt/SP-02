#ifndef STAGE_2_H
#define STAGE_2_H

#include "data_topic.h"

#include "full_seq_utils.h"




/* ===================================================
   SETUP FUNCTIONS
   =================================================== */

void setup_stage_2(rocket_state_t *rocket_state);

void setup_servomotors_stage_2(rocket_state_t *rocket_state);
void setup_data_acquisition_stage_2(rocket_state_t *rocket_state);
void setup_attitude_stage_2(rocket_state_t *rocket_state);




/* ===================================================
   LOOP FUNCTIONS
   =================================================== */

void loop_stage_2(rocket_state_t *rocket_state);




/* ===================================================
   SENSOR DATA CALLBACKS
   =================================================== */

void on_new_accel_frame(rocket_state_t *rocket_state, data_sub_t *sub);
void on_new_gyro_frame(rocket_state_t *rocket_state, data_sub_t *sub);
void on_new_pressure_frame(rocket_state_t *rocket_state, data_sub_t *sub);




/* ===================================================
   SECOND STAGE STATE MACHINE
   =================================================== */

typedef enum second_stage_initialisation_phase_t {
	SECOND_STAGE_INIT_PARA_ZERO,
	SECOND_STAGE_INIT_WAIT_PARA_ZERO,
	SECOND_STAGE_INIT_WAIT_STAGE_ASSEMBLY_CONFIRMATION,
	SECOND_STAGE_INIT_WAIT_JACK,
} second_stage_initialisation_phase_t;

typedef enum second_stage_flight_phase_t {
	SECOND_STAGE_FLIGHT_INITIALISATION,
	SECOND_STAGE_FLIGHT_WAIT_LAUNCH_CONFIRMATION,
	SECOND_STAGE_FLIGHT_WAIT_SEPARATION_CONFIRMATION,
	SECOND_STAGE_FLIGHT_WAIT_ATTITUDE_CONFIRMATION,
	SECOND_STAGE_FLIGHT_IGNITION,
	SECOND_STAGE_FLIGHT_WAIT_IGNITION_CONFIRMATION,
	SECOND_STAGE_FLIGHT_WAIT_APOGEE_CONFIRMATION,
	SECOND_STAGE_FLIGHT_APOGEE,
	SECOND_STAGE_FLIGHT_WAIT_DROGUE_CONFIRMATION,
	SECOND_STAGE_FLIGHT_WAIT_LANDING_CONFIRMATION,
	SECOND_STAGE_FLIGHT_IDLE,
} second_stage_flight_phase_t;

void second_stage_init_state_machine(rocket_state_t *rocket_state);
void second_stage_flight_state_machine(rocket_state_t *rocket_state);

#endif // STAGE_2_H
