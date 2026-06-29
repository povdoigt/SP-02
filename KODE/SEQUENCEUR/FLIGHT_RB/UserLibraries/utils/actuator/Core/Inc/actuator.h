/**
 ******************************************************************************
 * @file    actuator.h
 * @brief   Abstraction layer for mechanical actuators driven by STS3215 servos.
 *
 * Each actuator encapsulates:
 *   - A pointer to its underlying STS_Servo_t
 *   - A static configuration (homing parameters, nominal positions, angle range)
 *   - Runtime state (homing sub-state machine, homed flag)
 *
 * Homing — two usage patterns:
 *
 *   Non-blocking (state-machine friendly):
 *     Call Actuator_HomingStart() once to arm the sequence, then call
 *     Actuator_HomingProcess() on every tick.  It returns a
 *     Actuator_HomingStatus_t that the caller can switch on:
 *
 *       Actuator_HomingStart(&act);
 *       // ... in loop:
 *       switch (Actuator_HomingProcess(&act)) {
 *           case ACTUATOR_HOMING_IN_PROGRESS: break;
 *           case ACTUATOR_HOMING_DONE:        // proceed
 *           case ACTUATOR_HOMING_ERROR:       // handle fault
 *       }
 *
 *   Blocking (simple sequential code):
 *     Actuator_HomingRun(&act, timeout_ms) blocks until done or timeout.
 *
 * After homing, use Actuator_GoToPosition(act, pos_id) to reach any
 * nominal position by index (0 … num_positions-1).
 *
 * Position indices are intentionally generic (POS_0, POS_1, …). The caller
 * defines semantic aliases locally, e.g.:
 *   #define AF_POS_MOUNT   0
 *   #define AF_POS_CLOSED  1
 *   #define AF_POS_OPEN    2
 ******************************************************************************
 */

#ifndef ACTUATOR_H
#define ACTUATOR_H

#include "STS.h"
#include <stdbool.h>
#include <stdint.h>

#ifdef __cplusplus
extern "C" {
#endif


/* ============================================================
   Configuration constants
   ============================================================ */

/** Maximum number of nominal positions per actuator. Increase if needed. */
#define ACTUATOR_MAX_POSITIONS     8U


/* ============================================================
   Types
   ============================================================ */

/**
 * @brief Static configuration for a mechanical actuator.
 *
 * All fields are set once at compile-time and never modified at runtime.
 *
 * @note  homing_calibration_deg:
 *          The logical angle (in degrees) assigned to the physical hard-stop.
 *          Example: set to -10.0f so that all nominal positions in flight are
 *          positive; set to 720.0f for a multi-turn system whose stop is at
 *          the upper end of the travel.
 *
 * @note  positions_deg[]:
 *          Nominal positions reachable via Actuator_GoToPosition().
 *          Indices are 0-based; the caller defines semantic aliases.
 *          Positions need not be adjacent or monotonic.
 *
 * @note  angle_margin_deg:
 *          Margin added symmetrically around the computed [min, max] before
 *          writing CW/CCW limits to the servo EPROM.  The full set of angles
 *          used for the limit calculation is: homing_calibration_deg + all
 *          positions_deg[] entries.
 */
typedef struct {
    float   homing_speed_rpm;           /**< Homing speed (negative = CCW).            */
    float   homing_calibration_deg;     /**< Angle assigned to the hard-stop.           */

    float   positions_deg[ACTUATOR_MAX_POSITIONS]; /**< Nominal positions (deg).  */
    uint8_t num_positions;              /**< Number of valid entries in positions_deg.  */

    float   angle_margin_deg;           /**< Margin added around min/max for EPROM limits. */
} Actuator_Config_t;


/**
 * @brief  Status returned by Actuator_HomingProcess() each tick.
 */
typedef enum {
    ACTUATOR_HOMING_IDLE        = 0,    /**< HomingStart() not yet called.                 */
    ACTUATOR_HOMING_IN_PROGRESS = 1,    /**< Sequence running, call again next tick.       */
    ACTUATOR_HOMING_DONE        = 2,    /**< Hard-stop reached, servo calibrated.          */
    ACTUATOR_HOMING_ERROR       = 3,    /**< Communication error; sequence aborted.        */
} Actuator_HomingStatus_t;


/**
 * @brief  Internal sub-states of the homing sequence.
 *
 * Opaque to the caller; exposed in the header only so Actuator_t can be
 * stack- or statically-allocated without a heap allocation.
 */
typedef enum {
    _HOMING_STATE_IDLE       = 0,
    _HOMING_STATE_START      = 1,   /**< Arm: set speed-control mode + goal speed.    */
    _HOMING_STATE_MOVING     = 2,   /**< Poll overload flag each tick.                */
    _HOMING_STATE_SETTLING   = 3,   /**< Brief settle delay after stop detected.      */
    _HOMING_STATE_CALIBRATE  = 4,   /**< Stop motor, calibrate, switch to pos-ctrl.   */
    _HOMING_STATE_DONE       = 5,
    _HOMING_STATE_ERROR      = 6,
} _Actuator_HomingState_t;


/**
 * @brief Runtime state of a mechanical actuator.
 *
 * Initialise with Actuator_Init(). Never access fields directly.
 */
typedef struct {
    STS_Servo_t                 *servo;         /**< Underlying servo (not owned).     */
    Actuator_Config_t        config;        /**< Copy of the static config.        */
    bool                         is_homed;      /**< True after a successful homing.   */

    /* Homing sub-state machine (internal) */
    _Actuator_HomingState_t  _homing_state;
    uint32_t                     _homing_settle_tick; /**< HAL_GetTick() at stop detect. */
} Actuator_t;


/* ============================================================
   API — Initialisation
   ============================================================ */

/**
 * @brief  Initialise the actuator and configure the servo EPROM angle limits.
 *
 * Copies @p config into the actuator struct, links the servo pointer, then
 * computes the union of homing_calibration_deg and all positions_deg[] entries
 * to derive [min, max].  A margin of ±angle_margin_deg is added and the result
 * is written to the CW/CCW angle limit registers in the servo EPROM, enabling
 * the exact angular range this actuator needs (including negative angles or
 * multi-turn ranges when the geometry requires it).
 *
 * Does NOT perform homing. Call Actuator_HomingRun() or
 * Actuator_HomingStart() + Actuator_HomingProcess() afterwards.
 *
 * @param  act    Actuator to initialise.
 * @param  servo  Pointer to an already-initialised STS_Servo_t.
 * @param  config Pointer to a (typically const) configuration struct.
 * @retval HAL_OK on success, HAL_ERROR / HAL_TIMEOUT on servo comm failure.
 */
HAL_StatusTypeDef Actuator_Init(Actuator_t *act,
                                    STS_Servo_t *servo,
                                    const Actuator_Config_t *config);


/* ============================================================
   API — Homing (non-blocking)
   ============================================================ */

/**
 * @brief  Arm the homing sequence.
 *
 * Must be called once before the first Actuator_HomingProcess() tick.
 * Resets the internal sub-state machine and marks the actuator as not homed.
 *
 * @param  act  Actuator to home.
 */
void Actuator_HomingStart(Actuator_t *act);


/**
 * @brief  Advance the homing sequence by one step (non-blocking tick).
 *
 * Drives the internal sub-state machine one step forward.  Must be called
 * repeatedly (e.g. from a superloop or RTOS task) until it returns
 * ACTUATOR_HOMING_DONE or ACTUATOR_HOMING_ERROR.
 *
 * Sub-states driven internally:
 *   START → MOVING (poll overload) → SETTLING (hardware settle) → CALIBRATE → DONE
 *
 * @param  act  Actuator being homed.
 * @retval ACTUATOR_HOMING_IDLE        HomingStart() was not called.
 * @retval ACTUATOR_HOMING_IN_PROGRESS Still running; call again next tick.
 * @retval ACTUATOR_HOMING_DONE        Homing complete; actuator is calibrated.
 * @retval ACTUATOR_HOMING_ERROR       Communication failure; sequence aborted.
 */
Actuator_HomingStatus_t Actuator_HomingProcess(Actuator_t *act);


/* ============================================================
   API — Homing (blocking)
   ============================================================ */

/**
 * @brief  Run the full homing sequence to completion (blocking).
 *
 * Internally calls Actuator_HomingStart() then loops on
 * Actuator_HomingProcess() until DONE or ERROR, or until
 * @p timeout_ms milliseconds have elapsed.
 *
 * @param  act        Actuator to home.
 * @param  timeout_ms Maximum time to wait (ms). Pass 0 for no timeout.
 * @retval HAL_OK      Homing completed successfully.
 * @retval HAL_TIMEOUT Sequence did not complete within timeout_ms.
 * @retval HAL_ERROR   Communication error during the sequence.
 */
HAL_StatusTypeDef Actuator_HomingRun(Actuator_t *act, uint32_t timeout_ms);


/* ============================================================
   API — Motion
   ============================================================ */

/**
 * @brief  Command the actuator to a nominal position by index.
 *
 * @param  act     Actuator to move.
 * @param  pos_id  Index into config.positions_deg[] (0-based).
 * @retval HAL_OK      on success.
 * @retval HAL_ERROR   if pos_id >= num_positions or actuator not yet homed.
 */
HAL_StatusTypeDef Actuator_GoToPosition(Actuator_t *act, uint8_t pos_id);


/* ============================================================
   API — Telemetry
   ============================================================ */

/**
 * @brief  Read the current raw telemetry from the servo.
 *
 * Thin wrapper around STS_Servo_GetCurrentStatus(); provided here so callers
 * can stay within the Actuator API without touching STS directly.
 *
 * @param  act     Actuator to query.
 * @param  status  Output struct (raw units).
 * @retval HAL_OK on success.
 */
HAL_StatusTypeDef Actuator_GetCurrentStatus(Actuator_t *act,
                                                STS_Servo_Current_raw_t *status);


#ifdef __cplusplus
}
#endif
#endif /* ACTUATOR_H */
