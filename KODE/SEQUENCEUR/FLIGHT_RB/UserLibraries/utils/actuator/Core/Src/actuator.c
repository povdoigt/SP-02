/**
 ******************************************************************************
 * @file    actuator.c
 * @brief   Implementation of the mechanical actuator abstraction layer.
 ******************************************************************************
 */

#include "actuator.h"
#include "STS.h"

#include <string.h>

/** Settling time (ms) between stop detection and position calibration. */
#define HOMING_SETTLE_MS    100U


/* ============================================================
   Private helpers
   ============================================================ */

static inline uint16_t _deg_to_units(float deg)
{
    return STS_GetPositionInUnits(deg);
}

/**
 * @brief  Compute CW/CCW angle limits from config and write them to EPROM.
 *
 * Scans homing_calibration_deg and all positions_deg[] to find [min, max],
 * expands by ±angle_margin_deg, then writes CW = min and CCW = max to the
 * servo EPROM (same sign-bit encoding as goal position registers).
 *
 * Sequence: unlock EPROM → write CW → write CCW → lock EPROM.
 */
static HAL_StatusTypeDef _compute_and_set_angle_limits(STS_Servo_t *servo,
                                                        const Actuator_Config_t *config)
{
    float min_deg = config->homing_calibration_deg;
    float max_deg = config->homing_calibration_deg;

    for (uint8_t i = 0; i < config->num_positions; i++) {
        float p = config->positions_deg[i];
        if (p < min_deg) { min_deg = p; }
        if (p > max_deg) { max_deg = p; }
    }

    min_deg -= config->angle_margin_deg;
    max_deg += config->angle_margin_deg;

    HAL_StatusTypeDef res;

    uint8_t lock = 0;
    res = STS_Servo_WriteRegister(servo, STS_REG_LOCK_SYMBOL, &lock, 1);
    if (res != HAL_OK) { return res; }
    HAL_Delay(5);

    uint16_t cw_units = _deg_to_units(min_deg);
    uint8_t  cw[2]    = { (uint8_t)(cw_units & 0xFF), (uint8_t)(cw_units >> 8) };
    res = STS_Servo_WriteRegister(servo, STS_REG_CW_ANGLE_LIMIT_L, cw, 2);
    if (res != HAL_OK) { return res; }

    uint16_t ccw_units = _deg_to_units(max_deg);
    uint8_t  ccw[2]    = { (uint8_t)(ccw_units & 0xFF), (uint8_t)(ccw_units >> 8) };
    res = STS_Servo_WriteRegister(servo, STS_REG_CCW_ANGLE_LIMIT_L, ccw, 2);
    if (res != HAL_OK) { return res; }

    HAL_Delay(5);

    lock = 1;
    res = STS_Servo_WriteRegister(servo, STS_REG_LOCK_SYMBOL, &lock, 1);
    return res;
}


/* ============================================================
   Public API — Initialisation
   ============================================================ */

HAL_StatusTypeDef Actuator_Init(Actuator_t *act,
                                    STS_Servo_t *servo,
                                    const Actuator_Config_t *config)
{
    act->servo              = servo;
    act->is_homed           = false;
    act->_homing_state      = _HOMING_STATE_IDLE;
    act->_homing_settle_tick = 0;
    memcpy(&act->config, config, sizeof(Actuator_Config_t));

    return _compute_and_set_angle_limits(servo, config);
}


/* ============================================================
   Public API — Homing (non-blocking)
   ============================================================ */

void Actuator_HomingStart(Actuator_t *act)
{
    act->is_homed      = false;
    act->_homing_state = _HOMING_STATE_START;
}


Actuator_HomingStatus_t Actuator_HomingProcess(Actuator_t *act)
{
    HAL_StatusTypeDef res;

    switch (act->_homing_state) {

        case _HOMING_STATE_IDLE: {
            return ACTUATOR_HOMING_IDLE;
        }

        case _HOMING_STATE_START: {
            res = STS_Servo_SetOperatingMode(act->servo, STS_OP_MODE_SPEED_CONTROL);
            if (res != HAL_OK) { goto error; }

            int16_t speed_units = STS_GetSpeedInUnits(act->config.homing_speed_rpm);
            res = STS_Servo_SetGoalSpeed(act->servo, speed_units);
            if (res != HAL_OK) { goto error; }

            act->_homing_state = _HOMING_STATE_MOVING;
            return ACTUATOR_HOMING_IN_PROGRESS;
        }

        case _HOMING_STATE_MOVING: {
            bool in_overload = false;
            res = STS_Servo_InOverload(act->servo, &in_overload);
            if (res != HAL_OK) { goto error; }

            if (in_overload) {
                act->_homing_settle_tick = HAL_GetTick();
                act->_homing_state       = _HOMING_STATE_SETTLING;
            }
            return ACTUATOR_HOMING_IN_PROGRESS;
        }

        case _HOMING_STATE_SETTLING: {
            /* Non-blocking settle: check elapsed time each tick */
            if (HAL_GetTick() - act->_homing_settle_tick >= HOMING_SETTLE_MS) {
                act->_homing_state = _HOMING_STATE_CALIBRATE;
            }
            return ACTUATOR_HOMING_IN_PROGRESS;
        }

        case _HOMING_STATE_CALIBRATE: {
            /* Stop motor */
            res = STS_Servo_SetGoalSpeed(act->servo, 0);
            if (res != HAL_OK) { goto error; }

            /* Assign logical origin to the physical hard-stop */
            uint16_t calib_units = _deg_to_units(act->config.homing_calibration_deg);
            res = STS_Servo_PositionCalibration(act->servo, calib_units);
            if (res != HAL_OK) { goto error; }

            /* Return to position-control mode */
            res = STS_Servo_SetOperatingMode(act->servo, STS_OP_MODE_POSITION_CONTROL);
            if (res != HAL_OK) { goto error; }

            act->is_homed      = true;
            act->_homing_state = _HOMING_STATE_DONE;
            return ACTUATOR_HOMING_DONE;
        }

        case _HOMING_STATE_DONE: {
            return ACTUATOR_HOMING_DONE;
        }

        case _HOMING_STATE_ERROR:
        default: {
            return ACTUATOR_HOMING_ERROR;
        }
    }

error:
    act->_homing_state = _HOMING_STATE_ERROR;
    return ACTUATOR_HOMING_ERROR;
}


/* ============================================================
   Public API — Homing (blocking)
   ============================================================ */

HAL_StatusTypeDef Actuator_HomingRun(Actuator_t *act, uint32_t timeout_ms)
{
    Actuator_HomingStart(act);

    uint32_t t_start = HAL_GetTick();

    Actuator_HomingStatus_t status;
    do {
        status = Actuator_HomingProcess(act);

        if (timeout_ms != 0 && (HAL_GetTick() - t_start) >= timeout_ms) {
            /* Abort: cut speed and leave servo in a safe state */
            STS_Servo_SetGoalSpeed(act->servo, 0);
            STS_Servo_SetOperatingMode(act->servo, STS_OP_MODE_POSITION_CONTROL);
            act->_homing_state = _HOMING_STATE_ERROR;
            return HAL_TIMEOUT;
        }
    } while (status == ACTUATOR_HOMING_IN_PROGRESS);

    return (status == ACTUATOR_HOMING_DONE) ? HAL_OK : HAL_ERROR;
}


/* ============================================================
   Public API — Motion
   ============================================================ */

HAL_StatusTypeDef Actuator_GoToPosition(Actuator_t *act, uint8_t pos_id)
{
    if (!act->is_homed) {
        return HAL_ERROR;
    }
    if (pos_id >= act->config.num_positions) {
        return HAL_ERROR;
    }

    uint16_t pos_units = _deg_to_units(act->config.positions_deg[pos_id]);
    return STS_Servo_SetGoalPosition(act->servo, pos_units);
}


/* ============================================================
   Public API — Telemetry
   ============================================================ */

HAL_StatusTypeDef Actuator_GetCurrentStatus(Actuator_t *act,
                                                STS_Servo_Current_raw_t *status)
{
    return STS_Servo_GetCurrentStatus(act->servo, status);
}
