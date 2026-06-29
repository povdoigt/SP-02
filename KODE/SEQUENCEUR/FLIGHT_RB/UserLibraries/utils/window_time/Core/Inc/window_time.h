#ifndef WINDOW_TIME_H
#define WINDOW_TIME_H

#include <stdint.h>

typedef enum window_time_state_t {
    WINDOW_TIME_STATE_WAITING,
    WINDOW_TIME_STATE_ACTIVE,
    WINDOW_TIME_STATE_EXPIRED
} window_time_state_t;

typedef struct window_time_t {
    uint32_t start_time_ms; /* Timestamp de début de la fenêtre en millisecondes */
    uint32_t duration_ms;   /* Durée de la fenêtre en millisecondes */
    uint8_t id;                /* ID de la fenêtre pour l'identifier dans les événements */
} window_time_t;

void window_time_init(window_time_t *window, uint8_t id, uint32_t start_time_ms, uint32_t duration_ms);
window_time_state_t window_time_get_state(const window_time_t *window, uint32_t current_time_ms);

#endif /* WINDOW_TIME_H */