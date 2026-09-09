#include "window_time.h"

void window_time_init(window_time_t *window, uint8_t id, uint32_t start_time_ms, uint32_t duration_ms) {
    window->id = id;
    window->start_time_ms = start_time_ms;
    window->duration_ms = duration_ms;
}

window_time_state_t window_time_get_state(const window_time_t *window, uint32_t current_time_ms) {
    if (current_time_ms < window->start_time_ms) {
        return WINDOW_TIME_STATE_WAITING;
    }
    if (current_time_ms < window->start_time_ms + window->duration_ms) {
        return WINDOW_TIME_STATE_ACTIVE;
    }
    return WINDOW_TIME_STATE_EXPIRED;
}