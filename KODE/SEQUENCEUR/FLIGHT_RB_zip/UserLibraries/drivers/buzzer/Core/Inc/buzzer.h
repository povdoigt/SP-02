#ifndef BUZZER_H
#define BUZZER_H

#include "stm32f0xx_hal.h"

#include <stdbool.h>

#define BUZZER_NOTE_A4 440

#define BUZZER_NGGYU_SONG_SIZE 19
#define BUZZER_CCCP_SONG_SIZE 26
#define BUZZER_BEEP_SONG_SIZE 3


#define BUZZER_ACTIVE 1


typedef struct buzzer_t {
    TIM_HandleTypeDef *htim;
    uint32_t channel;

    bool ASYNC_busy;
} buzzer_t;

typedef struct BUZZER_SONG_BANK {
    float nggyu_freqs[BUZZER_NGGYU_SONG_SIZE];
    uint32_t nggyu_durations[BUZZER_NGGYU_SONG_SIZE];

    float cccp_freqs[BUZZER_CCCP_SONG_SIZE];
    uint32_t cccp_durations[BUZZER_CCCP_SONG_SIZE];

    float beep_0_freqs[BUZZER_BEEP_SONG_SIZE];
    uint32_t beep_0_durations[BUZZER_BEEP_SONG_SIZE];

    float beep_1_freqs[BUZZER_BEEP_SONG_SIZE];
    uint32_t beep_1_durations[BUZZER_BEEP_SONG_SIZE];

    float beep_2_freqs[BUZZER_BEEP_SONG_SIZE];
    uint32_t beep_2_durations[BUZZER_BEEP_SONG_SIZE];

    float beep_3_freqs[BUZZER_BEEP_SONG_SIZE];
    uint32_t beep_3_durations[BUZZER_BEEP_SONG_SIZE];
} BUZZER_SONG_BANK;

extern BUZZER_SONG_BANK buzzer_song_bank;


void BUZZER_set_song_bank(BUZZER_SONG_BANK *song_bank);

void BUZZER_Init(buzzer_t *buzzer, TIM_HandleTypeDef *htim, uint32_t channel);

void BUZZER_play_note(buzzer_t *buzzer, float *freqs, uint32_t *duration, int size);


// #define ASYNC_BUZZER_play_note_NUMBER 10

// typedef enum ASYNC_BUZZER_play_note_STATE {
//     ASYNC_play_note_WAIT_BUZZER,
//     ASYNC_play_note_START,
//     ASYNC_play_note_PLAY,
//     ASYNC_play_note_END,
// } ASYNC_BUZZER_play_note_STATE;

// typedef struct ASYNC_BUZZER_play_note_CONTEXT {
//     BUZZER   *buzzer;
//     float    *freqs;
//     uint32_t *durations;
//     size_t    size;
//     size_t    index;
//     uint32_t  next_time;

//     ASYNC_BUZZER_play_note_STATE state;
// } ASYNC_BUZZER_play_note_CONTEXT;

// extern TASK_POOL_CREATE(ASYNC_BUZZER_play_note);

// void ASYNC_BUZZER_play_note_init(TASK      *self,
//                           BUZZER    *buzzer,
//                           float     *freqs,
//                           uint32_t  *durations,
//                           int        size);
// TASK_RETURN ASYNC_BUZZER_play_note(SCHEDULER *scheduler, TASK *self);

#endif /* BUZZER_H */