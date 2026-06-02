
#include "buzzer.h"
#include "tim.h"

#include <string.h>
#include <math.h>


// Song generation

BUZZER_SONG_BANK buzzer_song_bank;

typedef enum note_t {
    NOTE_NOP = -10,
    NOTE_C   = -9,
    NOTE_Cs  = -8,
    NOTE_D   = -7,
    NOTE_Ds  = -6,
    NOTE_E   = -5,
    NOTE_F   = -4,
    NOTE_Fs  = -3,
    NOTE_G   = -2,
    NOTE_Gs  = -1,
    NOTE_A   =  0,   // Central note (A4 = 440 Hz)
    NOTE_As  = +1,
    NOTE_B   = +2
} note_t;

typedef enum octave_t {
    OCTAVE_0 = -4,
    OCTAVE_1 = -3,
    OCTAVE_2 = -2,
    OCTAVE_3 = -1,
    OCTAVE_4 =  0,  // Central octave (A4 = 440 Hz)
    OCTAVE_5 = +1,
    OCTAVE_6 = +2,
    OCTAVE_7 = +3,
    OCTAVE_8 = +4
} octave_t;

float get_note_freq(note_t note, octave_t octave);

void set_nggyu_song(float *freqs, uint32_t *durations);
void set_cccp_song(float *freqs, uint32_t *durations);
void set_beep_0_song(float *freqs, uint32_t *durations);
void set_beep_1_song(float *freqs, uint32_t *durations);
void set_beep_2_song(float *freqs, uint32_t *durations);
void set_beep_3_song(float *freqs, uint32_t *durations);


void BUZZER_set_song_bank(BUZZER_SONG_BANK *song_bank) {
    set_nggyu_song(song_bank->nggyu_freqs, song_bank->nggyu_durations);
    set_cccp_song(song_bank->cccp_freqs, song_bank->cccp_durations);
    set_beep_0_song(song_bank->beep_0_freqs, song_bank->beep_0_durations);
    set_beep_1_song(song_bank->beep_1_freqs, song_bank->beep_1_durations);
    set_beep_2_song(song_bank->beep_2_freqs, song_bank->beep_2_durations);
    set_beep_3_song(song_bank->beep_3_freqs, song_bank->beep_3_durations);
}


void BUZZER_Init(buzzer_t *buzzer, TIM_HandleTypeDef *htim, uint32_t channel) {
    buzzer->htim = htim;
    buzzer->channel = channel;
    buzzer->ASYNC_busy = false;

    uint32_t periode = htim->Init.Period + 1;

    __HAL_TIM_SET_COMPARE(htim, channel, periode / 2 - 1); // Set duty cycle to 50%
}

#if BUZZER_ACTIVE
void BUZZER_play_note(buzzer_t *buzzer, float* freqs, uint32_t* duration, int size) {
    TIM_set_frequency(buzzer->htim, 0);
    HAL_TIM_PWM_Start(buzzer->htim, buzzer->channel);
    for (int i = 0; i < size; i++) {
        TIM_set_frequency(buzzer->htim, freqs[i]);
        HAL_Delay(duration[i]);
    }
    TIM_set_frequency(buzzer->htim, 0);
    HAL_TIM_PWM_Stop(buzzer->htim, buzzer->channel);
}
#else
void BUZZER_play_note(buzzer_t *buzzer, float* freqs, uint32_t* duration, int size) {
    // Do nothing
}
#endif // BUZZER_ACTIVE







// TASK_POOL_CREATE(ASYNC_BUZZER_play_note);

// void ASYNC_BUZZER_play_note_init(TASK      *self,
//                           BUZZER    *buzzer,
//                           float     *freqs,
//                           uint32_t  *durations,
//                           int        size) {
//     ASYNC_BUZZER_play_note_CONTEXT *context = (ASYNC_BUZZER_play_note_CONTEXT*)self->context;

//     context->buzzer = buzzer;
//     context->freqs = freqs;
//     context->durations = durations;
//     context->size = size;
//     context->index = 0;
//     context->next_time = 0;
//     context->state = ASYNC_play_note_WAIT_BUZZER;
// }

// #if BUZZER_ACTIVE

// TASK_RETURN ASYNC_BUZZER_play_note(SCHEDULER *scheduler, TASK *self) {
//     ASYNC_BUZZER_play_note_CONTEXT *context = (ASYNC_BUZZER_play_note_CONTEXT*)self->context;

//     UNUSED(scheduler);

//     uint32_t current_time = HAL_GetTick();

//     switch (context->state) {
//     case ASYNC_play_note_WAIT_BUZZER: {
//         if (!(context->buzzer->ASYNC_busy) && (context->buzzer->htim->State == HAL_TIM_STATE_READY)) {
//             context->state = ASYNC_play_note_START;
//             context->buzzer->ASYNC_busy = true;
//         }
//         break; }
//     case ASYNC_play_note_START: {
//         TIM_set_frequency(context->buzzer->htim, 0);
//         HAL_TIM_PWM_Start(context->buzzer->htim, context->buzzer->channel);
//         context->state = ASYNC_play_note_PLAY;
//         break; }
//     case ASYNC_play_note_PLAY: {
//         if (current_time >= context->next_time) {
//             if (context->index < context->size) {
//                 TIM_set_frequency(context->buzzer->htim, context->freqs[context->index]);
//                 context->next_time = HAL_GetTick() + context->durations[context->index];
//                 context->index++;
//             } else {
//                 TIM_set_frequency(context->buzzer->htim, 0);
//                 HAL_TIM_PWM_Stop(context->buzzer->htim, context->buzzer->channel);
//                 context->state = ASYNC_play_note_END;
//             }
//         }
//         break; }
//     case ASYNC_play_note_END: {
//         context->buzzer->ASYNC_busy = false;
//         return TASK_RETURN_STOP;
//         break; }
//     }
//     return TASK_RETURN_IDLE;
// }

// #else

// TASK_RETURN ASYNC_BUZZER_play_note(SCHEDULER *scheduler, TASK *self) {
//     return TASK_RETURN_STOP;
// }

// #endif // BUZZER_ACTIVE


float get_note_freq(note_t note, octave_t octave) {
    if (note == NOTE_NOP) {
        return 0;
    }
    return BUZZER_NOTE_A4 * pow(2.0, ((float)note / (float)12) + (float)octave);
}

void set_nggyu_song(float *freqs, uint32_t *durations) {
    float _freqs[BUZZER_NGGYU_SONG_SIZE] = {
        get_note_freq(NOTE_C, OCTAVE_4),
        get_note_freq(NOTE_D, OCTAVE_4),
        get_note_freq(NOTE_G, OCTAVE_3),
        get_note_freq(NOTE_D, OCTAVE_4),
        get_note_freq(NOTE_E, OCTAVE_4),
        get_note_freq(NOTE_G, OCTAVE_4),
        get_note_freq(NOTE_F, OCTAVE_4),
        get_note_freq(NOTE_E, OCTAVE_4),
        get_note_freq(NOTE_C, OCTAVE_4),
        get_note_freq(NOTE_D, OCTAVE_4),
        get_note_freq(NOTE_G, OCTAVE_3),
        get_note_freq(NOTE_NOP, 0),
        get_note_freq(NOTE_G, OCTAVE_3),
        get_note_freq(NOTE_A, OCTAVE_3),

        get_note_freq(NOTE_C, OCTAVE_4),
        get_note_freq(NOTE_D, OCTAVE_4),
        get_note_freq(NOTE_G, OCTAVE_3),
        get_note_freq(NOTE_D, OCTAVE_4),
        get_note_freq(NOTE_E, OCTAVE_4),
    };

    uint32_t _durations[BUZZER_NGGYU_SONG_SIZE] = {
        1000,
        1000,
         500,
        1000,
        1000,
         250,
         250,
         250,
        1000,
        1000,
         500,
        1000,
         250,
         250,

        1000,
        1000,
         500,
        1000,
        1000,
    };

    for (int i = 0; i < BUZZER_NGGYU_SONG_SIZE; i++) {
        freqs[i] = _freqs[i];
        durations[i] = _durations[i];
    }
}

void set_cccp_song(float *freqs, uint32_t *durations) {
    float _freqs[BUZZER_CCCP_SONG_SIZE] = {
        get_note_freq(NOTE_As, OCTAVE_3),
        get_note_freq(NOTE_Ds, OCTAVE_4),
        get_note_freq(NOTE_As, OCTAVE_3),
        get_note_freq(NOTE_C , OCTAVE_4),
        get_note_freq(NOTE_D , OCTAVE_4),
        get_note_freq(NOTE_G , OCTAVE_3),
        0,
        get_note_freq(NOTE_G , OCTAVE_3),
        get_note_freq(NOTE_C , OCTAVE_4),
        get_note_freq(NOTE_As, OCTAVE_3),
        get_note_freq(NOTE_Gs, OCTAVE_3),
        get_note_freq(NOTE_As, OCTAVE_3),
        get_note_freq(NOTE_Ds, OCTAVE_3),
        get_note_freq(NOTE_F , OCTAVE_3),
        0,
        get_note_freq(NOTE_F , OCTAVE_3),
        get_note_freq(NOTE_G , OCTAVE_3),
        get_note_freq(NOTE_Gs, OCTAVE_3),
        0,
        get_note_freq(NOTE_Gs, OCTAVE_3),
        get_note_freq(NOTE_As, OCTAVE_3),
        get_note_freq(NOTE_C , OCTAVE_4),
        0,
        get_note_freq(NOTE_D , OCTAVE_4),
        get_note_freq(NOTE_Ds, OCTAVE_4),
        get_note_freq(NOTE_F , OCTAVE_4),
    };

    uint32_t _durations[BUZZER_CCCP_SONG_SIZE] = {
         500,
        1000,
         750,
         250,
         750,
         475,
          50,
         475,
         750,
         750,
         250,
         750,
         750,
         725,
          50,
         475,
         500,
         725,
          50,
         475,
         500,
         725,
          50,
         475,
         500,
         750,
    };

    for (int i = 0; i < BUZZER_CCCP_SONG_SIZE; i++) {
        freqs[i] = _freqs[i];
        durations[i] = _durations[i];
    }
}

void set_beep_0_song(float *freqs, uint32_t *durations) {
    float _freqs[BUZZER_BEEP_SONG_SIZE] = {
        get_note_freq(NOTE_C, OCTAVE_4),
        get_note_freq(NOTE_E, OCTAVE_4),
        get_note_freq(NOTE_G, OCTAVE_4),
    };

    uint32_t _durations[BUZZER_BEEP_SONG_SIZE] = {
        500,
        300,
        500,
    };

    for (int i = 0; i < BUZZER_BEEP_SONG_SIZE; i++) {
        freqs[i] = _freqs[i];
        durations[i] = _durations[i];
    }
}

void set_beep_1_song(float *freqs, uint32_t *durations) {
    float _freqs[BUZZER_BEEP_SONG_SIZE] = {
        get_note_freq(NOTE_E, OCTAVE_4),
        0,
        get_note_freq(NOTE_E, OCTAVE_4),
    };

    uint32_t _durations[BUZZER_BEEP_SONG_SIZE] = {
        250,
        250,
        250,
    };

    for (int i = 0; i < BUZZER_BEEP_SONG_SIZE; i++) {
        freqs[i] = _freqs[i];
        durations[i] = _durations[i];
    }
}

void set_beep_2_song(float *freqs, uint32_t *durations) {
    float _freqs[BUZZER_BEEP_SONG_SIZE] = {
        get_note_freq(NOTE_E, OCTAVE_4),
        get_note_freq(NOTE_G, OCTAVE_4),
        get_note_freq(NOTE_G, OCTAVE_4),
    };

    uint32_t _durations[BUZZER_BEEP_SONG_SIZE] = {
        250,
        250,
        250,
    };

    for (int i = 0; i < BUZZER_BEEP_SONG_SIZE; i++) {
        freqs[i] = _freqs[i];
        durations[i] = _durations[i];
    }
}

void set_beep_3_song(float *freqs, uint32_t *durations) {
    float _freqs[BUZZER_BEEP_SONG_SIZE] = {
        get_note_freq(NOTE_G, OCTAVE_5),
        0,
        get_note_freq(NOTE_G, OCTAVE_5),
    };

    uint32_t _durations[BUZZER_BEEP_SONG_SIZE] = {
        250,
        250,
        250,
    };

    for (int i = 0; i < BUZZER_BEEP_SONG_SIZE; i++) {
        freqs[i] = _freqs[i];
        durations[i] = _durations[i];
    }
}
