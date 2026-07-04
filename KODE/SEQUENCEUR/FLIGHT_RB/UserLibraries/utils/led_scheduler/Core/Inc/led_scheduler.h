#ifndef LED_SCHEDULER_H
#define LED_SCHEDULER_H

#include "waveform_def.h"
#include "float3.h"

#include <stdint.h>
#include <stdbool.h>

#define LED_SCHED_MAX_EVENTS 8

/**
 * Handle vers un event du pool statique.
 * Encode (index, génération) : la génération est incrémentée chaque fois
 * qu'un slot est réoccupé (retrait normal + réutilisation, ou éviction
 * forcée). Un handle vers un slot qui a changé d'occupant redevient
 * automatiquement invalide -- LedSched_Remove() sur un tel handle est un
 * no-op silencieux, il ne touchera jamais le nouvel occupant.
 */
typedef int32_t led_evt_handle_t;
#define LED_SCHED_HANDLE_INVALID ((led_evt_handle_t)-1)

/**
 * Politique d'éviction quand LedSched_Add() est appelé pool plein.
 *
 * LED_SCHED_NO_FORCE   : échoue (LED_SCHED_HANDLE_INVALID), pool inchangé.
 * LED_SCHED_SOFT_FORCE : évince l'event de plus basse priorité du pool,
 *                        seulement si sa priorité est STRICTEMENT plus
 *                        basse que celle du nouvel event. Sinon, échoue
 *                        comme LED_SCHED_NO_FORCE (aucun event moins
 *                        prioritaire à sacrifier).
 * LED_SCHED_HARD_FORCE : évince toujours l'event de plus basse priorité
 *                        du pool, même s'il est plus prioritaire que le
 *                        nouvel event. Ne peut échouer que si
 *                        LED_SCHED_MAX_EVENTS == 0.
 *
 * À priorité égale parmi les candidats à l'éviction, le plus ancien
 * (seq le plus petit) est choisi.
 *
 * Dans tous les cas, le handle de l'event évincé devient invalide : son
 * détenteur le découvrira au prochain LedSched_Remove() (no-op silencieux).
 */
typedef enum {
	LED_SCHED_NO_FORCE,
	LED_SCHED_SOFT_FORCE,
	LED_SCHED_HARD_FORCE,
} led_evt_force_t;

/**
 * Initialise le scheduler (vide le pool, réinitialise les générations).
 */
void LedSched_Init(void);

/**
 * Ajoute un event lumineux au scheduler.
 *
 * waveform         : waveform espace jouée par l'event. La waveform est
 *                    Waveform_Restart()-ée à l'ajout. Une même instance de
 *                    waveform_space_t ne doit pas être partagée entre deux
 *                    events actifs (t0/active sont internes à l'instance).
 * priority         : plus haut = plus prioritaire. À priorité égale, les
 *                    events sont résolus en FIFO (exclusion mutuelle : un
 *                    seul event joué à la fois par niveau de priorité).
 * transparent      : si true, la couleur de l'event s'additionne (bornée à
 *                    [0,1]) avec celle du niveau de priorité inférieur.
 *                    Si false, l'event stoppe la descente (sa couleur est
 *                    tout de même additionnée).
 * pause_timeout_ms : durée maximale de pause quand l'event est masqué.
 *                    Passé ce délai, la waveform est jouée en silence
 *                    (le temps avance, la couleur est ignorée).
 *                    0            = jamais mis en pause, joue toujours.
 *                    HAL_MAX_DELAY = pause sans limite, jamais forcé.
 * force            : politique d'éviction si le pool est plein, voir
 *                    led_evt_force_t. Passer LED_SCHED_NO_FORCE pour le
 *                    comportement standard (échoue si plein).
 *
 * Retourne le handle de l'event, ou LED_SCHED_HANDLE_INVALID si l'ajout
 * (et l'éviction éventuelle) a échoué.
 *
 * Note : une waveform périodique ne se termine jamais d'elle-même, elle
 * doit être retirée explicitement via LedSched_Remove().
 */
led_evt_handle_t LedSched_Add(waveform_space_t *waveform, uint8_t priority,
                              bool transparent, uint32_t pause_timeout_ms,
                              led_evt_force_t force);

/**
 * Retire un event du scheduler.
 * No-op silencieux si le handle est invalide OU périmé (slot réoccupé
 * par un autre event depuis, via retrait+réutilisation ou éviction).
 */
void LedSched_Remove(led_evt_handle_t evt);

/**
 * Efface tous les events du scheduler.
 */
void LedSched_Clear(void);

/**
 * Vérifie si un handle est valide. Un handle périmé (slot réoccupé par un autre event depuis) est
 * considéré comme invalide.
 */
bool LedSched_IsHandleValid(led_evt_handle_t evt);

/**
 * Résout tous les events et retourne la couleur composée à passer à
 * LED_RGB_SetColor(). Non-bloquant, à appeler à chaque loop.
 *
 * now_ms : temps courant (HAL_GetTick()).
 */
float3_t LedSched_Process(uint32_t now_ms);

#endif /* LED_SCHEDULER_H */
