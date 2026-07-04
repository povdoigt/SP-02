#include "led_scheduler.h"

#include <math.h>
#include <string.h>

typedef struct {
	waveform_space_t	*waveform;
	uint32_t			 pause_timeout_ms;
	uint32_t			 seq;				// ordre d'insertion, FIFO intra-priorité
	uint32_t			 pause_start_ms;	// début de l'épisode de pause courant
	uint32_t			 generation;		// incrémentée à chaque réoccupation du slot
	uint8_t				 priority;			// plus haut = prioritaire
	bool				 transparent;
	bool				 in_use;
	bool				 forced_background;	// timeout dépassé : joue en direct, masqué
} led_evt_t;

static led_evt_t	pool[LED_SCHED_MAX_EVENTS];
static uint32_t		next_seq;

#define LED_SCHED_INDEX_BITS 8
#define LED_SCHED_INDEX_MASK 0xFFu

static led_evt_handle_t make_handle(uint8_t index, uint32_t generation) {
	return (led_evt_handle_t)(((uint32_t)generation << LED_SCHED_INDEX_BITS) | index);
}

static uint8_t handle_index(led_evt_handle_t h) {
	return (uint8_t)((uint32_t)h & LED_SCHED_INDEX_MASK);
}

static uint32_t handle_generation(led_evt_handle_t h) {
	return ((uint32_t)h) >> LED_SCHED_INDEX_BITS;
}




/* ===================================================
   HELPERS
   =================================================== */

static float3_t clamp_add(float3_t a, float3_t b) {
	return (float3_t){
		.x = fminf(a.x + b.x, 1.0f),
		.y = fminf(a.y + b.y, 1.0f),
		.z = fminf(a.z + b.z, 1.0f),
	};
}

static bool evt_is_done(const led_evt_t *evt) {
	return !Waveform_IsActive((waveform_generic_t *)evt->waveform);
}

/**
 * Rend l'event visible : le sort de pause si besoin (Waveform_Pause gère
 * la continuité de phase en interne) et joue sa couleur réelle.
 */
static float3_t evt_reveal(led_evt_t *evt, uint32_t now_ms) {
	if (Waveform_IsPaused((waveform_generic_t *)evt->waveform)) {
		Waveform_Pause((waveform_generic_t *)evt->waveform, false, now_ms);
	}
	evt->forced_background = false;
	float3_t color = Waveform_Play_Space(evt->waveform, now_ms);
	if (evt_is_done(evt)) {
		evt->in_use = false;
	}
	return color;
}

/**
 * Gère un event masqué (non joué ce cycle) :
 * - pause_timeout_ms == 0     : jamais mis en pause, toujours joué en
 *                                direct, couleur ignorée.
 * - déjà en exécution forcée  : continue de jouer en direct (le timeout
 *                                a déjà été dépassé une fois).
 * - sinon                     : mis en pause ; passé pause_timeout_ms,
 *                                force-reprend (borné au timeout exact
 *                                via le current_time passé à Waveform_Pause)
 *                                puis bascule en exécution forcée.
 */
static void evt_background(led_evt_t *evt, uint32_t now_ms) {
	if (evt->pause_timeout_ms == 0 || evt->forced_background) {
		(void)Waveform_Play_Space(evt->waveform, now_ms);
		if (evt_is_done(evt)) {
			evt->in_use = false;
		}
		return;
	}

	if (!Waveform_IsPaused((waveform_generic_t *)evt->waveform)) {
		Waveform_Pause((waveform_generic_t *)evt->waveform, true, now_ms);
		evt->pause_start_ms = now_ms;
		return;
	}

	if ((now_ms - evt->pause_start_ms) < evt->pause_timeout_ms) {
		return;	// reste en pause
	}

	// Timeout dépassé : on force la reprise, mais en indiquant que le
	// changement d'état a eu lieu exactement au moment du timeout
	// (pause_start_ms + pause_timeout_ms), pas à now_ms. La phase interne
	// n'est donc décalée que du temps réellement "gelé" avant le forçage ;
	// à partir de maintenant la waveform tourne en temps réel, en silence.
	Waveform_Pause((waveform_generic_t *)evt->waveform, false,
	               evt->pause_start_ms + evt->pause_timeout_ms);
	evt->forced_background = true;
	(void)Waveform_Play_Space(evt->waveform, now_ms);
	if (evt_is_done(evt)) {
		evt->in_use = false;
	}
}

/**
 * Tête FIFO d'un niveau de priorité : event in_use de ce niveau
 * avec le plus petit seq (le plus ancien).
 */
static led_evt_t *fifo_head_of_priority(uint8_t priority) {
	led_evt_t *head = NULL;
	for (int8_t i = 0; i < LED_SCHED_MAX_EVENTS; i++) {
		led_evt_t *evt = &pool[i];
		if (!evt->in_use || evt->priority != priority) {
			continue;
		}
		if (head == NULL || evt->seq < head->seq) {
			head = evt;
		}
	}
	return head;
}




/**
 * Slot in_use de plus basse priorité (tie-break : le plus ancien, seq le
 * plus petit). Retourne -1 si le pool est vide.
 */
static int8_t find_lowest_priority_slot(void) {
	int8_t worst = -1;
	for (int8_t i = 0; i < LED_SCHED_MAX_EVENTS; i++) {
		if (!pool[i].in_use) {
			continue;
		}
		if (worst < 0
		    || pool[i].priority < pool[worst].priority
		    || (pool[i].priority == pool[worst].priority && pool[i].seq < pool[worst].seq)) {
			worst = i;
		}
	}
	return worst;
}




/* ===================================================
   API
   =================================================== */

void LedSched_Init(void) {
	memset(pool, 0, sizeof(pool));
	next_seq = 0;
}

led_evt_handle_t LedSched_Add(waveform_space_t *waveform, uint8_t priority,
                              bool transparent, uint32_t pause_timeout_ms,
                              led_evt_force_t force) {
	if (waveform == NULL) {
		return LED_SCHED_HANDLE_INVALID;
	}

	// 1. Slot libre disponible : pas besoin d'éviction, quel que soit `force`.
	int8_t target = -1;
	for (int8_t i = 0; i < LED_SCHED_MAX_EVENTS; i++) {
		if (!pool[i].in_use) {
			target = i;
			break;
		}
	}

	// 2. Pool plein : appliquer la politique d'éviction.
	if (target < 0) {
		if (force == LED_SCHED_NO_FORCE) {
			return LED_SCHED_HANDLE_INVALID;
		}

		int8_t victim = find_lowest_priority_slot();
		if (victim < 0) {
			return LED_SCHED_HANDLE_INVALID;	// LED_SCHED_MAX_EVENTS == 0
		}
		if (force == LED_SCHED_SOFT_FORCE && pool[victim].priority >= priority) {
			return LED_SCHED_HANDLE_INVALID;	// rien d'assez faible à évincer
		}

		// L'éviction rend le handle de la victime périmé : son détenteur
		// le découvrira au prochain LedSched_Remove() (no-op silencieux),
		// grâce à l'incrément de génération ci-dessous.
		target = victim;
	}

	pool[target].generation++;	// périme tout ancien handle vers ce slot
	uint32_t generation = pool[target].generation;

	pool[target] = (led_evt_t){
		.waveform          = waveform,
		.pause_timeout_ms  = pause_timeout_ms,
		.seq               = next_seq++,
		.pause_start_ms    = 0,
		.generation        = generation,
		.priority          = priority,
		.transparent       = transparent,
		.in_use            = true,
		.forced_background = false,
	};
	Waveform_Restart((waveform_generic_t *)waveform);
	return make_handle((uint8_t)target, generation);
}

void LedSched_Remove(led_evt_handle_t evt) {
	if (evt == LED_SCHED_HANDLE_INVALID) {
		return;
	}
	uint8_t index = handle_index(evt);
	if (index >= LED_SCHED_MAX_EVENTS) {
		return;
	}
	if (!pool[index].in_use || pool[index].generation != handle_generation(evt)) {
		return;	// handle périmé (slot déjà réoccupé) : no-op silencieux
	}
	// Pas de nettoyage nécessaire sur la waveform : le slot sera réécrit
	// intégralement (et Waveform_Restart-é) par le prochain LedSched_Add.
	pool[index].in_use = false;
}

void LedSched_Clear(void) {
	for (int8_t i = 0; i < LED_SCHED_MAX_EVENTS; i++) {
		pool[i].in_use = false;
		pool[i].generation++;	// périme tout ancien handle vers ce slot
	}
}

bool LedSched_IsHandleValid(led_evt_handle_t evt) {
	if (evt == LED_SCHED_HANDLE_INVALID) {
		return false;
	}
	uint8_t index = handle_index(evt);
	if (index >= LED_SCHED_MAX_EVENTS) {
		return false;
	}
	if (!pool[index].in_use || pool[index].generation != handle_generation(evt)) {
		return false;	// handle périmé (slot déjà réoccupé)
	}
	return true;
}

float3_t LedSched_Process(uint32_t now_ms) {
	float3_t	acc = FLOAT3_ZERO;
	bool		visible = true;		// la descente compose tant que true
	bool		have_prev_prio = false;
	uint8_t		prev_prio = 0;

	// Descente des niveaux de priorité, du plus haut au plus bas
	for (;;) {
		bool	found = false;
		uint8_t	prio = 0;
		for (int8_t i = 0; i < LED_SCHED_MAX_EVENTS; i++) {
			led_evt_t *evt = &pool[i];
			if (!evt->in_use) {
				continue;
			}
			if (have_prev_prio && evt->priority >= prev_prio) {
				continue;
			}
			if (!found || evt->priority > prio) {
				prio = evt->priority;
				found = true;
			}
		}
		if (!found) {
			break;
		}
		prev_prio = prio;
		have_prev_prio = true;

		led_evt_t *head = fifo_head_of_priority(prio);

		// Tous les events du niveau sauf la tête : en attente FIFO, masqués
		for (int8_t i = 0; i < LED_SCHED_MAX_EVENTS; i++) {
			led_evt_t *evt = &pool[i];
			if (!evt->in_use || evt->priority != prio || evt == head) {
				continue;
			}
			evt_background(evt, now_ms);
		}

		if (head == NULL) {
			continue;	// niveau vidé par evt_background (done en silence)
		}

		if (visible) {
			bool was_transparent = head->transparent;
			acc = clamp_add(acc, evt_reveal(head, now_ms));
			if (!head->in_use || !was_transparent) {
				visible = false;
			}
		} else {
			evt_background(head, now_ms);
		}
	}

	return acc;
}
