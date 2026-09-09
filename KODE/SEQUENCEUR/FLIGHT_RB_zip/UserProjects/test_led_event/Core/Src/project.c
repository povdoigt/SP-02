/**
 * example_full_demo.c — Démo complète du led_scheduler sous forme setup()/loop().
 *
 * Scénario auto-déroulant (piloté par HAL_GetTick), à flasher tel quel pour
 * observer sur la LED RGB tous les mécanismes du scheduler :
 *
 *   t (ms)   Événement                                Mécanisme démontré
 *   ------   --------------------------------------   -------------------------------
 *    0       Fond phase A : respiration violette      Event de fond opaque, plancher
 *                avec plancher lumineux                 lumineux via scale_add_const
 *    5000    Fond phase B : clignotement jaune         Swap de fond (Remove + Add),
 *                "écrêté" (sinus × porte)               scale_mult (clipping)
 *    8000    Flash de confirmation vert (one-shot)     Event à durée finie, fire-and-
 *                décroissance exponentielle              forget (pas de Remove requis)
 *    10000   3 blips cyan en rafale                    FIFO intra-priorité : joués
 *                                                        l'un après l'autre
 *    13000   Overlay batterie faible (TRANSPARENT)     Composition additive par-dessus
 *                pulse orange + teinte constante         le fond, space_add_const
 *    16000   Fond phase C : "charge" (dent de scie     space_add : deux waveforms
 *                verte + plancher bleu)                  espace additionnées
 *    19000   ERREUR CRITIQUE : heartbeat rouge         Priorité max opaque : masque
 *                (double battement)                      tout, timeout=0 (temps réel)
 *    20000   2 blips pendant l'erreur                  Timeout de pause : masqués,
 *                                                        finissent en silence, jamais vus
 *    24000   Fin de l'erreur                           Le fond + batterie réapparaissent
 *    26000   6 blips d'un coup -> pool PLEIN (8/8)     Saturation du pool
 *    26200   Tentative SOFT_FORCE à priorité 1        Échec propre : aucun candidat
 *                                                        strictement plus faible
 *    26400   RESET D'URGENCE blanc en HARD_FORCE       Éviction forcée : le fond
 *                (triple flash)                          (plus basse priorité) est évincé
 *    29500   Ré-ajout du fond phase C                  Récupération après éviction ;
 *                                                        les blips en attente défilent
 *
 * Priorités : fond=1, blips=2, confirmation=3, batterie=5 (transparent),
 *             erreur=9, urgence=10.
 */

#include "main.h"

#include "led.h"
#include "led_scheduler.h"
#include "waveform.h"

#include <stdbool.h>
#include <stdint.h>




/* ===================================================
   RESSOURCES STATIQUES
   =================================================== */

static led_rgb_t led_rgb;

/* Une instance de waveform_space_t par event potentiellement actif en même
 * temps (contrainte de la lib : t0/active vivent dans l'instance). Les blips
 * tournent sur un pool de 6 instances, réutilisées uniquement une fois
 * l'event précédent terminé et retiré. */
static waveform_space_t wf_base;
static waveform_space_t wf_confirm;
static waveform_space_t wf_battery;
static waveform_space_t wf_error;
static waveform_space_t wf_emergency;
static waveform_space_t wf_blip[6];
static uint8_t          blip_next;   /* rotation dans wf_blip[] */

static led_evt_handle_t evt_base    = LED_SCHED_HANDLE_INVALID;
static led_evt_handle_t evt_battery = LED_SCHED_HANDLE_INVALID;
static led_evt_handle_t evt_error   = LED_SCHED_HANDLE_INVALID;

static uint32_t t0_demo;




/* ===================================================
   WAVEFORMS — PHASE A : respiration violette à plancher
   ===================================================
   Composition sur 3 niveaux :
     sine ──> ×0.85 (mult_const) ──> +0.15 (add_const) ──> × violet
   Le plancher garantit que la LED ne s'éteint jamais complètement :
   l'intensité oscille entre 0.15 et 1.0. */

static const waveform_scale_mult_const_t phase_a_scaled_ctx = {
	.wave_function_mult = waveform_sine,
	.ctx_mult = &(waveform_sine_t){ .start = 0.0f, .end = 1.0f },
	.factor = 0.85f,
};

static const waveform_scale_add_const_t phase_a_floor_ctx = {
	.wave_function_add = waveform_scale_mult_const,
	.ctx_add = &phase_a_scaled_ctx,
	.offset = 0.15f,
};

static const waveform_space_mult_const_t phase_a_ctx = {
	.wave_function_mult = waveform_scale_add_const,
	.ctx_mult = &phase_a_floor_ctx,
	.factor = { .x = 1.0f, .y = 0.0f, .z = 1.0f },   /* violet */
};




/* ===================================================
   WAVEFORMS — PHASE B : clignotement jaune écrêté
   ===================================================
   scale_mult : sinus pleine période × porte [0.35, 0.85].
   La porte "découpe" la respiration : allumage/extinction brutaux
   avec un sommet arrondi — impossible avec une seule fonction de base. */

static const waveform_scale_mult_t phase_b_clip_ctx = {
	.wave_function_mult_1 = waveform_sine,
	.ctx_mult_1 = &(waveform_sine_t){ .start = 0.0f, .end = 1.0f },
	.wave_function_mult_2 = waveform_gate,
	.ctx_mult_2 = &(waveform_gate_t){ .start = 0.35f, .end = 0.85f },
};

static const waveform_space_mult_const_t phase_b_ctx = {
	.wave_function_mult = waveform_scale_mult,
	.ctx_mult = &phase_b_clip_ctx,
	.factor = { .x = 1.0f, .y = 0.8f, .z = 0.0f },   /* jaune */
};




/* ===================================================
   WAVEFORMS — PHASE C : charge (dent de scie verte + plancher bleu)
   ===================================================
   space_add : deux waveforms ESPACE additionnées.
   - rampe [0,1] × vert   : montée progressive, retombée sèche ("charge")
   - porte [0,1] × bleu   : plancher constant, la LED tire vers le teal */

static const waveform_space_mult_const_t phase_c_ramp_ctx = {
	.wave_function_mult = waveform_ramp,
	.ctx_mult = &(waveform_ramp_t){ .start = 0.0f, .end = 1.0f },
	.factor = { .x = 0.0f, .y = 1.0f, .z = 0.0f },   /* vert */
};

static const waveform_space_mult_const_t phase_c_floor_ctx = {
	.wave_function_mult = waveform_gate,
	.ctx_mult = &(waveform_gate_t){ .start = 0.0f, .end = 1.0f },
	.factor = { .x = 0.0f, .y = 0.0f, .z = 0.12f },  /* plancher bleu */
};

static const waveform_space_add_t phase_c_ctx = {
	.wave_function_add_1 = waveform_space_mult_const,
	.ctx_add_1 = &phase_c_ramp_ctx,
	.wave_function_add_2 = waveform_space_mult_const,
	.ctx_add_2 = &phase_c_floor_ctx,
};




/* ===================================================
   WAVEFORM — Flash de confirmation (one-shot)
   ===================================================
   Impulsion exponentielle : allumage instantané, décroissance douce.
   Non périodique : l'event se termine tout seul (fire-and-forget). */

static const waveform_space_mult_const_t confirm_ctx = {
	.wave_function_mult = waveform_impulse,
	.ctx_mult = &(waveform_impulse_t){ .start = 0.0f, .end = 1.0f, .k = 4.0f },
	.factor = { .x = 0.2f, .y = 1.0f, .z = 0.2f },   /* vert clair */
};




/* ===================================================
   WAVEFORM — Blip diagnostic (one-shot, double battement)
   ===================================================
   scale_add de deux portes : deux flashs courts par event. */

static const waveform_scale_add_t blip_beats_ctx = {
	.wave_function_add_1 = waveform_gate,
	.ctx_add_1 = &(waveform_gate_t){ .start = 0.0f, .end = 0.2f },
	.wave_function_add_2 = waveform_gate,
	.ctx_add_2 = &(waveform_gate_t){ .start = 0.4f, .end = 0.6f },
};

static const waveform_space_mult_const_t blip_ctx = {
	.wave_function_mult = waveform_scale_add,
	.ctx_mult = &blip_beats_ctx,
	.factor = { .x = 0.0f, .y = 0.9f, .z = 1.0f },   /* cyan */
};




/* ===================================================
   WAVEFORM — Batterie faible (overlay TRANSPARENT)
   ===================================================
   space_add_const : pulse triangulaire orange + teinte constante ambrée.
   L'offset constant assure une trace visuelle permanente de l'alerte,
   même entre deux pulses ; le tout s'ADDITIONNE au fond (transparent). */

static const waveform_space_mult_const_t battery_pulse_ctx = {
	.wave_function_mult = waveform_triangle,
	.ctx_mult = &(waveform_triangle_t){ .start = 0.0f, .end = 0.18f },
	.factor = { .x = 0.55f, .y = 0.18f, .z = 0.0f }, /* orange */
};

static const waveform_space_add_const_t battery_ctx = {
	.wave_function_add = waveform_space_mult_const,
	.ctx_add = &battery_pulse_ctx,
	.offset = { .x = 0.05f, .y = 0.015f, .z = 0.0f }, /* teinte ambrée permanente */
};




/* ===================================================
   WAVEFORM — Erreur critique (heartbeat rouge)
   ===================================================
   Double battement rapproché, silence, puis répétition (périodique). */

static const waveform_scale_add_t error_beats_ctx = {
	.wave_function_add_1 = waveform_gate,
	.ctx_add_1 = &(waveform_gate_t){ .start = 0.00f, .end = 0.08f },
	.wave_function_add_2 = waveform_gate,
	.ctx_add_2 = &(waveform_gate_t){ .start = 0.16f, .end = 0.24f },
};

static const waveform_space_mult_const_t error_ctx = {
	.wave_function_mult = waveform_scale_add,
	.ctx_mult = &error_beats_ctx,
	.factor = { .x = 1.0f, .y = 0.0f, .z = 0.0f },   /* rouge */
};




/* ===================================================
   WAVEFORM — Reset d'urgence (triple flash blanc, one-shot)
   ===================================================
   scale_add IMBRIQUÉ : add(porte1, add(porte2, porte3)).
   Trois flashs blancs répartis sur la durée de l'event. */

static const waveform_scale_add_t emergency_tail_ctx = {
	.wave_function_add_1 = waveform_gate,
	.ctx_add_1 = &(waveform_gate_t){ .start = 2.0f/6.0f, .end = 3.0f/6.0f },
	.wave_function_add_2 = waveform_gate,
	.ctx_add_2 = &(waveform_gate_t){ .start = 4.0f/6.0f, .end = 5.0f/6.0f },
};

static const waveform_scale_add_t emergency_beats_ctx = {
	.wave_function_add_1 = waveform_gate,
	.ctx_add_1 = &(waveform_gate_t){ .start = 0.0f, .end = 1.0f/6.0f },
	.wave_function_add_2 = waveform_scale_add,        /* imbrication */
	.ctx_add_2 = &emergency_tail_ctx,
};

static const waveform_space_mult_const_t emergency_ctx = {
	.wave_function_mult = waveform_scale_add,
	.ctx_mult = &emergency_beats_ctx,
	.factor = { .x = 1.0f, .y = 1.0f, .z = 1.0f },   /* blanc */
};




/* ===================================================
   DÉCLENCHEURS
   =================================================== */

static void set_base(const void *space_ctx, uint32_t period_ms) {
	/* Swap du fond : Remove (no-op si le fond a été évincé entre-temps,
	 * grâce à la génération du handle) puis Add de la nouvelle waveform.
	 * L'instance wf_base est réutilisable car l'ancien event est retiré. */
	LedSched_Remove(evt_base);
	Waveform_Init_Space(&wf_base, waveform_space_mult_const,
	                    space_ctx, period_ms, true);
	evt_base = LedSched_Add(&wf_base, /* prio */ 1, /* transparent */ false,
	                        HAL_MAX_DELAY, LED_SCHED_NO_FORCE);
}

static void set_base_space_add(uint32_t period_ms) {
	/* Variante pour la phase C dont la fonction racine est space_add. */
	LedSched_Remove(evt_base);
	Waveform_Init_Space(&wf_base, waveform_space_add, &phase_c_ctx,
	                    period_ms, true);
	evt_base = LedSched_Add(&wf_base, 1, false, HAL_MAX_DELAY, LED_SCHED_NO_FORCE);
}

static void fire_confirm(void) {
	/* One-shot fire-and-forget : non périodique, l'event sort du pool tout
	 * seul quand la waveform se termine ; inutile de garder le handle. */
	Waveform_Init_Space(&wf_confirm, waveform_space_mult_const,
	                    &confirm_ctx, 800, false);
	(void)LedSched_Add(&wf_confirm, 3, false, 3000, LED_SCHED_NO_FORCE);
}

static void fire_blip(uint32_t pause_timeout_ms, led_evt_force_t force) {
	/* Rotation sur le pool d'instances ; le scénario garantit qu'une
	 * instance n'est réutilisée qu'une fois son event précédent terminé. */
	waveform_space_t *wf = &wf_blip[blip_next];
	blip_next = (uint8_t)((blip_next + 1u) % 6u);
	Waveform_Init_Space(wf, waveform_space_mult_const, &blip_ctx, 600, false);
	(void)LedSched_Add(wf, 2, false, pause_timeout_ms, force);
}

static void battery_alert(bool on) {
	if (on && evt_battery == LED_SCHED_HANDLE_INVALID) {
		Waveform_Init_Space(&wf_battery, waveform_space_add_const,
		                    &battery_ctx, 2500, true);
		evt_battery = LedSched_Add(&wf_battery, 5, /* TRANSPARENT */ true,
		                           HAL_MAX_DELAY, LED_SCHED_NO_FORCE);
	} else if (!on && evt_battery != LED_SCHED_HANDLE_INVALID) {
		LedSched_Remove(evt_battery);
		evt_battery = LED_SCHED_HANDLE_INVALID;
	}
}

static void critical_error(bool on) {
	if (on && evt_error == LED_SCHED_HANDLE_INVALID) {
		Waveform_Init_Space(&wf_error, waveform_space_mult_const,
		                    &error_ctx, 900, true);
		/* timeout 0 : même si un jour quelque chose de plus prioritaire la
		 * masque, l'erreur continue de tourner en temps réel. */
		evt_error = LedSched_Add(&wf_error, 9, false, 0, LED_SCHED_NO_FORCE);
	} else if (!on && evt_error != LED_SCHED_HANDLE_INVALID) {
		LedSched_Remove(evt_error);
		evt_error = LED_SCHED_HANDLE_INVALID;
	}
}

static void fire_emergency(void) {
	/* HARD_FORCE : le pool est plein à ce moment du scénario ; l'event de
	 * plus basse priorité (le fond, prio 1) est évincé pour faire place.
	 * evt_base devient alors un handle périmé — sans danger : le prochain
	 * LedSched_Remove(evt_base) sera un no-op silencieux. */
	Waveform_Init_Space(&wf_emergency, waveform_space_mult_const,
	                    &emergency_ctx, 2400, false);
	(void)LedSched_Add(&wf_emergency, 10, false, 0, LED_SCHED_HARD_FORCE);
}




/* ===================================================
   SCÉNARIO
   =================================================== */

typedef enum {
	ACT_BASE_A,
	ACT_BASE_B,
	ACT_BASE_C,
	ACT_CONFIRM,
	ACT_BLIP_VISIBLE,     /* timeout 1500 : attend son tour dans la FIFO   */
	ACT_BLIP_MASKED,      /* timeout 800  : meurt en silence sous l'erreur */
	ACT_BLIP_FILL,        /* timeout inf. : reste en attente, remplit le pool */
	ACT_BATTERY_ON,
	ACT_ERROR_ON,
	ACT_ERROR_OFF,
	ACT_SOFT_FORCE_DEMO,  /* échec attendu : rien de strictement plus faible */
	ACT_EMERGENCY,
} scn_action_t;

static const struct {
	uint32_t     t_ms;
	scn_action_t action;
} scenario[] = {
	{     0, ACT_BASE_A },
	{  5000, ACT_BASE_B },
	{  8000, ACT_CONFIRM },
	{ 10000, ACT_BLIP_VISIBLE },
	{ 10050, ACT_BLIP_VISIBLE },
	{ 10100, ACT_BLIP_VISIBLE },
	{ 13000, ACT_BATTERY_ON },
	{ 16000, ACT_BASE_C },
	{ 19000, ACT_ERROR_ON },
	{ 20000, ACT_BLIP_MASKED },
	{ 20050, ACT_BLIP_MASKED },
	{ 24000, ACT_ERROR_OFF },
	{ 26000, ACT_BLIP_FILL },
	{ 26010, ACT_BLIP_FILL },
	{ 26020, ACT_BLIP_FILL },
	{ 26030, ACT_BLIP_FILL },
	{ 26040, ACT_BLIP_FILL },
	{ 26050, ACT_BLIP_FILL },
	{ 26200, ACT_SOFT_FORCE_DEMO },
	{ 26400, ACT_EMERGENCY },
	{ 29500, ACT_BASE_C },
};
#define SCENARIO_LEN (sizeof(scenario) / sizeof(scenario[0]))

static uint8_t scn_idx;

static void do_action(scn_action_t action) {
	switch (action) {
		case ACT_BASE_A:       set_base(&phase_a_ctx, 2000);            		break;
		case ACT_BASE_B:       set_base(&phase_b_ctx, 1400);            		break;
		case ACT_BASE_C:       set_base_space_add(1500);                		break;
		case ACT_CONFIRM:      fire_confirm();                          		break;
		case ACT_BLIP_VISIBLE: fire_blip(1500, LED_SCHED_NO_FORCE);     		break;
		case ACT_BLIP_MASKED:  fire_blip(800, LED_SCHED_NO_FORCE);      		break;
		case ACT_BLIP_FILL:    fire_blip(HAL_MAX_DELAY, LED_SCHED_NO_FORCE);	break;
		case ACT_BATTERY_ON:   battery_alert(true);                     		break;
		case ACT_ERROR_ON:     critical_error(true);                    		break;
		case ACT_ERROR_OFF:    critical_error(false);                   		break;
		case ACT_SOFT_FORCE_DEMO: {
			/* Pool plein de events prio >= 1, on tente d'ajouter à prio 1 :
			 * la victime potentielle (le fond, prio 1) n'est pas STRICTEMENT
			 * plus faible -> LED_SCHED_HANDLE_INVALID, pool intact. */
			waveform_space_t *wf = &wf_blip[blip_next];
			Waveform_Init_Space(wf, waveform_space_mult_const, &blip_ctx, 600, false);
			led_evt_handle_t h = LedSched_Add(wf, 1, false, 1000, LED_SCHED_SOFT_FORCE);
			(void)h; /* == LED_SCHED_HANDLE_INVALID, échec propre attendu */
			break;
		}
		case ACT_EMERGENCY:    fire_emergency();                        		break;
	}
}




/* ===================================================
   SETUP / LOOP
   =================================================== */

void setup(void) {
	LED_Init(&led_rgb.red,   &htim3, TIM_CHANNEL_1);
	LED_Init(&led_rgb.green, &htim3, TIM_CHANNEL_2);
	LED_Init(&led_rgb.blue,  &htim3, TIM_CHANNEL_3);
	LED_RGB_SetColor(&led_rgb, FLOAT3_ZERO);

	LedSched_Init();

	t0_demo = HAL_GetTick();
	scn_idx = 0;
}

void loop(void) {
	uint32_t now_ms = HAL_GetTick();

	/* Déroule les étapes du scénario dont l'échéance est passée */
	while (scn_idx < SCENARIO_LEN && (now_ms - t0_demo) >= scenario[scn_idx].t_ms) {
		do_action(scenario[scn_idx].action);
		scn_idx++;
	}

	LED_RGB_SetColor(&led_rgb, LedSched_Process(now_ms));
}