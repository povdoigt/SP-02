#ifndef CIRCULAR_BUFFER_H
#define CIRCULAR_BUFFER_H

#include <stdint.h>
#include <stddef.h>

#ifdef RTOS
#include "FreeRTOS.h"
#include "cmsis_os2.h"
#endif

#ifdef __cplusplus
extern "C" {
#endif

/* --------------------------------------------------------------------------
 *   Types et configuration
 * -------------------------------------------------------------------------- */

/**
 * @brief Politique à appliquer en cas de dépassement de capacité.
 */
typedef enum {
    CB_OVERWRITE_OLDEST = 0,  /**< Écrase la donnée la plus ancienne si plein. */
    CB_REJECT_NEW             /**< Rejette la nouvelle donnée si plein.       */
} cb_overflow_policy_t;

/**
 * @brief Codes de statut renvoyés par les fonctions du circular buffer.
 */
typedef enum {
    CB_OK = 0,                /**< Opération réussie. */
    CB_EMPTY,                 /**< Aucune donnée à lire. */
    CB_FULL,                  /**< Tampon plein, nouvelle écriture refusée. */
    CB_OVERWROTE_OLDEST,      /**< Donnée la plus ancienne écrasée. */
    CB_BAD_ARG                /**< Paramètre invalide (pointeur NULL, etc.). */
} cb_status_t;

/**
 * @brief Structure d’un tampon circulaire générique.
 *
 * La mémoire de stockage n’est pas allouée par le buffer.
 * L’appelant doit fournir un pointeur vers un bloc de taille :
 * `elem_size * capacity` octets.
 */
typedef struct circular_buffer_t {
    uint8_t *storage;               /**< Mémoire externe du buffer. */
    size_t   elem_size;             /**< Taille d’un élément (en octets). */
    size_t   capacity;              /**< Nombre maximal d’éléments. */
    size_t   head;                  /**< Prochaine position d’écriture. */
    size_t   tail;                  /**< Position du plus ancien élément. */
    size_t   count;                 /**< Nombre d’éléments actuellement stockés. */
    cb_overflow_policy_t policy;    /**< Politique en cas de dépassement. */
#ifdef RTOS
    // Verouillage par mutex
    osMutexId_t mutex_id;           /**< ID du mutex pour accès thread-safe. */
    StaticSemaphore_t mutex_cm;     /**< Mémoire statique pour le mutex. */
#endif
} circular_buffer_t;

/**
 * @brief Calcule la taille mémoire nécessaire pour stocker `n` éléments du type donné.
 */
#define CIRCULAR_BUFFER_BYTES(type, n) ((size_t)(sizeof(type) * (n)))

/* --------------------------------------------------------------------------
 *   Initialisation / reset
 * -------------------------------------------------------------------------- */

/**
 * @brief Initialise un tampon circulaire sur une mémoire externe.
 * @param cb        Pointeur vers la structure à initialiser.
 * @param storage   Mémoire fournie (doit être allouée par l’appelant).
 * @param elem_size Taille d’un élément (en octets).
 * @param capacity  Nombre maximal d’éléments.
 * @param policy    Politique en cas de dépassement.
 */
void cb_init(circular_buffer_t *cb,
             void *storage, size_t elem_size, size_t capacity,
             cb_overflow_policy_t policy);

/**
 * @brief Vide le buffer sans modifier la mémoire de stockage.
 */
void cb_reset(circular_buffer_t *cb);

void cb_free(circular_buffer_t *cb);

/* --------------------------------------------------------------------------
 *   Écriture / lecture destructive
 * -------------------------------------------------------------------------- */

/**
 * @brief Ajoute une donnée au buffer (copie complète).
 *
 * Si le buffer est plein :
 * - en mode CB_OVERWRITE_OLDEST, l’élément le plus ancien est écrasé,
 * - en mode CB_REJECT_NEW, l’écriture est refusée.
 */
cb_status_t cb_push(circular_buffer_t *cb, const void *elem);

/**
 * @brief Retire et copie la donnée la plus ancienne (FIFO).
 */
cb_status_t cb_pop(circular_buffer_t *cb, void *out);

/* --------------------------------------------------------------------------
 *   Accès pointeur (sans copie)
 * -------------------------------------------------------------------------- */

/**
 * @brief Retourne un pointeur constant vers un élément à un index absolu.
 *
 * @note Comportement "wrap permissif" :
 *       si `idx >= capacity`, l’indice est ramené automatiquement par modulo.
 *       Aucun contrôle n’est effectué sur la validité temporelle de la donnée.
 */
const void *cb_peek_ptr(circular_buffer_t *cb, size_t idx);

/**
 * @brief Retourne un pointeur constant vers un élément relatif à une origine.
 *
 * @param cb      Pointeur vers le buffer.
 * @param origin  Index de base (souvent cb->tail).
 * @param offset  Décalage relatif (peut être positif ou négatif).
 *
 * @note Le comportement est également "wrap permissif".
 */
const void *cb_peek_relative_ptr(circular_buffer_t *cb,
                                 size_t origin, int offset);

/* --------------------------------------------------------------------------
 *   Accès lecture (avec copie)
 * -------------------------------------------------------------------------- */

/**
 * @brief Copie la donnée à un index absolu dans le buffer.
 *
 * Équivalent à : memcpy(out, cb_peek_ptr(cb, idx), elem_size)
 */
cb_status_t cb_peek(circular_buffer_t *cb, size_t idx, void *out);

/**
 * @brief Copie la donnée à un offset relatif à une origine donnée.
 *
 * Équivalent à : memcpy(out, cb_peek_relative_ptr(cb, origin, offset), elem_size)
 */
cb_status_t cb_peek_relative(circular_buffer_t *cb,
                             size_t origin, int offset, void *out);

#ifdef __cplusplus
}
#endif

#endif /* CIRCULAR_BUFFER_H */
