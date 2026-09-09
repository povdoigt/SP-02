#ifndef DT_SEQ_TEST_H
#define DT_SEQ_TEST_H

#include "data_topic.h"
#include "test.h"

#define DT_seq_test_N_TESTS 12

extern TEST_case_table_t DT_seq_test_cases[DT_seq_test_N_TESTS];

/* ========================================================================
 * T0 – Protection NULL / arguments invalides
 *   Appelle les fonctions API avec des pointeurs NULL ou un abonné
 *   non attaché. Verifie que DT_BAD_ARG est retourné sans crash.
 * ======================================================================== */
void DT_seq_test_t0_null_args(TEST_case_t *tc);

/* ========================================================================
 * T1 – Publish / read FIFO basique (1 abonné)
 *   Publie {10, 20, 30}, lit avec data_sub_read dans l'ordre FIFO.
 *   Verifie les valeurs, les codes de retour DT_OK, et que num_to_read
 *   décroît correctement.
 * ======================================================================== */
void DT_seq_test_t1_publish_read_fifo(TEST_case_t *tc);

/* ========================================================================
 * T2 – DT_EMPTY quand aucune donnée disponible
 *   Attache un abonné FROM_NOW avant toute publication.
 *   data_sub_read et data_sub_peek retournent DT_EMPTY.
 *   Verifie que num_to_read == 0.
 * ======================================================================== */
void DT_seq_test_t2_empty_read(TEST_case_t *tc);

/* ========================================================================
 * T3 – DATA_ATTACH_FROM_OLDEST récupère l'historique existant
 *   Publie {10, 20, 30} avant d'attacher l'abonné en mode FROM_OLDEST.
 *   Verifie que num_to_read == 3 et que les lectures donnent
 *   {10, 20, 30} dans l'ordre, suivi de DT_EMPTY.
 * ======================================================================== */
void DT_seq_test_t3_attach_from_oldest(TEST_case_t *tc);

/* ========================================================================
 * T4 – DATA_ATTACH_FROM_NOW ignore l'historique existant
 *   Publie {10, 20, 30} avant d'attacher l'abonné en mode FROM_NOW.
 *   Verifie que num_to_read == 0.
 *   Publie ensuite 40 : num_to_read == 1, lecture retourne 40.
 * ======================================================================== */
void DT_seq_test_t4_attach_from_now(TEST_case_t *tc);

/* ========================================================================
 * T5 – data_sub_num_to_read suit le compteur de publications
 *   Attache FROM_NOW, publie 1, 2 puis 3 elements successivement.
 *   Verifica que num_to_read vaut 1, 2, 3 a chaque étape, et décroît
 *   de 1 apres chaque data_sub_read.
 * ======================================================================== */
void DT_seq_test_t5_num_to_read(TEST_case_t *tc);

/* ========================================================================
 * T6 – Indépendance de deux abonnés (tails séparées)
 *   Attache sub1 et sub2 FROM_NOW.
 *   Publie {10, 20, 30}.
 *   sub1 lit {10, 20}. Vérifie que num_to_read(sub2) == 3 (inchangé).
 *   sub2 lit {10, 20, 30} complètement.
 *   num_to_read(sub1) == 1, num_to_read(sub2) == 0 en fin.
 * ======================================================================== */
void DT_seq_test_t6_two_subscribers(TEST_case_t *tc);

/* ========================================================================
 * T7 – DT_DATA_LOSS quand un abonné est dépassé
 *   capacity=3 / OVERWRITE_OLDEST, attache sub FROM_NOW (aucune pub).
 *   Publie 4 éléments : delta = pub_seq - last_seq = 4 > cap = 3.
 *   data_sub_peek_relative_ptr retourne DT_DATA_LOSS (et synce le sub).
 *   Après sync : num_to_read == 0.
 *   Nouvelle publication → lecture DT_OK correcte.
 * ======================================================================== */
void DT_seq_test_t7_data_loss(TEST_case_t *tc);

/* ========================================================================
 * T8 – data_sub_sync réaligne l'abonné sur la tête courante
 *   Publie {10, 20, 30}, puis appelle data_sub_sync.
 *   Verifie que num_to_read == 0 et que data_sub_read retourne DT_EMPTY.
 *   Publie ensuite 40 : lecture == 40.
 * ======================================================================== */
void DT_seq_test_t8_sync(TEST_case_t *tc);

/* ========================================================================
 * T9 – data_sub_detach / sub_count / re-attachement
 *   Attache sub1 et sub2 → sub_count == 2.
 *   Detache sub1 → sub_count == 1, sub1.attached == 0.
 *   Publie 99 → sub2 lit 99 (DT_OK), sub1 ne peut pas lire.
 *   Re-attache sub1 FROM_NOW → sub_count == 2, num_to_read(sub1) == 0.
 * ======================================================================== */
void DT_seq_test_t9_detach_reattach(TEST_case_t *tc);

/* ========================================================================
 * T10 – data_sub_peek : lecture non destructive (tail inchangée)
 *   Publie {10, 20, 30}. Effectue plusieurs data_sub_peek (index 0,1,2).
 *   Verifie les valeurs et que :
 *     - num_to_read reste 3 après les peeks.
 *     - sub->tail est inchangé.
 *     - data_sub_read suivant retourne bien 10 (premier élément).
 * ======================================================================== */
void DT_seq_test_t10_peek_non_destructive(TEST_case_t *tc);

/* ========================================================================
 * T11 – Politique CB_REJECT_NEW → DT_FULL à la publication
 *   capacity=2 / REJECT_NEW.
 *   Publie 10, 20 → DT_OK. Publie 30 → DT_FULL (buffer plein).
 *   Verifie que pub_seq == 2 (pas d'incrément) et que les données
 *   lues sont {10, 20} (30 non stocké).
 * ======================================================================== */
void DT_seq_test_t11_reject_new_policy(TEST_case_t *tc);

#endif /* DT_SEQ_TEST_H */
