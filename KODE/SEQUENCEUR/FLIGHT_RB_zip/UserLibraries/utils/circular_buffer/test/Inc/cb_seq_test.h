#ifndef CB_SEQ_TEST_H
#define CB_SEQ_TEST_H

#include "circular_buffer.h"
#include "test.h"

#define CB_seq_test_N_TESTS 11

extern TEST_case_table_t CB_seq_test_cases[CB_seq_test_N_TESTS];

/* ========================================================================
 * T0 – Protection NULL / arguments invalides
 *   Appelle cb_init, cb_push et cb_pop avec des pointeurs NULL.
 *   Verifie qu'aucun crash ne se produit et que les codes d'erreur
 *   corrects sont retournes (CB_BAD_ARG).
 * ======================================================================== */
void CB_seq_test_t0_null_args(TEST_case_t *tc);

/* ========================================================================
 * T1 – Ordre FIFO basique (push / pop 3 elements)
 *   Pousse {10, 20, 30} et depile dans l'ordre. Verifie :
 *   - cb_push retourne CB_OK pour chaque insertion.
 *   - cb_pop retourne les elements dans l'ordre FIFO exact.
 *   - count == 0 apres epuisement.
 * ======================================================================== */
void CB_seq_test_t1_fifo_order(TEST_case_t *tc);

/* ========================================================================
 * T2 – Lecture sur buffer vide → CB_EMPTY
 *   cb_pop sur un buffer fraichement initialise doit retourner CB_EMPTY
 *   sans modifier la memoire de sortie.
 * ======================================================================== */
void CB_seq_test_t2_empty_read(TEST_case_t *tc);

/* ========================================================================
 * T3 – Buffer plein avec politique CB_REJECT_NEW → CB_FULL
 *   capacity=3, pousse 4 elements. Le 4eme doit retourner CB_FULL.
 *   Les 3 premiers elements doivent rester intacts en memoire.
 * ======================================================================== */
void CB_seq_test_t3_full_reject(TEST_case_t *tc);

/* ========================================================================
 * T4 – Buffer plein avec politique CB_OVERWRITE_OLDEST
 *   capacity=3, pousse {1, 2, 3, 4}. Le 4eme push doit retourner
 *   CB_OVERWROTE_OLDEST. Apres overwrite, le premier pop renvoie 2
 *   (l'element 1 a ete ecrase), puis 3, puis 4.
 * ======================================================================== */
void CB_seq_test_t4_full_overwrite(TEST_case_t *tc);

/* ========================================================================
 * T5 – cb_reset vide logiquement le buffer
 *   Pousse 2 elements, cb_reset, verifie count == 0 et que cb_pop
 *   retourne CB_EMPTY. Puis repousse 1 element pour confirmer que le
 *   buffer est reutilisable apres reset.
 * ======================================================================== */
void CB_seq_test_t5_reset(TEST_case_t *tc);

/* ========================================================================
 * T6 – cb_peek : couverture complete du wrap permissif sur l'index absolu
 *   Logique : cb_peek(cb, idx) lit le slot physique  idx % capacity.
 *   capacity=4, storage[0..3]={10,20,30,40}.
 *   Scenarios verifies :
 *     A) idx in [0,3]         -> acces directs.
 *     B) idx == cap           -> wrap  -> slot 0 = 10.
 *     C) idx == cap+2         -> wrap  -> slot 2 = 30.
 *     D) idx == 2*cap         -> double wrap -> slot 0 = 10.
 *     E) idx == 17            -> 17%4=1 -> slot 1 = 20.
 *     F) non destructif       -> count inchange apres tous les peeks.
 * ======================================================================== */
void CB_seq_test_t6_peek_absolute(TEST_case_t *tc);

/* ========================================================================
 * T7 – cb_peek_relative : couverture complete du wrap permissif
 *   Logique : index_physique = (origin + offset) % capacity  (signe gere).
 *   capacity=5, storage[0..4]={10,20,30,40,50}.
 *   Scenarios verifies :
 *     A) offsets +0..+4 depuis origin=0    -> acces directs.
 *     B) offset == +cap    (0+5)%5=0       -> slot 0 = 10.
 *     C) offset == +cap+2  (0+7)%5=2       -> slot 2 = 30.
 *     D) offset == -cap    (0-5+5)%5=0     -> slot 0 = 10.
 *     E) offset == -1      (0-1+5)%5=4     -> slot 4 = 50.
 *     F) offset == -2      (0-2+5)%5=3     -> slot 3 = 40.
 *     G) origin=2, +3      (2+3)%5=0       -> slot 0 = 10.
 *     H) origin=2, -3      (2-3+5)%5=4     -> slot 4 = 50.
 *     I) origin=1, +11     (1+11)%5=2      -> slot 2 = 30.
 *     J) origin=1, -11     (1-11+15)%5=0   -> slot 0 = 10.
 *     K) non destructif    -> count inchange apres tous les peeks.
 * ======================================================================== */
void CB_seq_test_t7_peek_relative(TEST_case_t *tc);

/* ========================================================================
 * T8 – Wrap-around (head et tail traversent la frontiere de capacite)
 *   capacity=4, pousse {1,2,3,4}, pop {1,2}, pousse {5,6}.
 *   Le head et le tail doivent avoir effectue un wrap.
 *   Verifie que le pop donne {3, 4, 5, 6} dans l'ordre.
 * ======================================================================== */
void CB_seq_test_t8_wraparound(TEST_case_t *tc);

/* ========================================================================
 * T9 – Elem_size variable : type float
 *   Cree un buffer de capacity=4 avec elem_size=sizeof(float).
 *   Pousse {1.5f, 2.5f, 3.5f}, pop et verifie les valeurs en virgule
 *   flottante (tolerance 1e-6).
 * ======================================================================== */
void CB_seq_test_t9_float_elemsize(TEST_case_t *tc);

/* ========================================================================
 * T10 – Cycle remplissage / vidange repete
 *   capacity=8. Cycle 1 : pousse [0..7], pop et verifie [0..7].
 *   Cycle 2 : pousse [8..15], pop et verifie [8..15].
 *   Verifie que le buffer est parfaitement reutilisable apres un cycle
 *   complet (head, tail, count tous remis a zero logiquement).
 * ======================================================================== */
void CB_seq_test_t10_fill_drain_cycle(TEST_case_t *tc);

#endif /* CB_SEQ_TEST_H */
