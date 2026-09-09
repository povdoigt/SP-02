#include "cb_seq_test.h"

#include <stdio.h>
#include <math.h>
#include <string.h>

/* ========================================================================
 * Table des cas de test
 * ======================================================================== */

TEST_case_table_t CB_seq_test_cases[CB_seq_test_N_TESTS] = {
    { .case_info = { .name = "T0 NULL args"       }, .func = CB_seq_test_t0_null_args         },
    { .case_info = { .name = "T1 FIFO order"      }, .func = CB_seq_test_t1_fifo_order        },
    { .case_info = { .name = "T2 Empty read"      }, .func = CB_seq_test_t2_empty_read        },
    { .case_info = { .name = "T3 Full REJECT_NEW" }, .func = CB_seq_test_t3_full_reject       },
    { .case_info = { .name = "T4 Full OVERWRITE"  }, .func = CB_seq_test_t4_full_overwrite    },
    { .case_info = { .name = "T5 Reset"           }, .func = CB_seq_test_t5_reset             },
    { .case_info = { .name = "T6 Peek absolute"   }, .func = CB_seq_test_t6_peek_absolute     },
    { .case_info = { .name = "T7 Peek relative"   }, .func = CB_seq_test_t7_peek_relative     },
    { .case_info = { .name = "T8 Wrap-around"     }, .func = CB_seq_test_t8_wraparound        },
    { .case_info = { .name = "T9 Float elem_size" }, .func = CB_seq_test_t9_float_elemsize    },
    { .case_info = { .name = "T10 Fill/Drain x2"  }, .func = CB_seq_test_t10_fill_drain_cycle },
};

#define CAP6 4u
#define CAP7 5u

/* ========================================================================
 * T0 – Protection NULL / arguments invalides
 * ======================================================================== */
void CB_seq_test_t0_null_args(TEST_case_t *tc) {
    tc->result = R_FAIL;

    /* cb_init avec arguments invalides – doit retourner CB_BAD_ARG sans crasher */
    cb_status_t s;
    s = cb_init(NULL, NULL, sizeof(uint32_t), 4, CB_REJECT_NEW);
    TEST_ASSERT(s == CB_BAD_ARG, "cb_init(NULL,NULL,...) retourne %d != CB_BAD_ARG", s);

    uint8_t storage[4 * sizeof(uint32_t)];
    s = cb_init(NULL, storage, sizeof(uint32_t), 4, CB_REJECT_NEW);
    TEST_ASSERT(s == CB_BAD_ARG, "cb_init(NULL,storage,...) retourne %d != CB_BAD_ARG", s);

    circular_buffer_t cb;
    s = cb_init(&cb, NULL, sizeof(uint32_t), 4, CB_REJECT_NEW);
    TEST_ASSERT(s == CB_BAD_ARG, "cb_init(cb,NULL,...) retourne %d != CB_BAD_ARG", s);
    s = cb_init(&cb, storage, 0,             4, CB_REJECT_NEW);
    TEST_ASSERT(s == CB_BAD_ARG, "cb_init(cb,storage,0,...) retourne %d != CB_BAD_ARG", s);
    s = cb_init(&cb, storage, sizeof(uint32_t), 0, CB_REJECT_NEW);
    TEST_ASSERT(s == CB_BAD_ARG, "cb_init(cb,...,cap=0,...) retourne %d != CB_BAD_ARG", s);

    /* cb_push / cb_pop avec NULL cb */
    uint32_t val = 42;

    s = cb_push(NULL, &val);
    TEST_ASSERT(s == CB_BAD_ARG, "cb_push(NULL,val) retourne %d != CB_BAD_ARG", s);

    s = cb_pop(NULL, &val);
    TEST_ASSERT(s == CB_BAD_ARG, "cb_pop(NULL,out) retourne %d != CB_BAD_ARG", s);

    /* cb_push avec elem NULL sur buffer valide */
    s = cb_init(&cb, storage, sizeof(uint32_t), 4, CB_REJECT_NEW);
    TEST_ASSERT(s == CB_OK, "cb_init valide retourne %d != CB_OK", s);
    s = cb_push(&cb, NULL);
    TEST_ASSERT(s == CB_BAD_ARG, "cb_push(cb,NULL) retourne %d != CB_BAD_ARG", s);
    s = cb_free(&cb);
    TEST_ASSERT(s == CB_OK, "cb_free retourne %d != CB_OK", s);

    snprintf(tc->detail, sizeof(tc->detail), "Tous les NULL/bad-arg correctement rejetes");
    tc->result = R_PASS;
}

/* ========================================================================
 * T1 – Ordre FIFO basique
 * ======================================================================== */
void CB_seq_test_t1_fifo_order(TEST_case_t *tc) {
    tc->result = R_FAIL;

    uint32_t storage[3];
    circular_buffer_t cb;
    cb_status_t s = cb_init(&cb, storage, sizeof(uint32_t), 3, CB_REJECT_NEW);
    TEST_ASSERT(s == CB_OK, "cb_init retourne %d != CB_OK", s);

    uint32_t in[3] = {10u, 20u, 30u};

    for (int i = 0; i < 3; i++) {
        s = cb_push(&cb, &in[i]);
        TEST_ASSERT(s == CB_OK, "push[%d] retourne %d != CB_OK", i, s);
    }
    TEST_ASSERT(cb.count == 3u, "count=%u != 3 apres 3 push", (unsigned)cb.count);

    uint32_t out;
    for (int i = 0; i < 3; i++) {
        s = cb_pop(&cb, &out);
        TEST_ASSERT(s == CB_OK,       "pop[%d] retourne %d != CB_OK", i, s);
        TEST_ASSERT(out == in[i],     "pop[%d]=%u attendu %u", i, (unsigned)out, (unsigned)in[i]);
    }
    TEST_ASSERT(cb.count == 0u, "count=%u != 0 apres 3 pop", (unsigned)cb.count);

    s = cb_free(&cb);
    TEST_ASSERT(s == CB_OK, "cb_free retourne %d != CB_OK", s);
    snprintf(tc->detail, sizeof(tc->detail), "FIFO {10,20,30} OK, count=0 en fin");
    tc->result = R_PASS;
}

/* ========================================================================
 * T2 – Lecture sur buffer vide
 * ======================================================================== */
void CB_seq_test_t2_empty_read(TEST_case_t *tc) {
    tc->result = R_FAIL;

    uint32_t storage[4];
    circular_buffer_t cb;
    cb_status_t s = cb_init(&cb, storage, sizeof(uint32_t), 4, CB_REJECT_NEW);
    TEST_ASSERT(s == CB_OK, "cb_init retourne %d != CB_OK", s);

    uint32_t out = 0xDEADBEEFu;
    s = cb_pop(&cb, &out);

    TEST_ASSERT(s == CB_EMPTY,          "cb_pop sur vide retourne %d != CB_EMPTY", s);
    TEST_ASSERT(out == 0xDEADBEEFu,     "cb_pop ne doit pas ecrire sur out (out=0x%08X)", (unsigned)out);
    TEST_ASSERT(cb.count == 0u,         "count=%u != 0", (unsigned)cb.count);

    s = cb_free(&cb);
    TEST_ASSERT(s == CB_OK, "cb_free retourne %d != CB_OK", s);
    snprintf(tc->detail, sizeof(tc->detail), "CB_EMPTY retourne, out intact, count=0");
    tc->result = R_PASS;
}

/* ========================================================================
 * T3 – Buffer plein avec CB_REJECT_NEW
 * ======================================================================== */
void CB_seq_test_t3_full_reject(TEST_case_t *tc) {
    tc->result = R_FAIL;

    uint32_t storage[3];
    circular_buffer_t cb;
    uint32_t val;
    cb_status_t s = cb_init(&cb, storage, sizeof(uint32_t), 3, CB_REJECT_NEW);
    TEST_ASSERT(s == CB_OK, "cb_init retourne %d != CB_OK", s);

    for (uint32_t i = 1u; i <= 3u; i++) {
        s = cb_push(&cb, &i);
        TEST_ASSERT(s == CB_OK, "push #%u retourne %d != CB_OK", (unsigned)i, s);
    }
    TEST_ASSERT(cb.count == 3u, "count=%u != 3 avant push de trop", (unsigned)cb.count);

    val = 99u;
    s = cb_push(&cb, &val);
    TEST_ASSERT(s == CB_FULL, "4eme push retourne %d != CB_FULL", s);
    TEST_ASSERT(cb.count == 3u, "count=%u != 3 apres refus", (unsigned)cb.count);

    /* Verifie que les 3 premiers elements sont intacts */
    for (uint32_t i = 1u; i <= 3u; i++) {
        uint32_t out = 0u;
        s = cb_pop(&cb, &out);
        TEST_ASSERT(s  == CB_OK, "pop #%u retourne %d != CB_OK", (unsigned)i, s);
        TEST_ASSERT(out == i,    "pop #%u=%u attendu %u", (unsigned)i, (unsigned)out, (unsigned)i);
    }

    s = cb_free(&cb);
    TEST_ASSERT(s == CB_OK, "cb_free retourne %d != CB_OK", s);
    snprintf(tc->detail, sizeof(tc->detail), "CB_FULL ok, {1,2,3} intacts, 99 rejete");
    tc->result = R_PASS;
}

/* ========================================================================
 * T4 – Buffer plein avec CB_OVERWRITE_OLDEST
 * ======================================================================== */
void CB_seq_test_t4_full_overwrite(TEST_case_t *tc) {
    tc->result = R_FAIL;

    uint32_t storage[3];
    circular_buffer_t cb;
    cb_status_t s = cb_init(&cb, storage, sizeof(uint32_t), 3, CB_OVERWRITE_OLDEST);
    TEST_ASSERT(s == CB_OK, "cb_init retourne %d != CB_OK", s);
    for (uint32_t i = 1u; i <= 3u; i++) {
        s = cb_push(&cb, &i);
        TEST_ASSERT(s == CB_OK, "push #%u retourne %d != CB_OK", (unsigned)i, s);
    }

    uint32_t val = 4u;
    s = cb_push(&cb, &val);
    TEST_ASSERT(s == CB_OVERWROTE_OLDEST, "4eme push retourne %d != CB_OVERWROTE_OLDEST", s);
    TEST_ASSERT(cb.count == 3u, "count=%u != 3 apres overwrite", (unsigned)cb.count);

    /* Apres overwrite de 1 par 4 : FIFO doit etre {2, 3, 4} */
    uint32_t expected[3] = {2u, 3u, 4u};
    for (int i = 0; i < 3; i++) {
        uint32_t out = 0u;
        s = cb_pop(&cb, &out);
        TEST_ASSERT(s   == CB_OK,          "pop[%d] retourne %d != CB_OK", i, s);
        TEST_ASSERT(out == expected[i],     "pop[%d]=%u attendu %u", i, (unsigned)out, (unsigned)expected[i]);
    }

    s = cb_free(&cb);
    TEST_ASSERT(s == CB_OK, "cb_free retourne %d != CB_OK", s);
    snprintf(tc->detail, sizeof(tc->detail), "CB_OVERWROTE_OLDEST ok, FIFO={2,3,4} apres push 4");
    tc->result = R_PASS;
}

/* ========================================================================
 * T5 – cb_reset vide logiquement le buffer
 * ======================================================================== */
void CB_seq_test_t5_reset(TEST_case_t *tc) {
    tc->result = R_FAIL;

    uint32_t storage[4];
    circular_buffer_t cb;
    cb_status_t s = cb_init(&cb, storage, sizeof(uint32_t), 4, CB_REJECT_NEW);
    TEST_ASSERT(s == CB_OK, "cb_init retourne %d != CB_OK", s);

    uint32_t val;
    val = 11u; cb_push(&cb, &val);
    val = 22u; cb_push(&cb, &val);
    TEST_ASSERT(cb.count == 2u, "count=%u != 2 avant reset", (unsigned)cb.count);

    s = cb_reset(&cb);
    TEST_ASSERT(s == CB_OK, "cb_reset retourne %d != CB_OK", s);
    TEST_ASSERT(cb.count == 0u, "count=%u != 0 apres reset", (unsigned)cb.count);

    uint32_t out = 0xDEADBEEFu;
    s = cb_pop(&cb, &out);
    TEST_ASSERT(s == CB_EMPTY, "cb_pop apres reset retourne %d != CB_EMPTY", s);

    /* Le buffer doit etre reutilisable apres reset */
    val = 55u;
    s = cb_push(&cb, &val);
    TEST_ASSERT(s == CB_OK, "push apres reset retourne %d != CB_OK", s);
    s = cb_pop(&cb, &out);
    TEST_ASSERT(s  == CB_OK && out == 55u,
              "pop apres reset : s=%d out=%u (attendu 55)", s, (unsigned)out);

    s = cb_free(&cb);
    TEST_ASSERT(s == CB_OK, "cb_free retourne %d != CB_OK", s);
    snprintf(tc->detail, sizeof(tc->detail), "Reset: count=0, CB_EMPTY, reutilisable OK");
    tc->result = R_PASS;
}

/* ========================================================================
 * T6 – cb_peek : comportement wrap permissif sur l'index absolu
 *
 * Logique testee : cb_peek(cb, idx) est equivalent a
 *   storage[ idx % capacity ]  (wrap permissif, pas de garde).
 *
 * Scenarios :
 *   A) idx in [0, cap-1]  : acces directs aux slots physiques.
 *   B) idx == capacity    : wrap -> slot 0.
 *   C) idx == capacity+1  : wrap -> slot 1.
 *   D) idx == 2*capacity  : double wrap -> slot 0 a nouveau.
 *   E) idx >> capacity (grand index) : modulo exact.
 *   F) non destructif : count inchange apres tous les peek.
 * ======================================================================== */
void CB_seq_test_t6_peek_absolute(TEST_case_t *tc) {
    tc->result = R_FAIL;

    /* capacity = 4, on remplit les 4 slots : storage[0..3] = {10,20,30,40} */
#define CAP6 4u
    uint32_t storage[CAP6];
    circular_buffer_t cb;
    cb_status_t s = cb_init(&cb, storage, sizeof(uint32_t), CAP6, CB_REJECT_NEW);
    TEST_ASSERT(s == CB_OK, "cb_init retourne %d != CB_OK", s);

    for (uint32_t i = 0u; i < CAP6; i++) {
        uint32_t v = (i + 1u) * 10u;   /* 10, 20, 30, 40 */
        cb_push(&cb, &v);
    }
    /* head == 0 (wrap), tail == 0 : slots physiques [0..3] = {10,20,30,40} */

    /* --- A) acces normal in [0, cap-1] --- */
    uint32_t expected_slot[CAP6] = {10u, 20u, 30u, 40u};
    for (size_t i = 0u; i < CAP6; i++) {
        uint32_t out = 0u;
        s = cb_peek(&cb, i, &out);
        TEST_ASSERT(s   == CB_OK,              "[A] peek[%u] retourne %d", (unsigned)i, s);
        TEST_ASSERT(out == expected_slot[i],    "[A] peek[%u]=%u attendu %u",
                  (unsigned)i, (unsigned)out, (unsigned)expected_slot[i]);
    }

    /* --- B) idx == capacity  ->  idx % 4 == 0  -> slot 0 = 10 --- */
    {
        uint32_t out = 0u;
        s = cb_peek(&cb, CAP6, &out);
        TEST_ASSERT(s   == CB_OK,  "[B] peek[cap] retourne %d", s);
        TEST_ASSERT(out == 10u,    "[B] peek[cap]=%u attendu 10", (unsigned)out);
    }

    /* --- C) idx == capacity+1  ->  idx % 4 == 1  -> slot 1 = 20 --- */
    {
        uint32_t out = 0u;
        s = cb_peek(&cb, CAP6 + 1u, &out);
        TEST_ASSERT(s   == CB_OK,  "[C] peek[cap+1] retourne %d", s);
        TEST_ASSERT(out == 20u,    "[C] peek[cap+1]=%u attendu 20", (unsigned)out);
    }

    /* --- D) idx == 2*capacity  ->  idx % 4 == 0  -> slot 0 = 10 --- */
    {
        uint32_t out = 0u;
        s = cb_peek(&cb, 2u * CAP6, &out);
        TEST_ASSERT(s   == CB_OK,  "[D] peek[2*cap] retourne %d", s);
        TEST_ASSERT(out == 10u,    "[D] peek[2*cap]=%u attendu 10", (unsigned)out);
    }

    /* --- E) grand index : idx == 17  ->  17 % 4 == 1  -> slot 1 = 20 --- */
    {
        uint32_t out = 0u;
        s = cb_peek(&cb, 17u, &out);
        TEST_ASSERT(s   == CB_OK,  "[E] peek[17] retourne %d", s);
        TEST_ASSERT(out == 20u,    "[E] peek[17]=%u attendu 20 (17%%4=1)", (unsigned)out);
    }

    /* --- F) non destructif --- */
    TEST_ASSERT(cb.count == CAP6, "[F] count=%u != %u apres tous les peek",
              (unsigned)cb.count, (unsigned)CAP6);

    s = cb_free(&cb);
    TEST_ASSERT(s == CB_OK, "cb_free retourne %d != CB_OK", s);
    snprintf(tc->detail, sizeof(tc->detail),
             "peek wrap-permissif: idx%%cap OK (B,C,D,E), non-destructif OK");
    tc->result = R_PASS;
#undef CAP6
}

/* ========================================================================
 * T7 – cb_peek_relative : couverture complete du wrap permissif
 *
 * Logique : index_physique = (origin + offset) % capacity  (modulo signe)
 *   - wrap_add gere les negatifs via  ((x % n) + n) % n.
 *
 * Etat initial :
 *   capacity=5, push {10,20,30,40,50} => storage[0..4]={10,20,30,40,50}
 *   head=0 (wrap), tail=0.
 *
 * Scenarios :
 *   A) offset positif [0..4]  depuis origin=0  -> slots directs.
 *   B) offset == +cap         depuis origin=0  -> (0+5)%5=0 -> 10.
 *   C) offset == +cap+2       depuis origin=0  -> (0+7)%5=2 -> 30.
 *   D) offset == -(cap)       depuis origin=0  -> (0-5)%5=0 -> 10.
 *   E) offset == -1           depuis origin=0  -> (0-1+5)%5=4 -> 50.
 *   F) offset == -2           depuis origin=0  -> (0-2+5)%5=3 -> 40.
 *   G) origin arbitraire=2, offset=+3 -> (2+3)%5=0 -> 10.
 *   H) origin arbitraire=2, offset=-3 -> (2-3+5)%5=4 -> 50.
 *   I) grand offset positif  -> origin=1, offset=+11 -> (1+11)%5=2 -> 30.
 *   J) grand offset negatif  -> origin=1, offset=-11 -> (1-11+15)%5=0 -> 10
 *      (wrap_add : ((1-11) % 5 + 5) % 5 = ((-10)%5+5)%5 = (0+5)%5 = 0).
 *   K) non destructif : count inchange.
 * ======================================================================== */
void CB_seq_test_t7_peek_relative(TEST_case_t *tc) {
    tc->result = R_FAIL;

#define CAP7 5u
    uint32_t storage[CAP7];
    circular_buffer_t cb;
    cb_status_t s = cb_init(&cb, storage, sizeof(uint32_t), CAP7, CB_REJECT_NEW);
    TEST_ASSERT(s == CB_OK, "cb_init retourne %d != CB_OK", s);

    /* Remplit les 5 slots : storage[i] = (i+1)*10 */
    for (uint32_t i = 0u; i < CAP7; i++) {
        uint32_t v = (i + 1u) * 10u;   /* 10, 20, 30, 40, 50 */
        cb_push(&cb, &v);
    }
    /* head=0 (wrap), tail=0. storage[0..4] = {10,20,30,40,50} */

    /* --- A) offsets positifs normaux depuis origin=0 --- */
    uint32_t slot_val[CAP7] = {10u, 20u, 30u, 40u, 50u};
    for (int off = 0; off < (int)CAP7; off++) {
        uint32_t out = 0u;
        s = cb_peek_relative(&cb, 0u, off, &out);
        TEST_ASSERT(s   == CB_OK,          "[A] peek_rel(0,+%d) retourne %d", off, s);
        TEST_ASSERT(out == slot_val[off],   "[A] peek_rel(0,+%d)=%u attendu %u",
                  off, (unsigned)out, (unsigned)slot_val[off]);
    }

    /* --- B) offset == +CAP7  ->  (0+5)%5=0  -> storage[0]=10 --- */
    {
        uint32_t out = 0u;
        s = cb_peek_relative(&cb, 0u, (int)CAP7, &out);
        TEST_ASSERT(s   == CB_OK, "[B] peek_rel(0,+cap) retourne %d", s);
        TEST_ASSERT(out == 10u,   "[B] peek_rel(0,+cap)=%u attendu 10", (unsigned)out);
    }

    /* --- C) offset == +CAP7+2  ->  (0+7)%5=2  -> storage[2]=30 --- */
    {
        uint32_t out = 0u;
        s = cb_peek_relative(&cb, 0u, (int)CAP7 + 2, &out);
        TEST_ASSERT(s   == CB_OK, "[C] peek_rel(0,+cap+2) retourne %d", s);
        TEST_ASSERT(out == 30u,   "[C] peek_rel(0,+cap+2)=%u attendu 30", (unsigned)out);
    }

    /* --- D) offset == -cap  ->  (0-5+5)%5=0  -> storage[0]=10 --- */
    {
        uint32_t out = 0u;
        s = cb_peek_relative(&cb, 0u, -(int)CAP7, &out);
        TEST_ASSERT(s   == CB_OK, "[D] peek_rel(0,-cap) retourne %d", s);
        TEST_ASSERT(out == 10u,   "[D] peek_rel(0,-cap)=%u attendu 10", (unsigned)out);
    }

    /* --- E) offset == -1  ->  (0-1+5)%5=4  -> storage[4]=50 --- */
    {
        uint32_t out = 0u;
        s = cb_peek_relative(&cb, 0u, -1, &out);
        TEST_ASSERT(s   == CB_OK, "[E] peek_rel(0,-1) retourne %d", s);
        TEST_ASSERT(out == 50u,   "[E] peek_rel(0,-1)=%u attendu 50", (unsigned)out);
    }

    /* --- F) offset == -2  ->  (0-2+5)%5=3  -> storage[3]=40 --- */
    {
        uint32_t out = 0u;
        s = cb_peek_relative(&cb, 0u, -2, &out);
        TEST_ASSERT(s   == CB_OK, "[F] peek_rel(0,-2) retourne %d", s);
        TEST_ASSERT(out == 40u,   "[F] peek_rel(0,-2)=%u attendu 40", (unsigned)out);
    }

    /* --- G) origin=2, offset=+3  ->  (2+3)%5=0  -> storage[0]=10 --- */
    {
        uint32_t out = 0u;
        s = cb_peek_relative(&cb, 2u, +3, &out);
        TEST_ASSERT(s   == CB_OK, "[G] peek_rel(2,+3) retourne %d", s);
        TEST_ASSERT(out == 10u,   "[G] peek_rel(2,+3)=%u attendu 10", (unsigned)out);
    }

    /* --- H) origin=2, offset=-3  ->  (2-3+5)%5=4  -> storage[4]=50 --- */
    {
        uint32_t out = 0u;
        s = cb_peek_relative(&cb, 2u, -3, &out);
        TEST_ASSERT(s   == CB_OK, "[H] peek_rel(2,-3) retourne %d", s);
        TEST_ASSERT(out == 50u,   "[H] peek_rel(2,-3)=%u attendu 50", (unsigned)out);
    }

    /* --- I) grand offset positif : origin=1, offset=+11
     *   (1+11)%5 = 12%5 = 2  -> storage[2]=30 --- */
    {
        uint32_t out = 0u;
        s = cb_peek_relative(&cb, 1u, +11, &out);
        TEST_ASSERT(s   == CB_OK, "[I] peek_rel(1,+11) retourne %d", s);
        TEST_ASSERT(out == 30u,   "[I] peek_rel(1,+11)=%u attendu 30 ((1+11)%%5=2)",
                  (unsigned)out);
    }

    /* --- J) grand offset negatif : origin=1, offset=-11
     *   wrap_add: ((1-11)%5 + 5)%5 = ((-10)%5+5)%5 = (0+5)%5 = 0  -> storage[0]=10
     *   (note: (-10)%5 == 0 en C, pas de reste negatif ici)         --- */
    {
        uint32_t out = 0u;
        s = cb_peek_relative(&cb, 1u, -11, &out);
        TEST_ASSERT(s   == CB_OK, "[J] peek_rel(1,-11) retourne %d", s);
        TEST_ASSERT(out == 10u,   "[J] peek_rel(1,-11)=%u attendu 10 ((1-11)%%5=0)",
                  (unsigned)out);
    }

    /* --- K) non destructif --- */
    TEST_ASSERT(cb.count == CAP7, "[K] count=%u != %u apres tous les peek_relative",
              (unsigned)cb.count, (unsigned)CAP7);

    s = cb_free(&cb);
    TEST_ASSERT(s == CB_OK, "cb_free retourne %d != CB_OK", s);
    snprintf(tc->detail, sizeof(tc->detail),
             "peek_rel wrap-permissif: A-J OK (pos/neg/grand/origin arb.), non-dest.");
    tc->result = R_PASS;
#undef CAP7
}

/* ========================================================================
 * T8 – Wrap-around (head et tail traversent la frontiere capacite)
 * ======================================================================== */
void CB_seq_test_t8_wraparound(TEST_case_t *tc) {
    tc->result = R_FAIL;

    uint32_t storage[4];
    circular_buffer_t cb;
    cb_status_t s = cb_init(&cb, storage, sizeof(uint32_t), 4, CB_REJECT_NEW);
    TEST_ASSERT(s == CB_OK, "cb_init retourne %d != CB_OK", s);

    /* Remplit le buffer : slots [0..3] = {1,2,3,4}, head wraps a 0 */
    for (uint32_t i = 1u; i <= 4u; i++) {
        cb_push(&cb, &i);
    }

    /* Pop 2 elements : tail passe de 0 a 2 */
    uint32_t out;
    cb_pop(&cb, &out);
    TEST_ASSERT(out == 1u, "pop #1 = %u attendu 1", (unsigned)out);
    cb_pop(&cb, &out);
    TEST_ASSERT(out == 2u, "pop #2 = %u attendu 2", (unsigned)out);
    TEST_ASSERT(cb.count == 2u && cb.tail == 2u && cb.head == 0u,
              "Etat apres 2 pop: count=%u tail=%u head=%u",
              (unsigned)cb.count, (unsigned)cb.tail, (unsigned)cb.head);

    /* Push 5 et 6 : head passe par le slot 0 puis 1 (wrap effectif) */
    uint32_t v5 = 5u, v6 = 6u;
    cb_push(&cb, &v5);
    cb_push(&cb, &v6);
    TEST_ASSERT(cb.count == 4u, "count=%u != 4 apres 2 push supplementaires", (unsigned)cb.count);

    /* Le FIFO doit sortir {3, 4, 5, 6} */
    uint32_t expected[4] = {3u, 4u, 5u, 6u};
    for (int i = 0; i < 4; i++) {
        s = cb_pop(&cb, &out);
        TEST_ASSERT(s   == CB_OK,          "pop[%d] retourne %d != CB_OK", i, s);
        TEST_ASSERT(out == expected[i],     "pop[%d]=%u attendu %u", i, (unsigned)out, (unsigned)expected[i]);
    }

    TEST_ASSERT(cb.count == 0u, "count=%u != 0 apres vidange complete", (unsigned)cb.count);

    s = cb_free(&cb);
    TEST_ASSERT(s == CB_OK, "cb_free retourne %d != CB_OK", s);
    snprintf(tc->detail, sizeof(tc->detail), "Wrap-around: FIFO={3,4,5,6} correct, count=0");
    tc->result = R_PASS;
}

/* ========================================================================
 * T9 – elem_size variable : type float
 * ======================================================================== */
void CB_seq_test_t9_float_elemsize(TEST_case_t *tc) {
    tc->result = R_FAIL;

    float storage[4];
    circular_buffer_t cb;
    cb_status_t s = cb_init(&cb, storage, sizeof(float), 4, CB_REJECT_NEW);
    TEST_ASSERT(s == CB_OK, "cb_init retourne %d != CB_OK", s);

    float in[3] = {1.5f, 2.5f, 3.5f};
    for (int i = 0; i < 3; i++) {
        s = cb_push(&cb, &in[i]);
        TEST_ASSERT(s == CB_OK, "push float[%d] retourne %d != CB_OK", i, s);
    }

    for (int i = 0; i < 3; i++) {
        float out = 0.0f;
        s = cb_pop(&cb, &out);
        TEST_ASSERT(s == CB_OK,                  "pop float[%d] retourne %d != CB_OK", i, s);
        TEST_ASSERT(fabsf(out - in[i]) < 1e-6f,  "pop float[%d]=%.6f attendu %.6f", i, (double)out, (double)in[i]);
    }

    TEST_ASSERT(cb.count == 0u, "count=%u != 0 apres pop complet", (unsigned)cb.count);

    s = cb_free(&cb);
    TEST_ASSERT(s == CB_OK, "cb_free retourne %d != CB_OK", s);
    snprintf(tc->detail, sizeof(tc->detail), "elem_size=float: {1.5,2.5,3.5} push/pop OK");
    tc->result = R_PASS;
}

/* ========================================================================
 * T10 – Cycle remplissage / vidange repete
 * ======================================================================== */
void CB_seq_test_t10_fill_drain_cycle(TEST_case_t *tc) {
    tc->result = R_FAIL;

#define CB_TEST_CAP 8u

    uint32_t storage[CB_TEST_CAP];
    circular_buffer_t cb;
    cb_status_t s = cb_init(&cb, storage, sizeof(uint32_t), CB_TEST_CAP, CB_REJECT_NEW);
    TEST_ASSERT(s == CB_OK, "cb_init retourne %d != CB_OK", s);

    /* --- Cycle 1 : valeurs 0..7 --- */
    for (uint32_t i = 0u; i < CB_TEST_CAP; i++) {
        s = cb_push(&cb, &i);
        TEST_ASSERT(s == CB_OK, "[C1] push[%u] retourne %d != CB_OK", (unsigned)i, s);
    }
    TEST_ASSERT(cb.count == CB_TEST_CAP, "[C1] count=%u != %u", (unsigned)cb.count, (unsigned)CB_TEST_CAP);

    for (uint32_t i = 0u; i < CB_TEST_CAP; i++) {
        uint32_t out = 0xFFu;
        s = cb_pop(&cb, &out);
        TEST_ASSERT(s   == CB_OK, "[C1] pop[%u] retourne %d != CB_OK", (unsigned)i, s);
        TEST_ASSERT(out == i,     "[C1] pop[%u]=%u attendu %u", (unsigned)i, (unsigned)out, (unsigned)i);
    }
    TEST_ASSERT(cb.count == 0u, "[C1] count=%u != 0 apres vidange", (unsigned)cb.count);

    /* --- Cycle 2 : valeurs 8..15 --- */
    for (uint32_t i = 0u; i < CB_TEST_CAP; i++) {
        uint32_t v = i + CB_TEST_CAP;
        s = cb_push(&cb, &v);
        TEST_ASSERT(s == CB_OK, "[C2] push[%u] retourne %d != CB_OK", (unsigned)v, s);
    }

    for (uint32_t i = 0u; i < CB_TEST_CAP; i++) {
        uint32_t out = 0xFFu;
        uint32_t expected = i + CB_TEST_CAP;
        s = cb_pop(&cb, &out);
        TEST_ASSERT(s   == CB_OK,         "[C2] pop[%u] retourne %d != CB_OK", (unsigned)i, s);
        TEST_ASSERT(out == expected,       "[C2] pop[%u]=%u attendu %u", (unsigned)i, (unsigned)out, (unsigned)expected);
    }
    TEST_ASSERT(cb.count == 0u, "[C2] count=%u != 0 apres vidange", (unsigned)cb.count);

#undef CB_TEST_CAP

    s = cb_free(&cb);
    TEST_ASSERT(s == CB_OK, "cb_free retourne %d != CB_OK", s);
    snprintf(tc->detail, sizeof(tc->detail), "2 cycles fill/drain cap=8 OK, count=0 final");
    tc->result = R_PASS;
}
