#include "dt_seq_test.h"

#include <stdio.h>
#include <string.h>
#include <stdint.h>

/* ========================================================================
 * Table des cas de test
 * ======================================================================== */

TEST_case_table_t DT_seq_test_cases[DT_seq_test_N_TESTS] = {
    { .case_info = { .name = "T0  NULL args"        }, .func = DT_seq_test_t0_null_args         },
    { .case_info = { .name = "T1  Pub/Read FIFO"    }, .func = DT_seq_test_t1_publish_read_fifo },
    { .case_info = { .name = "T2  Empty read"       }, .func = DT_seq_test_t2_empty_read        },
    { .case_info = { .name = "T3  Attach OLDEST"    }, .func = DT_seq_test_t3_attach_from_oldest},
    { .case_info = { .name = "T4  Attach NOW"       }, .func = DT_seq_test_t4_attach_from_now   },
    { .case_info = { .name = "T5  num_to_read"      }, .func = DT_seq_test_t5_num_to_read       },
    { .case_info = { .name = "T6  Two subscribers"  }, .func = DT_seq_test_t6_two_subscribers   },
    { .case_info = { .name = "T7  Data loss"        }, .func = DT_seq_test_t7_data_loss         },
    { .case_info = { .name = "T8  Sync"             }, .func = DT_seq_test_t8_sync              },
    { .case_info = { .name = "T9  Detach/Reattach"  }, .func = DT_seq_test_t9_detach_reattach   },
    { .case_info = { .name = "T10 Peek non-dest."   }, .func = DT_seq_test_t10_peek_non_destructive },
    { .case_info = { .name = "T11 REJECT_NEW"       }, .func = DT_seq_test_t11_reject_new_policy},
};

/* ========================================================================
 * T0 – Protection NULL / arguments invalides
 * ======================================================================== */
void DT_seq_test_t0_null_args(TEST_case_t *tc) {
    tc->result = R_FAIL;

    data_status_t s;

    /* data_topic_publish avec pointeurs NULL */
    uint32_t storage[4];
    data_topic_t topic;
    data_topic_init(&topic, storage, sizeof(uint32_t), 4, CB_REJECT_NEW);

    s = data_topic_publish(NULL, &(uint32_t){1u});
    TEST_ASSERT(s == DT_BAD_ARG, "publish(NULL,val) retourne %d != DT_BAD_ARG", s);

    s = data_topic_publish(&topic, NULL);
    TEST_ASSERT(s == DT_BAD_ARG, "publish(topic,NULL) retourne %d != DT_BAD_ARG", s);

    /* data_sub_attach avec pointeurs NULL */
    data_sub_t sub = {0};
    s = data_sub_attach(NULL, &topic, DATA_ATTACH_FROM_NOW);
    TEST_ASSERT(s == DT_BAD_ARG, "attach(NULL,topic) retourne %d != DT_BAD_ARG", s);

    s = data_sub_attach(&sub, NULL, DATA_ATTACH_FROM_NOW);
    TEST_ASSERT(s == DT_BAD_ARG, "attach(sub,NULL) retourne %d != DT_BAD_ARG", s);

    /* data_sub_detach sur abonné non attaché ou NULL */
    s = data_sub_detach(NULL);
    TEST_ASSERT(s == DT_BAD_ARG, "detach(NULL) retourne %d != DT_BAD_ARG", s);
    /* sub.attached == 0 */
    s = data_sub_detach(&sub);
    TEST_ASSERT(s == DT_BAD_ARG, "detach(non-attache) retourne %d != DT_BAD_ARG", s);

    /* data_sub_sync sur NULL ou non attaché */
    s = data_sub_sync(NULL);
    TEST_ASSERT(s == DT_BAD_ARG, "sync(NULL) retourne %d != DT_BAD_ARG", s);
    s = data_sub_sync(&sub);
    TEST_ASSERT(s == DT_BAD_ARG, "sync(non-attache) retourne %d != DT_BAD_ARG", s);

    /* data_sub_read avec NULL */
    uint32_t out;
    s = data_sub_read(NULL, &out);
    TEST_ASSERT(s == DT_BAD_ARG, "read(NULL,out) retourne %d != DT_BAD_ARG", s);

    /* data_sub_read sur non attaché */
    s = data_sub_read(&sub, &out);
    TEST_ASSERT(s == DT_BAD_ARG, "read(non-attache,out) retourne %d != DT_BAD_ARG", s);

    data_topic_free(&topic);
    snprintf(tc->detail, sizeof(tc->detail), "Tous les NULL/bad-arg correctement rejetes");
    tc->result = R_PASS;
}

/* ========================================================================
 * T1 – Publish / read FIFO basique
 * ======================================================================== */
void DT_seq_test_t1_publish_read_fifo(TEST_case_t *tc) {
    tc->result = R_FAIL;

    uint32_t storage[4];
    data_topic_t topic;
    data_topic_init(&topic, storage, sizeof(uint32_t), 4, CB_OVERWRITE_OLDEST);

    data_sub_t sub = {0};
    data_status_t s = data_sub_attach(&sub, &topic, DATA_ATTACH_FROM_NOW);
    TEST_ASSERT(s == DT_OK, "attach retourne %d != DT_OK", s);

    uint32_t in[3] = {10u, 20u, 30u};
    for (int i = 0; i < 3; i++) {
        s = data_topic_publish(&topic, &in[i]);
        TEST_ASSERT(s == DT_OK, "publish[%d] retourne %d != DT_OK", i, s);
    }
    TEST_ASSERT(topic.pub_seq == 3u, "pub_seq=%u != 3", (unsigned)topic.pub_seq);
    TEST_ASSERT(data_sub_num_to_read(&sub) == 3u,
              "num_to_read=%u != 3 apres 3 publish", (unsigned)data_sub_num_to_read(&sub));

    for (int i = 0; i < 3; i++) {
        uint32_t out = 0u;
        s = data_sub_read(&sub, &out);
        TEST_ASSERT(s   == DT_OK,  "read[%d] retourne %d != DT_OK", i, s);
        TEST_ASSERT(out == in[i],  "read[%d]=%u attendu %u", i, (unsigned)out, (unsigned)in[i]);
    }
    TEST_ASSERT(data_sub_num_to_read(&sub) == 0u,
              "num_to_read=%u != 0 apres 3 lectures", (unsigned)data_sub_num_to_read(&sub));

    data_sub_detach(&sub);
    data_topic_free(&topic);
    snprintf(tc->detail, sizeof(tc->detail), "FIFO {10,20,30} OK, pub_seq=3, num_to_read=0");
    tc->result = R_PASS;
}

/* ========================================================================
 * T2 – DT_EMPTY quand aucune donnée disponible
 * ======================================================================== */
void DT_seq_test_t2_empty_read(TEST_case_t *tc) {
    tc->result = R_FAIL;

    uint32_t storage[4];
    data_topic_t topic;
    data_topic_init(&topic, storage, sizeof(uint32_t), 4, CB_REJECT_NEW);

    data_sub_t sub = {0};
    data_sub_attach(&sub, &topic, DATA_ATTACH_FROM_NOW);

    TEST_ASSERT(data_sub_num_to_read(&sub) == 0u,
              "num_to_read=%u != 0 antes de cualquier pub", (unsigned)data_sub_num_to_read(&sub));

    uint32_t out = 0xDEADBEEFu;
    data_status_t s = data_sub_read(&sub, &out);
    TEST_ASSERT(s == DT_EMPTY, "read sur vide retourne %d != DT_EMPTY", s);
    TEST_ASSERT(out == 0xDEADBEEFu, "out modifié alors que DT_EMPTY (out=0x%08X)", (unsigned)out);

    s = data_sub_peek(&sub, &out, 0);
    TEST_ASSERT(s == DT_EMPTY, "peek sur vide retourne %d != DT_EMPTY", s);

    data_sub_detach(&sub);
    data_topic_free(&topic);
    snprintf(tc->detail, sizeof(tc->detail), "DT_EMPTY pour read et peek, out non modifie");
    tc->result = R_PASS;
}

/* ========================================================================
 * T3 – DATA_ATTACH_FROM_OLDEST récupère l'historique existant
 * ======================================================================== */
void DT_seq_test_t3_attach_from_oldest(TEST_case_t *tc) {
    tc->result = R_FAIL;

    uint32_t storage[4];
    data_topic_t topic;
    data_topic_init(&topic, storage, sizeof(uint32_t), 4, CB_OVERWRITE_OLDEST);

    uint32_t in[3] = {10u, 20u, 30u};
    for (int i = 0; i < 3; i++) {
        data_topic_publish(&topic, &in[i]);
    }

    /* Attache APRES les 3 publications */
    data_sub_t sub = {0};
    data_status_t s = data_sub_attach(&sub, &topic, DATA_ATTACH_FROM_OLDEST);
    TEST_ASSERT(s == DT_OK, "attach FROM_OLDEST retourne %d != DT_OK", s);
    TEST_ASSERT(data_sub_num_to_read(&sub) == 3u,
              "num_to_read=%u != 3 apres ATTACH_FROM_OLDEST", (unsigned)data_sub_num_to_read(&sub));

    for (int i = 0; i < 3; i++) {
        uint32_t out = 0u;
        s = data_sub_read(&sub, &out);
        TEST_ASSERT(s   == DT_OK,  "read[%d] retourne %d != DT_OK", i, s);
        TEST_ASSERT(out == in[i],  "read[%d]=%u attendu %u", i, (unsigned)out, (unsigned)in[i]);
    }
    s = data_sub_read(&sub, &(uint32_t){0u});
    TEST_ASSERT(s == DT_EMPTY, "4eme read retourne %d != DT_EMPTY", s);

    data_sub_detach(&sub);
    data_topic_free(&topic);
    snprintf(tc->detail, sizeof(tc->detail), "FROM_OLDEST: num=3, {10,20,30} lus, DT_EMPTY OK");
    tc->result = R_PASS;
}

/* ========================================================================
 * T4 – DATA_ATTACH_FROM_NOW ignore l'historique existant
 * ======================================================================== */
void DT_seq_test_t4_attach_from_now(TEST_case_t *tc) {
    tc->result = R_FAIL;

    uint32_t storage[4];
    data_topic_t topic;
    data_topic_init(&topic, storage, sizeof(uint32_t), 4, CB_OVERWRITE_OLDEST);

    uint32_t in[3] = {10u, 20u, 30u};
    for (int i = 0; i < 3; i++) {
        data_topic_publish(&topic, &in[i]);
    }

    /* Attache APRES les 3 publications en mode FROM_NOW */
    data_sub_t sub = {0};
    data_sub_attach(&sub, &topic, DATA_ATTACH_FROM_NOW);
    TEST_ASSERT(data_sub_num_to_read(&sub) == 0u,
              "num_to_read=%u != 0 apres ATTACH_FROM_NOW", (unsigned)data_sub_num_to_read(&sub));

    data_status_t s = data_sub_read(&sub, &(uint32_t){0u});
    TEST_ASSERT(s == DT_EMPTY, "read avant any new pub retourne %d != DT_EMPTY", s);

    /* Publie une nouvelle valeur : l'abonné doit la voir */
    uint32_t v40 = 40u;
    data_topic_publish(&topic, &v40);
    TEST_ASSERT(data_sub_num_to_read(&sub) == 1u,
              "num_to_read=%u != 1 apres pub 40", (unsigned)data_sub_num_to_read(&sub));

    uint32_t out = 0u;
    s = data_sub_read(&sub, &out);
    TEST_ASSERT(s   == DT_OK,  "read apres pub 40 retourne %d != DT_OK", s);
    TEST_ASSERT(out == 40u,    "read=%u attendu 40", (unsigned)out);

    data_sub_detach(&sub);
    data_topic_free(&topic);
    snprintf(tc->detail, sizeof(tc->detail), "FROM_NOW: historique ignore, pub 40 lue OK");
    tc->result = R_PASS;
}

/* ========================================================================
 * T5 – data_sub_num_to_read suit le compteur de publications
 * ======================================================================== */
void DT_seq_test_t5_num_to_read(TEST_case_t *tc) {
    tc->result = R_FAIL;

    uint32_t storage[8];
    data_topic_t topic;
    data_topic_init(&topic, storage, sizeof(uint32_t), 8, CB_REJECT_NEW);

    data_sub_t sub = {0};
    data_sub_attach(&sub, &topic, DATA_ATTACH_FROM_NOW);

    uint32_t v;
    v = 1u; data_topic_publish(&topic, &v);
    TEST_ASSERT(data_sub_num_to_read(&sub) == 1u, "[1pub] num=%u != 1",
              (unsigned)data_sub_num_to_read(&sub));

    v = 2u; data_topic_publish(&topic, &v);
    TEST_ASSERT(data_sub_num_to_read(&sub) == 2u, "[2pub] num=%u != 2",
              (unsigned)data_sub_num_to_read(&sub));

    v = 3u; data_topic_publish(&topic, &v);
    TEST_ASSERT(data_sub_num_to_read(&sub) == 3u, "[3pub] num=%u != 3",
              (unsigned)data_sub_num_to_read(&sub));

    /* Chaque read décrémente d'un */
    for (uint32_t expected = 2u; expected != (uint32_t)-1u; expected--) {
        data_sub_read(&sub, &(uint32_t){0u});
        uint32_t n = data_sub_num_to_read(&sub);
        TEST_ASSERT(n == expected, "apres read: num=%u attendu %u",
                  (unsigned)n, (unsigned)expected);
        if (expected == 0u) break;
    }

    data_sub_detach(&sub);
    data_topic_free(&topic);
    snprintf(tc->detail, sizeof(tc->detail), "num_to_read: 1->2->3 puis 2->1->0 apres reads OK");
    tc->result = R_PASS;
}

/* ========================================================================
 * T6 – Indépendance de deux abonnés
 * ======================================================================== */
void DT_seq_test_t6_two_subscribers(TEST_case_t *tc) {
    tc->result = R_FAIL;

    uint32_t storage[4];
    data_topic_t topic;
    data_topic_init(&topic, storage, sizeof(uint32_t), 4, CB_OVERWRITE_OLDEST);

    data_sub_t sub1 = {0}, sub2 = {0};
    data_sub_attach(&sub1, &topic, DATA_ATTACH_FROM_NOW);
    data_sub_attach(&sub2, &topic, DATA_ATTACH_FROM_NOW);
    TEST_ASSERT(topic.sub_count == 2u, "sub_count=%u != 2", (unsigned)topic.sub_count);

    uint32_t in[3] = {10u, 20u, 30u};
    for (int i = 0; i < 3; i++) data_topic_publish(&topic, &in[i]);

    /* sub1 lit les 2 premiers */
    for (int i = 0; i < 2; i++) {
        uint32_t out = 0u;
        data_status_t s = data_sub_read(&sub1, &out);
        TEST_ASSERT(s   == DT_OK,  "[sub1] read[%d] retourne %d != DT_OK", i, s);
        TEST_ASSERT(out == in[i],  "[sub1] read[%d]=%u attendu %u",
                  i, (unsigned)out, (unsigned)in[i]);
    }

    /* sub2 ne doit pas avoir avancé */
    TEST_ASSERT(data_sub_num_to_read(&sub2) == 3u,
              "[sub2] num=%u != 3 (avance par sub1?)", (unsigned)data_sub_num_to_read(&sub2));
    TEST_ASSERT(data_sub_num_to_read(&sub1) == 1u,
              "[sub1] num=%u != 1 apres 2 reads", (unsigned)data_sub_num_to_read(&sub1));

    /* sub2 lit les 3 éléments indépendamment */
    for (int i = 0; i < 3; i++) {
        uint32_t out = 0u;
        data_status_t s = data_sub_read(&sub2, &out);
        TEST_ASSERT(s   == DT_OK,  "[sub2] read[%d] retourne %d != DT_OK", i, s);
        TEST_ASSERT(out == in[i],  "[sub2] read[%d]=%u attendu %u",
                  i, (unsigned)out, (unsigned)in[i]);
    }
    TEST_ASSERT(data_sub_num_to_read(&sub2) == 0u,
              "[sub2] num=%u != 0 apres 3 reads", (unsigned)data_sub_num_to_read(&sub2));

    data_sub_detach(&sub1);
    data_sub_detach(&sub2);
    data_topic_free(&topic);
    snprintf(tc->detail, sizeof(tc->detail),
             "sub1={10,20} lu, sub2={10,20,30} intact, tails indep.");
    tc->result = R_PASS;
}

/* ========================================================================
 * T7 – DT_DATA_LOSS quand un abonné est dépassé
 * ======================================================================== */
void DT_seq_test_t7_data_loss(TEST_case_t *tc) {
    tc->result = R_FAIL;

#define CAP7 3u
    uint32_t storage[CAP7];
    data_topic_t topic;
    data_topic_init(&topic, storage, sizeof(uint32_t), CAP7, CB_OVERWRITE_OLDEST);

    data_sub_t sub = {0};
    data_sub_attach(&sub, &topic, DATA_ATTACH_FROM_NOW);
    /* sub.tail=0, sub.last_seq=0 */

    /* Publie CAP+1 éléments : delta = 4 > cap = 3 → data loss */
    for (uint32_t i = 1u; i <= CAP7 + 1u; i++) {
        data_topic_publish(&topic, &i);
    }
    TEST_ASSERT(data_sub_num_to_read(&sub) > CAP7,
              "num_to_read=%u devrait etre > cap=%u",
              (unsigned)data_sub_num_to_read(&sub), (unsigned)CAP7);

    /* Utilise peek_relative_ptr (bas niveau) pour observer DT_DATA_LOSS */
    const void *ptr = NULL;
    data_status_t s = data_sub_peek_relative_ptr(&sub, &ptr, sub.tail, 0);
    TEST_ASSERT(s == DT_DATA_LOSS,
              "peek_relative_ptr retourne %d != DT_DATA_LOSS apres overflow", s);

    /* La sync automatique doit avoir eu lieu */
    TEST_ASSERT(data_sub_num_to_read(&sub) == 0u,
              "num_to_read=%u != 0 apres sync automatique",
              (unsigned)data_sub_num_to_read(&sub));

    /* Après sync : nouvelles publications lisibles normalement */
    uint32_t v99 = 99u;
    data_topic_publish(&topic, &v99);
    TEST_ASSERT(data_sub_num_to_read(&sub) == 1u,
              "num_to_read=%u != 1 apres pub post-sync",
              (unsigned)data_sub_num_to_read(&sub));

    uint32_t out = 0u;
    s = data_sub_read(&sub, &out);
    TEST_ASSERT(s   == DT_OK, "read post-sync retourne %d != DT_OK", s);
    TEST_ASSERT(out == 99u,   "read post-sync=%u attendu 99", (unsigned)out);

    data_sub_detach(&sub);
    data_topic_free(&topic);
    snprintf(tc->detail, sizeof(tc->detail),
             "DT_DATA_LOSS detecte, sync auto, lecture post-sync=99 OK");
    tc->result = R_PASS;
#undef CAP7
}

/* ========================================================================
 * T8 – data_sub_sync réaligne sur la tête courante
 * ======================================================================== */
void DT_seq_test_t8_sync(TEST_case_t *tc) {
    tc->result = R_FAIL;

    uint32_t storage[4];
    data_topic_t topic;
    data_topic_init(&topic, storage, sizeof(uint32_t), 4, CB_OVERWRITE_OLDEST);

    data_sub_t sub = {0};
    data_sub_attach(&sub, &topic, DATA_ATTACH_FROM_NOW);

    uint32_t in[3] = {10u, 20u, 30u};
    for (int i = 0; i < 3; i++) data_topic_publish(&topic, &in[i]);
    TEST_ASSERT(data_sub_num_to_read(&sub) == 3u,
              "num_to_read=%u != 3 avant sync", (unsigned)data_sub_num_to_read(&sub));

    /* Sync → l'abonné saute tout l'historique non lu */
    data_status_t s = data_sub_sync(&sub);
    TEST_ASSERT(s == DT_OK, "sync retourne %d != DT_OK", s);
    TEST_ASSERT(data_sub_num_to_read(&sub) == 0u,
              "num_to_read=%u != 0 apres sync", (unsigned)data_sub_num_to_read(&sub));
    TEST_ASSERT(sub.tail     == topic.cb.head,
              "sub.tail=%u != cb.head=%u apres sync",
              (unsigned)sub.tail, (unsigned)topic.cb.head);
    TEST_ASSERT(sub.last_seq == topic.pub_seq,
              "sub.last_seq=%u != pub_seq=%u apres sync",
              (unsigned)sub.last_seq, (unsigned)topic.pub_seq);

    s = data_sub_read(&sub, &(uint32_t){0u});
    TEST_ASSERT(s == DT_EMPTY, "read apres sync retourne %d != DT_EMPTY", s);

    /* Publie 40 : lisible normalement après sync */
    uint32_t v40 = 40u;
    data_topic_publish(&topic, &v40);
    uint32_t out = 0u;
    s = data_sub_read(&sub, &out);
    TEST_ASSERT(s   == DT_OK, "read post-sync retourne %d != DT_OK", s);
    TEST_ASSERT(out == 40u,   "read post-sync=%u attendu 40", (unsigned)out);

    data_sub_detach(&sub);
    data_topic_free(&topic);
    snprintf(tc->detail, sizeof(tc->detail),
             "Sync: num=0, tail=head, last_seq=pub_seq, pub 40 lue OK");
    tc->result = R_PASS;
}

/* ========================================================================
 * T9 – data_sub_detach / sub_count / re-attachement
 * ======================================================================== */
void DT_seq_test_t9_detach_reattach(TEST_case_t *tc) {
    tc->result = R_FAIL;

    uint32_t storage[4];
    data_topic_t topic;
    data_topic_init(&topic, storage, sizeof(uint32_t), 4, CB_OVERWRITE_OLDEST);

    data_sub_t sub1 = {0}, sub2 = {0};
    data_sub_attach(&sub1, &topic, DATA_ATTACH_FROM_NOW);
    data_sub_attach(&sub2, &topic, DATA_ATTACH_FROM_NOW);
    TEST_ASSERT(topic.sub_count == 2u, "sub_count=%u != 2 apres 2 attach",
              (unsigned)topic.sub_count);

    /* Détache sub1 */
    data_status_t s = data_sub_detach(&sub1);
    TEST_ASSERT(s == DT_OK, "detach sub1 retourne %d != DT_OK", s);
    TEST_ASSERT(topic.sub_count == 1u, "sub_count=%u != 1 apres detach",
              (unsigned)topic.sub_count);
    TEST_ASSERT(sub1.attached == 0, "sub1.attached=%d != 0", sub1.attached);

    /* sub1 détaché ne peut pas lire */
    s = data_sub_read(&sub1, &(uint32_t){0u});
    TEST_ASSERT(s == DT_BAD_ARG, "read sur sub1 detache retourne %d != DT_BAD_ARG", s);

    /* sub2 voit correctement la publication */
    uint32_t v99 = 99u;
    data_topic_publish(&topic, &v99);
    uint32_t out = 0u;
    s = data_sub_read(&sub2, &out);
    TEST_ASSERT(s   == DT_OK, "sub2 read retourne %d != DT_OK", s);
    TEST_ASSERT(out == 99u,   "sub2 read=%u attendu 99", (unsigned)out);

    /* Re-attache sub1 */
    s = data_sub_attach(&sub1, &topic, DATA_ATTACH_FROM_NOW);
    TEST_ASSERT(s == DT_OK, "re-attach sub1 retourne %d != DT_OK", s);
    TEST_ASSERT(topic.sub_count == 2u, "sub_count=%u != 2 apres re-attach",
              (unsigned)topic.sub_count);
    TEST_ASSERT(data_sub_num_to_read(&sub1) == 0u,
              "sub1 num_to_read=%u != 0 apres re-attach FROM_NOW",
              (unsigned)data_sub_num_to_read(&sub1));

    data_sub_detach(&sub1);
    data_sub_detach(&sub2);
    data_topic_free(&topic);
    snprintf(tc->detail, sizeof(tc->detail),
             "Detach: sub_count=1, sub2 lit 99, re-attach OK, count=2");
    tc->result = R_PASS;
}

/* ========================================================================
 * T10 – data_sub_peek : lecture non destructive
 * ======================================================================== */
void DT_seq_test_t10_peek_non_destructive(TEST_case_t *tc) {
    tc->result = R_FAIL;

    uint32_t storage[4];
    data_topic_t topic;
    data_topic_init(&topic, storage, sizeof(uint32_t), 4, CB_OVERWRITE_OLDEST);

    data_sub_t sub = {0};
    data_sub_attach(&sub, &topic, DATA_ATTACH_FROM_NOW);

    uint32_t in[3] = {10u, 20u, 30u};
    for (int i = 0; i < 3; i++) data_topic_publish(&topic, &in[i]);

    size_t tail_before = sub.tail;

    /* Peek aux index physiques 0,1,2 : lit le slot physique du CB */
    uint32_t expected[3] = {10u, 20u, 30u};
    for (int i = 0; i < 3; i++) {
        uint32_t out = 0u;
        data_status_t s = data_sub_peek(&sub, &out, i);
        TEST_ASSERT(s   == DT_OK,        "peek[%d] retourne %d != DT_OK", i, s);
        TEST_ASSERT(out == expected[i],   "peek[%d]=%u attendu %u",
                  i, (unsigned)out, (unsigned)expected[i]);
    }

    /* Non destructif : tail et num_to_read inchangés */
    TEST_ASSERT(sub.tail == tail_before,
              "sub.tail=%u modifie par peek (etait %u)",
              (unsigned)sub.tail, (unsigned)tail_before);
    TEST_ASSERT(data_sub_num_to_read(&sub) == 3u,
              "num_to_read=%u != 3 apres peeks", (unsigned)data_sub_num_to_read(&sub));

    /* data_sub_read suivant retourne le premier element du FIFO */
    uint32_t out = 0u;
    data_status_t s = data_sub_read(&sub, &out);
    TEST_ASSERT(s   == DT_OK, "read apres peeks retourne %d != DT_OK", s);
    TEST_ASSERT(out == 10u,   "read apres peeks=%u attendu 10 (tete FIFO)", (unsigned)out);

    data_sub_detach(&sub);
    data_topic_free(&topic);
    snprintf(tc->detail, sizeof(tc->detail),
             "peek[0,1,2]={10,20,30} OK, tail intact, read->10 OK");
    tc->result = R_PASS;
}

/* ========================================================================
 * T11 – Politique CB_REJECT_NEW → DT_FULL à la publication
 * ======================================================================== */
void DT_seq_test_t11_reject_new_policy(TEST_case_t *tc) {
    tc->result = R_FAIL;

    uint32_t storage[2];
    data_topic_t topic;
    data_topic_init(&topic, storage, sizeof(uint32_t), 2, CB_REJECT_NEW);

    data_sub_t sub = {0};
    data_sub_attach(&sub, &topic, DATA_ATTACH_FROM_NOW);

    uint32_t v;
    data_status_t s;

    v = 10u; s = data_topic_publish(&topic, &v);
    TEST_ASSERT(s == DT_OK, "publish 10 retourne %d != DT_OK", s);
    v = 20u; s = data_topic_publish(&topic, &v);
    TEST_ASSERT(s == DT_OK, "publish 20 retourne %d != DT_OK", s);

    TEST_ASSERT(topic.pub_seq == 2u, "pub_seq=%u != 2 avant reject", (unsigned)topic.pub_seq);

    v = 30u; s = data_topic_publish(&topic, &v);
    TEST_ASSERT(s == DT_FULL, "publish 30 retourne %d != DT_FULL (buffer plein)", s);

    /* pub_seq ne doit pas avoir incrementé (publication refusée) */
    TEST_ASSERT(topic.pub_seq == 2u, "pub_seq=%u != 2 apres reject", (unsigned)topic.pub_seq);
    TEST_ASSERT(data_sub_num_to_read(&sub) == 2u,
              "num_to_read=%u != 2", (unsigned)data_sub_num_to_read(&sub));

    /* Lecture : seulement {10, 20}, 30 non stocké */
    uint32_t out;
    out = 0u; s = data_sub_read(&sub, &out);
    TEST_ASSERT(s == DT_OK && out == 10u, "1er read: s=%d out=%u attendu 10", s, (unsigned)out);
    out = 0u; s = data_sub_read(&sub, &out);
    TEST_ASSERT(s == DT_OK && out == 20u, "2eme read: s=%d out=%u attendu 20", s, (unsigned)out);
    out = 0u; s = data_sub_read(&sub, &out);
    TEST_ASSERT(s == DT_EMPTY, "3eme read retourne %d != DT_EMPTY", s);

    data_sub_detach(&sub);
    data_topic_free(&topic);
    snprintf(tc->detail, sizeof(tc->detail),
             "DT_FULL ok, pub_seq=2 stable, {10,20} lus, 30 rejete");
    tc->result = R_PASS;
}
