
#ifndef TOOLS_H
#define TOOLS_H

#include "stm32l4xx_hal.h"

#include <stdlib.h>
#include <string.h>
#include <math.h> 

// #define MAX_LENGTH_SPI_BUF 1024

typedef union FLOAT_U32_UNION {
    float    float_type;
    uint32_t uint32_type;
} FLOAT_U32_UNION;



void float_format(char* buff, float num, int precision, int width);


// Renvoie le nombre de champ d'une ligne type csv.
int count_nbr_elems(char buffer[], char sep);

// Modifie un [buffer] caracterisant la ligne d'un fichier de csv afin de decouper les differentes chaines
// de caracteres. Remplie un tableau de pointeur pointant vers ces differentes chaines de caracteres.
void set_elems_from_csv(char **elems, char buffer[], char sep, int nbr_elems);

// Remplie un [buffer] caracterisant la ligne d'un fichier de csv a partir d'un tableau de
// pointeur pointant vers les differentes chaines de caracteres.
void set_line_to_csv(char **elems, char buffer[], char sep, int nbr_elems);

#endif // TOOLS_H