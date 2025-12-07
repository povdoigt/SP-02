#include "tools.h"

#include <stdio.h>

void float_format(char *buff, float num, int precision, int width) {
    float num_abs = num < 0 ? -num : num;

    long int num_up = (int)num_abs;
    long int num_dw = (int)((num_abs - num_up) * pow(10, precision));

    int num_up_width = width - precision - 2; // 1 for the dot and 1 for the sign
    long int max_up = pow(10, num_up_width) - 1;

    num_up = num_up > max_up ? max_up : num_up;

    if (num < 0) {
        sprintf(buff, "-%lu.%lu", num_up, num_dw);
    } else {
        sprintf(buff, "+%lu.%lu", num_up, num_dw);
    }
}


int count_nbr_elems(char buffer[], char sep) {
    int i = 1;
    for (int j = 0; buffer[j]; ++j) {
        if (buffer[j] == sep)
            ++i;
    }
    return i;
}

void set_elems_from_csv(char **elems, char buffer[], char sep, int nbr_elems) {
    int i = 0;
    int count = 0;

    buffer[strlen(buffer) - 1] = '\0';
    while (buffer[i] != '\0' && count < nbr_elems) {
        elems[count] = buffer + i;
        ++count;

        while (buffer[i] != sep && buffer[i] != '\0')
            ++i;

        if (buffer[i] == sep) {
            buffer[i] = '\0';
        }
        ++i;
    }
} 

void set_line_to_csv(char **elems, char buffer[], char sep, int nbr_elems) {
    char str_sep[2] = {sep, '\0'}; 
    buffer[0] = '\0';
    for (int i = 0; i < nbr_elems; ++i) {
        strcat(buffer, elems[i]);
        if (i < nbr_elems - 1)
            strcat(buffer, str_sep);
    }
    strcat(buffer, "\n");
}
