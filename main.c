#include <stdio.h>
#include <stdlib.h>
#include "file.h"

int main() {
    pile* p = CreePile();
    Empiler(&p, 10);
    Empiler(&p, 20);
    Empiler(&p, 30);

    printf("\nÉtat de la pile après empilement :\n");
    afficherPile(p);

    printf("Taille de la pile : %d\n", taille(p));
    printf("Le sommet de la pile : %d\n", Sommet(p));

    printf("\nDépilement : %d\n", Depiler(&p));
    afficherPile(p);

    printf("Taille de la pile : %d\n", taille(p));
    printf("Dépilement : %d\n", Depiler(&p));
    afficherPile(p);

    printf("Taille de la pile : %d\n", taille(p));

    LibererPile(p);
    free(p);
    p = NULL;

    return 0;
}
