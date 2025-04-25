#include "file.h"

pile* CreePile() {
    pile* p = (pile*)malloc(sizeof(pile));
    if (p == NULL) {
        printf("Erreur d'allocation mémoire\n");
        exit(EXIT_FAILURE);
    }
    p->size = 0;
    p->sommet = NULL;
    return p;
}

void Empiler(pile** p, int val) {
    Element* new = (Element*)malloc(sizeof(Element));
    if (new == NULL) {
        printf("Erreur d'allocation mémoire\n");
        exit(EXIT_FAILURE);
    }
    new->val = val;
    new->suivant = (*p)->sommet;
    (*p)->sommet = new;
    (*p)->size++;
}

int Depiler(pile** p) {
    if ((*p)->sommet == NULL) {
        printf("Pile vide\n");
        exit(EXIT_FAILURE);
    }
    Element* current = (*p)->sommet;
    int val = current->val;
    (*p)->sommet = current->suivant;
    free(current);
    (*p)->size--;
    return val;
}

int Sommet(pile* p) {
    if (p->sommet == NULL) {
        printf("Pile vide\n");
        exit(EXIT_FAILURE);
    }
    return p->sommet->val;
}

int taille(pile* p) {
    return p->size;
}

int EstVide(pile* p) {
    return p->sommet == NULL;
}

void afficherPile(pile* p) {
    Element* current = p->sommet;
    printf("Pile :\n");
    while (current != NULL) {
        printf("%d\n", current->val);
        current = current->suivant;
    }
}

void LibererPile(pile* p) {
    while (!EstVide(p)) {
        Depiler(&p);
    }
}
