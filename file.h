#include <stdio.h>
#include <stdlib.h>

typedef struct Element {
    int val;
    struct Element* suivant;
} Element;

typedef struct pile {
    int size;
    Element* sommet;
} pile;

pile* CreePile();
void Empiler(pile** p, int val);
int Depiler(pile** p);
int Sommet(pile* p);
int taille(pile* p);
int EstVide(pile* p);
void afficherPile(pile* p);
void LibererPile(pile* p);


