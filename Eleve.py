"""
   ****************************************************************************************************
   * Dans cette partie Nous allons                                                                    *
   *                                                                                                  *
   * - sauvegarder le fichier fileexcel dans une liste d'objets contenant toutes les informationss    *
   *     d'un eleve                                                                                   *
   * - les fonctions contenant :                                                                      *
   *     -- la liste des eleves ayant la moyenne                                                      *
   *     -- la liste des eleves ages de plus de 20ans                                                 *
   *     -- la moyenne de l'ecole                                                                     *
   *     -- le pourcentage des filles                                                                 *
   *     -- le pourcentage des garcons                                                                *
   *     -- la region qui a les meilleurs eleves                                                      *
   *                                                                                                  *
   * **************************************************************************************************
"""
import xlrd
from operator import attrgetter
from collections import Counter

class Eleve:

    nombreEleves = 0

    def __init__(self,nom, prenom, adresse, moyenne, age, region, specialite, sexe):
        self.nom = nom
        self.prenom = prenom
        self.adresse = adresse
        self.moyenne = moyenne
        self.age = age
        self.region = region
        self.specialite = specialite
        self.sexe = sexe


    def __repr__(self):
            return "Eleve({}  {}  {}  {}  {}  {}  {}  {})".format(self.nom, self.prenom, self.adresse, self.moyenne, self.age, self.region, self.specialite, self.sexe)


document = xlrd.open_workbook("excel/fileexcel.xlsx")
feuille = document.sheet_by_index(0)
cols = feuille.ncols
rows = feuille.nrows

def ListeEleve():
    liste = []
    for r in range(1, rows):
        nom = feuille.cell_value(r, 0)
        prenom = feuille.cell_value(r, 1)
        adresse = feuille.cell_value(r, 2)
        moyenne = feuille.cell_value(r, 3)
        age = feuille.cell_value(r, 4)
        region = feuille.cell_value(r, 5)
        specialite = feuille.cell_value(r, 6)
        sexe = feuille.cell_value(r, 7)
        element = Eleve(nom, prenom, adresse, moyenne, age, region, specialite, sexe)
        liste += [element]
    return liste

def ListeEleveAyantLaMoyenne(): #fonctio retournant la liste des eleves ayant la moyenne
    liste= ListeEleve()
    eleveAyantLaMoyenne = []
    for i, element in enumerate(liste):
        if(element.moyenne >=10):
            eleveAyantLaMoyenne += [element]
    return eleveAyantLaMoyenne

def ListeEleveAyantPlus20():#liste des eleves qui ont plus de 20 ans
    liste = ListeEleve()
    eleveAyantPlus20 = []
    for i, element in enumerate(liste):
        if(element.age >20):
            eleveAyantPlus20 += [element]
    return eleveAyantPlus20

def MoyenneEcole():#Moyenne de toute l'école
    liste = ListeEleve()

    somme =0
    nombre= 0
    for i, element in enumerate(liste):
        somme += liste[i].moyenne
        nombre += 1
    return round(somme/nombre, 2)

def PourcentageFille(): #Pourcentage des filles de l'école
    liste= ListeEleve()
    nombreEleve = 0
    nombreFille = 0
    for i, element in enumerate(liste):
        nombreEleve += 1
        if(element.sexe =="F"):
            nombreFille += 1
    return round((nombreFille*100/nombreEleve), 2)

def PourcentageGarcon(): #Pourcentage des garcons de l'éléve
    return 100 - PourcentageFille()

def RegoinAyantPlusForteMoyenne(): #Région qui a les meilleurs élèves
    liste = ListeEleve()
    moyenneParRegion = []
    liste = sorted(liste, key=attrgetter("region"))
    longueur = len(liste)+1
    longueur = longueur

    i = 0
    j = 0
    k = 0
    while(i < longueur - 1):
        moyenne = 0
        compteur = 0
        for j in range(longueur-1):
            if(liste[j].region == liste[i].region):
                moyenne += liste[j].moyenne
                compteur += 1
                k += 1
        element = [liste[i].region, round(moyenne/compteur, 2)]
        moyenneParRegion += [element]
        i = k
    max = moyenneParRegion[0][1]
    i = 0
    while(i < len(moyenneParRegion)):
        if(moyenneParRegion[i][1] > max):
            max = moyenneParRegion[i][1]
            j = i
        i += 1

    return moyenneParRegion[j][0]



def statistique():
    liste = [MoyenneEcole(), PourcentageFille(), PourcentageGarcon(), RegoinAyantPlusForteMoyenne()]
    return liste