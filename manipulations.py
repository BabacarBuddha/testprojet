"""
    *************************************************************************
    *    Dans cette partie nous allons extraire puis creer les              *
    *    fichier excel demandes                                             *
    *                                                                       *
    *************************************************************************
"""
from Eleve import *
import xlwt

Ecole = xlwt.Workbook()#Creation du fichier excel

#Pour la liste des eleves ayant la moyenne
moyenne = Ecole.add_sheet("EleveAyantlamoyenne")#Creation d'une feuille
liste1 = ListeEleveAyantLaMoyenne()

for i in range(len(liste1)):
    moyenne.write(i, 0, liste1[i].prenom)
    moyenne.write(i, 1, liste1[i].nom)
    moyenne.write(i, 2, liste1[i].moyenne)

#Pour le liste des eleves ages de plus de 20
age = Ecole.add_sheet("EleveAyantPlusde20")
liste2 = ListeEleveAyantPlus20()
for i in range(len(liste2)):
    age.write(i, 0, liste2[i].prenom)
    age.write(i, 1, liste2[i].nom)
    age.write(i, 2, liste2[i].age)

#Statistique Gloabal de l'école
statist = Ecole.add_sheet("StatistiqueGlobal")

stat = statistique()
for i in range(len(statistique())):
    statist.write(0, i, statistique()[i])


Ecole.save('Ecole.xls')