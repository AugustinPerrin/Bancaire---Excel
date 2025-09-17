###############################
#Script 
#Scrapper de releves bancaires
#Le nom du fichier à scrapper doit être passé dans l'appel au script
#Augustin PERRIN
###############################
import os.path
import scrapper as sc
import sys
#chemin du dossier parent
parent_dir = os.path.dirname(os.path.dirname(__file__))
#chemin complet
pdf_path = os.path.join(parent_dir, str(sys.argv[1]))
excel_path = "releves_2025.xlsm"

##########
#Script
#Ouvrant le fichier, scrappant les données, réalisant un tri automatique puis laisse l'utilisateur trier manuellement, insère les données dans le fichier et fait la mise
#en forme des données
##########
test = sc.scrap()
test.ouverture(pdf_path)
print(test.trouver_date())
lignes=test.credit_debit()
test.tri_automatique()
test.tri_manuel()
test.add_new_sheet(excel_path)

exit()

######################
#Exemple d'utilisation
######################
#test = sc.scrap()
#test.ouverture(pdf_path)           #ouverture du fichier pdf pour scrapper les données
#test.trouver_date()                #récupération de la date du fichier pdf
#test.credit_debit()                #trie les flux en crédit ou débit
#test.tri_automatique()             #trie automatiquement les crédits dans des catégories
#test.tri_manuel()                  #trie manuellement les crédits restants dans des catégories
#test.add_new_sheet(excel_path)               #Envoie les données dans une nouvelle feuille du fichier spécifié
###############################