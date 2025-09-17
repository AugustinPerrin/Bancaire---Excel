###############################
#Fichier python implémentant 
#le scrapping d'un fichier
#pdf de releves bancaires,
#récupérant, stockant les données
#et les envoyant dans un xlsx
#Augustin PERRIN
###############################

#Bibliothèques
import pdfplumber                       #lecteur fichier pdf
import re                               #expression régulières           
from openpyxl import Workbook           #obsolète
from openpyxl import load_workbook      #obsolète
import xlwings as xw                    #to write in excel
import bdd                              #BDD des catégories pour tri automatique


class scrap:

    #classe implémentant le scrapping
    #du releves et tout le reste
    #Pour le moment, à chaque instance de scrap est associé un scrapping
    

    def __init__(self):
        #Initialisation des attributs

        self.fichier = ""           
        self.texte = ""             
        self.date = []
        self.credit = {}
        self.debit = {}
        
        self.nom_categories = ["Logement","Fixe","Abonnement","Transport","Nourriture","Equipement","Habit","Divertissement","Autre"]
        self.categories = {nom : 0 for nom in self.nom_categories}
    
    def ouverture(self,nom_fichier):
        #Méthode ouvrant un fichier avec pdfplumber
        #nom_fichier : nom du fichier pdf à ouvrir
        self.fichier = nom_fichier
        with pdfplumber.open(self.fichier) as pdf:
            for page in pdf.pages:
                self.texte += page.extract_text()
    
    def trouver_date(self):
        #Méthode identifiant la date dans le texte scrappé
        #necessite d'avoir ouvert un fichier
        #Pattern pour les dates au format DD "Mois" YYYY
        pattern_date = r'\d{1,2}\s(?:Janvier|Février|Mars|Avril|Mai|Juin|Juillet|Août|Septembre|Octobre|Novembre|Décembre)\s\d{4}'  #travailler sur l'expression régulière à mettre
        
        #On suppose que la première date correspondante est celle de l'émission du fichier
        match = re.search(pattern_date,self.texte,re.IGNORECASE)
        if match:
            self.date = match.group(0)     #Récupère la première occurence trouvée
        else:
            self.date = None
        return self.date
    
    def credit_debit(self):
        #methode stockant les lignes de debit et credit
        
        #recuperation des lignes
        lignes = self.texte.split('\n')
        pattern = r'^\d{2}.\d{2}\s\d{2}.\d{2}'
        liste_lignes = []
        for ligne in lignes:
            if re.match(pattern,ligne.strip()):
                ligne_clean = ligne.strip()
                liste_lignes.append(ligne_clean)
            
                #Traitement
                montant_pattern = r'\s\d+[,]\d{2}'
                bigmontant_pattern = r'\s\d\s\d+[,]\d{2}'
                montant = re.search(montant_pattern,ligne_clean)
                if montant:
                    #convertir le montant en float
                    valeur = float(montant.group(0).replace(',','.'))
                    if "Virement" in ligne_clean or  "Avoir" in ligne_clean:
                        self.credit[ligne_clean[12:]] = valeur          #12 pour enlever la date au début de la ligne
                    else:
                        self.debit[ligne_clean[12:]] = valeur           #pareil
                bigmontant = re.search(bigmontant_pattern,ligne_clean)
                if bigmontant:
                    #convertir le montant en float
                    valeur = float(bigmontant.group(0).replace(" ", "").replace(',','.'))
                    if "Virement" in ligne_clean or "Avoir" in ligne_clean:
                        self.credit[ligne_clean[12:]] = valeur          #12 pour enlever la date au début de la ligne
                    else:
                        self.debit[ligne_clean[12:]] = valeur           #pareil
                
        return liste_lignes
        
    def tri_manuel(self):
        #methode triant les credits et debits dans les categories
        #tri à la main des catégories
        print("Voici les catégories pour les débits :")
        print(self.nom_categories)
        print("A chaque ligne, taper les deux premieres lettres de la categorie correspondante puis 'entrée'")
        code_categorie = ["lo","fi","ab","tr","no","eq","ha","di","au"]
        for key in self.debit:
            print("\n Dépense : \n")
            print(key)
            print("Type input : ")
            code_input = input()
            try:
                indice = code_categorie.index(str(code_input))
                print(self.debit[key])
                self.categories[self.nom_categories[indice]] += self.debit[key]
                print(self.categories[self.nom_categories[indice]])
            except:
                print("Erreur d'input, valeur non prise en compte")
        print("Fin du tri\n")
        print(self.categories)

    def tri_automatique(self):
        #methode triant les debits dans les categories
        #tri automatiquement

        liste_cles = []
        for key,montant in self.debit.items():
            flag = False

            for nom ,mots in bdd.dictionnaire_credit.items():
                for mot in mots:
                    if mot in key.lower():
                        self.categories[nom] += montant
                        liste_cles.append(key)
                        flag = True
                        continue
                    if flag :
                        continue
                if flag:
                    continue
        for key in liste_cles:
            self.debit.pop(key)

    def sheet_excel(self):
        print("Tapez Entrée le cas échéant..")
        test = input()

        wb = Workbook()
        ws = wb.active
        ws.title = "Relevé bancaire"+self.date

        ws['A1'] = "Date du relevé"
        ws['B1'] = self.date

        ws['A3'] = "Catégorie"
        ws['B3'] = "Montant"

        row = 5

        for categorie, montant in self.categories.items():
            ws[f'A{row}'] = categorie
            ws[f'B{row}'] = montant
            row += 1
        
        ws['A{}'.format(row+2)] = "Détail des débits"
        ws['A{}'.format(row+3)] = "Description" 
        ws['B{}'.format(row+3)] = "Montant"
        
        row += 4
        for description, montant in self.debit.items():
            ws[f'A{row}'] = description
            ws[f'B{row}'] = montant
            row += 1
        
        #Sauvegarde du fichier
        filename = f"releve_bancaire_{self.date.replace(' ', '_')}.xlsx"
        wb.save(filename)
        print(f"Fichier Excel créé : {filename}")

    def add_new_sheet(self, existing_file):
        print("Veillez à avoir fermé le doc excel\n tapez Entrée le cas échéant..")
        test = input()

        try:
            ws = xw.Book(existing_file)
            nom_temporaire = "temp"
            ws.sheets.add(nom_temporaire)
            sheet = ws.sheets[nom_temporaire]
            
        
            sheet['A1'].value = "Date du relevé"
            sheet['B1'].value = self.date

            sheet['A3'].value = "Débit"
            sheet['A4'].value = "Catégorie"
            sheet['B4'].value = "Montant"

            row = 5
            for categorie, montant in self.categories.items():
                sheet[f'A{row}'].value = categorie
                sheet[f'B{row}'].value = montant
                row += 1
            
            sheet['D3'].value = "Crédit"
            sheet['D4'].value = "Détail"
            sheet["E4"].value = "Montant"
            row = 5
            for detail, montant in self.credit.items():
                sheet[f'D{row}'].value = detail
                sheet[f'E{row}'].value = montant
                row += 1
            
        except PermissionError:
            print(f"ERREUR: Impossible de sauvegarder le fichier {existing_file}")
            print("Assurez-vous que le fichier n'est pas ouvert dans une autre application")
        except Exception as e:
            print(f"ERREUR OUIO OUOI: {str(e)}")

        #Appel de la macro miseenpage du fichier Excel
        ws.macro('MiseEnPage')()
        print("Modification bien prise en compte")


    def test(self):
        #méthode test pour le développement du code
        print("Debit:",self.debit,"\n")
        print("Crédit:",self.credit)
        print("Avant tri :",self.categories)
        self.re_categories()
        print("Apres tri :",self.categories)

    def modification(self):
        #méthode permettant de modifier un
        #débit ou crédit mal catégorisé
        return None
    
    def miseenpage(self):
        #méthode permettant de lancer
        #le code VBA de miseenpage
        return None

        