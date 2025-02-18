import shutil
import subprocess
import os
# import SetRefinedSheet
import warnings
warnings.simplefilter("ignore", UserWarning)
# Ignorer les warnings spécifiques d'openpyxl
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.styles.stylesheet")

AdministraveRegionURL = "https://www.meilleursagents.com/prix-immobilier/calvados-14"
dossierSource = f"{AdministraveRegionURL.split("/")[-1]}"
template = "./ModeleRentaDistance.xlsx"
outputFile = f"{dossierSource.split('/')[-1]}.xlsx"
RefinedSheetName = "Refined"
os.chdir(os.path.dirname(os.path.abspath(__file__)))

def setAmenitiesSheets(outputFile):
    departmentToFilter = outputFile.split('.')[0].split("-")[-1]
    subprocess.run(["python", "./SetAmenitiesSheets.py", outputFile, departmentToFilter])

def setPopulationSheet(outputFile):
    departmentToFilter = outputFile.split('.')[0].split("-")[-1]
    subprocess.run(["python", "./SetPopulationSheet.py", outputFile, departmentToFilter])

def setCitiesSheet(outputFile):
    departmentToFilter = outputFile.split('.')[0].split("-")[-1]
    subprocess.run(["python", "./SetCitiesSheet.py", outputFile, departmentToFilter])

def setRefinedSheet(outputFile):
    departmentToFilter = outputFile.split('.')[0].split("-")[-1]
    subprocess.run(["python", "./SetRefinedSheet.py", outputFile, departmentToFilter, AdministraveRegionURL])

def remakeReferenceSheet(outputFile):
    subprocess.run(["python", "./RemakeReferenceSheet.py", outputFile])

def setUnemploymentSheet(outputFile):
    subprocess.run(["python", "./GetUnemployment.py", outputFile])

def setVacantAccomodationSheet(outputFile) : 
    departmentToFilter = outputFile.split('.')[0].split("-")[-1]
    subprocess.run(["python", "./SetVacantAccomodationSheet.py", outputFile, departmentToFilter])

def cleanSheets(outputFile) :   
    subprocess.run(["python", "./CleanSheets.py", outputFile])

def main():
    global outputFile
    # if not os.path.exists(outputFile) : 
    #     print("Dupplication du modele")
    #     shutil.copy(template, outputFile)
    # setVacantAccomodationSheet(outputFile)
    # setRefinedSheet(outputFile)
    # setAmenitiesSheets()  
    # remakeReferenceSheet()
    # setCitiesSheet()
    # setPopulationSheet()

    for item in os.listdir("./data/neighborhood/"):
        dossierSource = f"./data/neighborhood/{item}"
        outputFile = f"{dossierSource.split('/')[-1]}.xlsx"
        if os.path.exists(outputFile) : 
            print(f"Suppression du fichier {outputFile}")
            os.remove(outputFile)
        if not os.path.exists(outputFile) : 
            print("Dupplication du modele")
            shutil.copy(template, outputFile)
        setVacantAccomodationSheet(outputFile)
        setUnemploymentSheet(outputFile)
        setRefinedSheet(outputFile)
        setAmenitiesSheets(outputFile)  
        remakeReferenceSheet(outputFile)
        setCitiesSheet(outputFile)
        setPopulationSheet(outputFile)

    

    # 1. Scraper la région administrative
    # SetRefinedSheet.scrapeAdministrativeRegionURL(AdministraveRegionURL)

    # # 4. Extraire les données et les enregistrer dans Excel
    # wb, feuille_bdd = SetRefinedSheet.charger_ou_creer_excel(outputFile, RefinedSheetName)
    # SetRefinedSheet.effacer_contenu_feuille(feuille_bdd, wb)
    # SetRefinedSheet.extractData(dossierSource, feuille_bdd, wb)
    # SetRefinedSheet.sauvegarder_excel(wb, outputFile)

if __name__ == "__main__":
    main()
