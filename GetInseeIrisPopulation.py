import requests
import zipfile
import io
import os

# UserURL = "https://www.insee.fr/fr/statistiques/8201904?sommaire=8202287"
url = "https://www.insee.fr/fr/statistiques/fichier/8201904/base-cc-evol-struct-pop-2021_csv.zip"

# Répertoire où extraire le contenu
output_dir = "./PopulationIRIS"

def download_and_extract_zip(url, output_dir):
    try:
        # Effectuer la requête GET
        response = requests.get(url)
        response.raise_for_status()  # Vérifie que la requête a réussi

        # Charger le contenu ZIP dans la mémoire
        with zipfile.ZipFile(io.BytesIO(response.content)) as z:
            # Créer le répertoire de sortie s'il n'existe pas
            os.makedirs(output_dir, exist_ok=True)
            
            # Extraire tous les fichiers dans le répertoire de sortie
            z.extractall(output_dir)
            print(f"Fichiers extraits dans le dossier : {output_dir}")
    except requests.exceptions.RequestException as e:
        print(f"Erreur lors du téléchargement : {e}")
    except zipfile.BadZipFile as e:
        print(f"Erreur lors de l'extraction du fichier ZIP : {e}")

# Télécharger et extraire le fichier
download_and_extract_zip(url, output_dir)
