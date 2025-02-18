import requests
import os
import pandas as pd

# URL du fichier CSV
url = "https://www.data.gouv.fr/fr/datasets/r/51606633-fb13-4820-b795-9a2a575a72f1"

# Nom des fichiers à enregistrer
csv_filename = "./cities/cities.csv"
xlsx_filename = "./cities/cities.xlsx"

# Créer le dossier s'il n'existe pas
os.makedirs(os.path.dirname(csv_filename), exist_ok=True)

# Télécharger le fichier
response = requests.get(url, stream=True)
response.raise_for_status()  # Vérifie si la requête a réussi

# Écrire le fichier CSV
with open(csv_filename, "wb") as file:
    for chunk in response.iter_content(chunk_size=8192):
        file.write(chunk)

print(f"Fichier CSV téléchargé avec succès : {csv_filename}")

# Convertir CSV en XLSX
df = pd.read_csv(csv_filename, delimiter=",")
df.to_excel(xlsx_filename, index=False)

print(f"Fichier XLSX créé avec succès : {xlsx_filename}")
