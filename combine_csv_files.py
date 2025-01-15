import pandas as pd
import glob
import os

# Définir le chemin du dossier contenant les fichiers .xlsx
folder_path = "C:/Users/MSI/Desktop/hexagone/bigdata/Projet_Scraping_Bi/data"

# Récupérer tous les fichiers Excel (.xlsx) dans le dossier
file_paths = glob.glob(os.path.join(folder_path, '*.xlsx'))

# Vérifier si des fichiers sont trouvés
if not file_paths:
    print("No files found.")
else:
    # Liste pour stocker les DataFrames
    dataframes = []

    # Itérer sur chaque fichier Excel trouvé
    for file_path in file_paths:
        try:
            # Lire le fichier Excel dans un DataFrame
            df = pd.read_excel(file_path)
            
            # Supprimer la colonne 'Column1' 
            columns_to_remove = ['Column1']
            df = df.drop(columns=[col for col in columns_to_remove if col in df.columns], errors='ignore')
            
            # Ajouter une nouvelle colonne 'SiteID' avec le nom du fichier (sans extension)
            SiteID = os.path.splitext(os.path.basename(file_path))[0]
            df['SiteID'] = SiteID
            
            # Ajouter le DataFrame à la liste
            dataframes.append(df)
        except Exception as e:
            # Gérer les erreurs de lecture
            print(f"Error reading {file_path}: {e}")

    # Si des DataFrames ont été chargés, les concaténer
    if dataframes:
        # Concaténer tous les DataFrames en un seul
        combined_df = pd.concat(dataframes, ignore_index=True)
        
        # Ajouter une colonne 'Id' pour numéroter chaque ligne
        combined_df['Id'] = range(1, len(combined_df) + 1)
        
        # Définir le chemin du fichier de sortie
        output_file = os.path.join(folder_path, 'resultat_scraping.xlsx')
        
        # Sauvegarder le DataFrame combiné dans un fichier Excel
        combined_df.to_excel(output_file, index=False)
        
        print(f"File saved: {output_file}")
    else:
        print("No valid data found.")
