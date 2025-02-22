from package.DataExtractOF import get_xlsx_of_the_day
import pandas as pd
from pathlib import Path
import os
from package.resourcesPath import AppContext

def load_excel_data():
    """Récupère le nom du fichier du jour et charge la DataFrame."""
    fichierExtractOF = get_xlsx_of_the_day()
    ctx = AppContext.get()
    resources_dir = Path(ctx.get_resource('Fichies_config/dummy.txt')).parent
    # Si la fonction retourne None, on utilise le premier .xlsx trouvé dans le répertoire parent
    if fichierExtractOF is None:
        potential_xlsx_files = list(resources_dir.glob("*.xlsx"))
        if not potential_xlsx_files:
            raise FileNotFoundError("Aucun fichier .xlsx trouvé dans le répertoire parent.")
        file_path = potential_xlsx_files[0]
    else:
        # Sinon, on construit le chemin comme d'habitude
        file_path = resources_dir / fichierExtractOF

    if not file_path.exists():
        raise FileNotFoundError(f"Le fichier {file_path} est introuvable.")

    columns_to_load = [0, 1, 2, 3, 7, 9, 10, 11, 13]
    df = pd.read_excel(file_path, usecols=columns_to_load)
    df.attrs['source_file'] = str(file_path)  # Stocke le chemin complet du fichier

    # Réarrangement des colonnes
    df = df[[df.columns[i] for i in [0, 1, 2, 3, 7, 8, 5, 6, 4]]]

    new_column_names = {
        "Article": "Référence",
        "Désignatio": "Designation",
        "Numéroopér": "Opération",
        "Designatio": "Désignation OP",
        "Quantité": "Qté. prévu",
        "Quantitéli": "Qté. livrée"
    }
    df = df.rename(columns=new_column_names)

    # Renommer la dernière colonne en 'DateFinOF'
    df = df.rename(columns={df.columns[-1]: "DateFinOF"})
    return df
