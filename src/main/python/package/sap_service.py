# services/sap_service.py

import subprocess
from pathlib import Path
from package.resourcesPath import AppContext
from package.credentials_manager import CredentialsManager

ctx = AppContext.get()
resources_dir = ctx.get_resource('Fichies_config/dummy.txt')

def run_sap_transaction(arg1, arg2):
    """Exécute TransactionSAP.vbs avec les arguments donnés."""
    try:
        # Récupération des identifiants SAP
        username, password = CredentialsManager.get_credentials()
        
        # Chemin d'accès au script VBS dans le dossier parent du parent
        vbs_path = Path(resources_dir).parent / "TransactionSAP.vbs"
        if not vbs_path.exists():
            raise FileNotFoundError("Fichier TransactionSAP.vbs introuvable !")

        # Exécuter le script VBS avec les identifiants récupérés
        subprocess.run(
            ["cscript", str(vbs_path), arg1, arg2, username, password],
            check=True,
            shell=True,
            encoding='utf-8'
        )
    except subprocess.CalledProcessError as e:
        raise Exception(f"Erreur d'exécution VBS : {str(e)}")
    except ValueError as e:
        # Erreur spécifique si les identifiants ne sont pas fournis
        raise Exception(f"Erreur d'authentification SAP : {str(e)}")
    except Exception as e:
        raise Exception(f"Erreur inattendue SAP : {str(e)}")
