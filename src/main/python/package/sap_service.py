# services/sap_service.py

import subprocess
from pathlib import Path
from  package.resourcesPath import AppContext

ctx = AppContext.get()
resources_dir = ctx.get_resource('Fichies_config/dummy.txt')

def run_sap_transaction(arg1, arg2):
    """Exécute TransactionSAP.vbs avec les arguments donnés."""
    try:
        # Chemin d'accès au script VBS dans le dossier parent du parent
        vbs_path = Path(resources_dir).parent / "TransactionSAP.vbs"
        if not vbs_path.exists():
            raise FileNotFoundError("Fichier TransactionSAP.vbs introuvable !")

        subprocess.run(
            ["cscript", str(vbs_path), arg1, arg2],
            check=True,
            shell=True,
            encoding='utf-8'
        )
    except subprocess.CalledProcessError as e:
        raise Exception(f"Erreur d'exécution VBS : {str(e)}")
    except Exception as e:
        raise Exception(f"Erreur inattendue SAP : {str(e)}")
