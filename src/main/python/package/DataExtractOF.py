from pathlib import Path
import datetime
import win32com.client
from package.resourcesPath import AppContext
from PySide6.QtWidgets import QMessageBox

def convert_xls_to_xlsx(chemin_xls: Path, chemin_xlsx: Path):
    """
    Convertit un fichier .xls en .xlsx via Excel COM, sans copie locale du .xls.
    :param chemin_xls: Chemin réseau vers le fichier .xls
    :param chemin_xlsx: Chemin local où enregistrer le .xlsx
    """
    try:
        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.Visible = False  # Excel invisible pendant la conversion
        
        wb = excel_app.Workbooks.Open(str(chemin_xls))
        wb.SaveAs(str(chemin_xlsx), FileFormat=51)  # 51 correspond au format .xlsx
        wb.Close()
        excel_app.Quit()
        
        print(f"Conversion réussie : {chemin_xls.name} -> {chemin_xlsx.name}")
    except Exception as e:
        print(f"Echec de la conversion en .xlsx : {e}")

def get_xlsx_of_the_day():
    """
    Vérifie l'existence d'un fichier .xls sur un lecteur réseau
    (nommé extraction_OF_du_dd_mm_yyyy.xls), le convertit en .xlsx,
    et le stocke dans le même dossier que ce script.

    Retourne le seul nom du fichier .xlsx (ex. 'extraction_OF_du_30_01_2025.xlsx')
    si déjà présent ou converti avec succès. Retourne None si le fichier est
    introuvable ou si la conversion a échoué.
    """
    # Récupération de la date du jour
    aujourdhui = datetime.date.today()
    jour = aujourdhui.day
    mois = aujourdhui.month
    annee = aujourdhui.year

    # Préfixe commun du fichier source / destination
    prefixe_fichier = f"extraction_OF_du_{jour:02d}_{mois:02d}_{annee}"

    # Chemin réseau du .xls
    chemin_source = Path(f"W:/CHARGE_SAP/Extraction_OF/{prefixe_fichier}.xls")
    print("Chemin du fichier source :", chemin_source)

    # Chemin local pour la version .xlsx
    ctx = AppContext.get()
    resources_dir = ctx.get_resource('Fichies_config/dummy.txt')
    chemin_xlsx = Path(resources_dir).parent / f"{prefixe_fichier}.xlsx"

    # 1) Vérification de l'existence locale du .xlsx
    if chemin_xlsx.exists():
        print(f"Le fichier '{chemin_xlsx.name}' existe déjà dans le répertoire du script. Aucune action.")
        # QMessageBox.information(None, "Fichier utilisé", f"Le fichier utilisé est : {chemin_xlsx.name}")
        return chemin_xlsx.name
    else:
        # 2) Vérifier si le .xls existe sur le réseau
        if not chemin_source.exists():
            print("ERREUR : Le fichier source n'existe pas sur le réseau.")
            return None
        else:
            # 3) Conversion directe
            convert_xls_to_xlsx(chemin_source, chemin_xlsx)

            # 4) Vérifier la bonne création du .xlsx
            if chemin_xlsx.exists():
                print("Le fichier .xlsx a été créé correctement.")

                # Suppression des autres fichiers .xlsx dans le répertoire de destination
                destination_dir = chemin_xlsx.parent
                for fichier in destination_dir.glob("*.xlsx"):
                    if fichier != chemin_xlsx:
                        try:
                            fichier.unlink()
                            print(f"Fichier supprimé : {fichier.name}")
                        except Exception as e:
                            print(f"Echec de suppression du fichier {fichier.name} : {e}")

                # QMessageBox.information(None, "Fichier utilisé", f"Le fichier utilisé est : {chemin_xlsx.name}")
                return chemin_xlsx.name
            else:
                print("Le fichier .xlsx n'a pas été créé correctement.")
                return None

if __name__ == "__main__":
    resultat = get_xlsx_of_the_day()
    if resultat:
        print("Nom du fichier du jour :", resultat)
    else:
        print("Pas de fichier créé ou échec lors de la conversion.")
