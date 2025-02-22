# business/logic.py

import re
import pandas as pd

from package.ipr_service import run_ipr
from package.sap_service import run_sap_transaction
from package.data_extract_service import load_excel_data
from pathlib import Path
from package.resourcesPath import AppContext
from package.ZP20_json import nom_app
import subprocess

ctx = AppContext.get()
resources_dir = ctx.get_resource('Fichies_config/dummy.txt')
vbs_path = Path(resources_dir).parent / "Transaction.vbs"

class BusinessLogic:
    def __init__(self, parent_window=None):
        self.parent_window = parent_window
        self.current_df = pd.DataFrame()
        self.filtered_df = pd.DataFrame()

        # Mapping action -> argument pour SAP
        self.action_mapping = {
            "1.Nomenclature (ZP20)": "/nzp20",
            "IPR": "IPR",   # Signe qu'on doit aller vers IPR
            "Saisie": "saisie",
            "Note infos": "infos",
            "Temps Gamme": "TempsGamme",
            "Saisie Temps OF": "TempsOF",
            "2.Nomenclature (CS03)": "/ncs03",
            "Modif Temps Gamme": "ModifTempsGamme",
            "COOIS sur OF": "/ncoois_of",
            "COOIS sur REF": "/ncoois_ref",
            "4.MAG Stock (MD04)": "/nmd04",
            "MAG Besoins (MCBZ)": "/nmcbz",
            "3.Nomenclature Code suppérieur (CS15)": "/ncs15",  # attention duplication?
            "6.MAG Mouvement stocks (MD51)": "/nmb51",
            "5.MAG Emplacement (LS24)": "/nls24",
            "7.ZPAN": "/nzpan",
            "Nomenclature interactive": "Nomenclature_interactive"
        }

    def load_data(self):
        """Charge le DataFrame depuis un service spécialisé."""
        df = load_excel_data()  # Appel service/data_extract_service.py
        self.current_df = df
        self.filtered_df = df.copy()

    def filter_data(self, filter_texts):
        """Applique la recherche sur chaque colonne (>=3 char)."""
        df = self.current_df.copy()
        for i, text in enumerate(filter_texts):
            if text:
                safe_pattern = re.escape(text)
                df = df[df.iloc[:, i].astype(str).str.contains(
                    safe_pattern, case=False, na=False, regex=True
                )]
        self.filtered_df = df
        return df

    def reset_filters(self):
        """Réinitialise le DataFrame filtré."""
        self.filtered_df = self.current_df

    def execute_action(self, action_name, input_value):
        """Gère l’exécution d’une action (IPR, SAP, ou Note infos)."""
        arg1 = self.action_mapping.get(action_name, "")

        if arg1 == "IPR":
            # IPR => appel direct
            run_ipr(input_value)
        elif arg1 == "infos":
            # Afficher la note infos
            self.show_infos(input_value)
        elif arg1 == "Nomenclature_interactive":
            subprocess.run(
                ["cscript", str(vbs_path), input_value],
                check=True,
                shell=True,
                encoding='utf-8'
            )
            self.my_zp20_widget = nom_app()
        else:
            # Exécuter la transaction SAP
            run_sap_transaction(arg1, input_value)

    def show_infos(self, of_value):
        """Logique pour afficher la note infos (data sur l'OF)."""
        df = self.filtered_df if not self.filtered_df.empty else self.current_df
        try:
            # Filtre sur la colonne 'OF' ou 'Référence'
            result = df[
                (df['OF'].astype(str) == of_value) |
                (df['Référence'].astype(str) == of_value)
            ][['OF', 'Référence', 'Designation', 'Client']].drop_duplicates()

            if result.empty:
                raise Exception("Aucune donnée trouvée pour cette référence/OF.")

            note_content = "INFOS PRODUCTION\n\n"
            for _, row in result.iterrows():
                note_content += (
                    f"Numéro d'OF: {row['OF']}\n"
                    f"Référence: {row['Référence']}\n"
                    f"Désignation: {row['Designation']}\n"
                    f"Client: {row['Client']}\n"
                    "--------------------------------------------------\n"
                )

            # Création d'un fichier temporaire
            import tempfile, os
            with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.txt', encoding='utf-8') as f:
                f.write(note_content)
                temp_path = f.name
            # Ouverture dans le bloc-notes
            os.startfile(temp_path)

        except Exception as e:
            # On relance l'exception pour remonter le message dans l'UI
            raise Exception(f"Erreur infos : {e}")
