from pathlib import Path
import win32com.client as win32
import tkinter as tk
from tkinter import messagebox, simpledialog
import pythoncom
import sys
from PySide6.QtWidgets import QMessageBox
from PySide6.QtCore import Qt

class NonModalMessageBox:
    def __init__(self):
        self.root = None
        self._current_msg = None  # Pour garder une référence aux messages

    def show_message(self, title, message, icon="info"):
        if not self.root:
            return
            
        # Fermer le message précédent s'il existe
        if self._current_msg:
            self._current_msg.close()
            
        msg = QMessageBox(self.root)
        msg.setWindowTitle(title)
        msg.setText(message)
        msg.setStandardButtons(QMessageBox.Ok)
        msg.setMinimumWidth(400)  # Largeur minimum
        
        if icon == "error":
            msg.setIcon(QMessageBox.Critical)
        elif icon == "warning":
            msg.setIcon(QMessageBox.Warning)
        else:
            msg.setIcon(QMessageBox.Information)
        
        msg.setWindowModality(Qt.NonModal)
        
        # Connecter le signal finished pour nettoyer la référence
        msg.finished.connect(lambda: self._clear_message(msg))
        
        self._current_msg = msg
        msg.show()
        
    def _clear_message(self, msg):
        if self._current_msg == msg:
            self._current_msg = None

message_box = NonModalMessageBox()

def rech_ipr(codecar):
    """Fonction principale qui peut être appelée par d'autres scripts"""
    if not codecar:
        message_box.show_message("Erreur", "Aucun code fourni", "error")
        return

    # Message initial de recherche
    message_box.show_message("Recherche en cours", f"Recherche de l'IPR pour le code {codecar}...")

    pythoncom.CoInitialize()
    codecar = codecar.replace('/', '-')

    base_path = Path(r"S:\Methodes Production")
    # base_path = Path(r"C:\Destination")
    directories = [
        ("0- IPR VALIDE", False),
        ("1- IPR AUTORISEES", False),
        ("2- IPR en COURS", True),
        ("3- IPR ARCHIVES", True)
    ]

    found = False

    for dir_name, _ in directories:
        dir_full = base_path / dir_name
        
        # Recherche fichiers Excel
        xls_files = list(dir_full.glob(f"{codecar}.xls*"))
        if xls_files:
            _handle_excel_find(xls_files[0], dir_name)
            found = True
            break

        # Recherche fichiers Word
        doc_files = list(dir_full.glob(f"{codecar}.doc*"))
        if doc_files:
            _handle_word_find(doc_files[0], dir_name)
            found = True
            break

    if not found:
        message_box.show_message("Information", "Aucun IPR trouvé")

def _handle_excel_find(file_path, dir_name):
    if "ARCHIVES" in dir_name:
        message_box.show_message("Attention", "Code trouvé dans IPR ARCHIVES, ne pas utiliser", "warning")
        return

    msg_mapping = {
        "VALIDE": "IPR VALIDE",
        "AUTORISEES": "IPR AUTORISES", 
        "COURS": "IPR en cours\nN'UTILISER QUE LES POSTES EN VERT"
    }
    key = next(k for k in msg_mapping if k in dir_name)
    message_box.show_message("Information", f"Code trouvé dans {msg_mapping[key]}")

    try:
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = True
        wb = excel.Workbooks.Open(str(file_path))
        if not wb.ReadOnly:
            wb.Save()
    except Exception as e:
        message_box.show_message("Erreur", f"Erreur Excel: {e}", "error")

def _handle_word_find(file_path, dir_name):
    if "ARCHIVES" in dir_name:
        message_box.show_message("Attention", "Code trouvé dans IPR ARCHIVES (Word), ne pas utiliser", "warning")
        return

    msg_mapping = {
        "VALIDE": "IPR VALIDE (Word)",
        "AUTORISEES": "IPR AUTORISES (Word)",
        "COURS": "IPR en cours (Word)\nN'UTILISER QUE LES POSTES EN VERT"
    }
    key = next(k for k in msg_mapping if k in dir_name)
    message_box.show_message("Information", f"Code trouvé dans {msg_mapping[key]}")

    try:
        word = win32.Dispatch('Word.Application')
        word.Visible = True
        doc = word.Documents.Open(str(file_path))
        doc.Activate()
    except Exception as e:
        message_box.show_message("Erreur", f"Erreur Word: {e}", "error")

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    
    codecar = sys.argv[1] if len(sys.argv) > 1 else simpledialog.askstring("Recherche IPR", "Entrez le code à rechercher:")
    
    if codecar:
        rech_ipr(codecar.strip())
    else:
        message_box.show_message("Annulé", "Aucun code saisi")
    
    root.mainloop()
