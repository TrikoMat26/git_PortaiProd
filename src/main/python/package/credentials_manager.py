import keyring
from PySide6.QtWidgets import QInputDialog, QLineEdit
from typing import Tuple, Optional

class CredentialsManager:
    SERVICE_NAME = "PortailProd_SAP"
    
    @staticmethod
    def get_credentials() -> Tuple[str, str]:
        """Récupère les identifiants SAP stockés ou demande à l'utilisateur."""
        username = keyring.get_password(CredentialsManager.SERVICE_NAME, "username")
        password = keyring.get_password(CredentialsManager.SERVICE_NAME, "password")
        
        if not username or not password:
            return CredentialsManager._prompt_credentials()
        
        return username, password
    
    @staticmethod
    def _prompt_credentials() -> Tuple[str, str]:
        """Demande les identifiants à l'utilisateur et les stocke."""
        username, ok = QInputDialog.getText(
            None, 
            "Identifiants SAP",
            "Nom d'utilisateur:",
            QLineEdit.Normal
        )
        if not ok or not username:
            raise ValueError("Username required")

        password, ok = QInputDialog.getText(
            None,
            "Identifiants SAP",
            "Mot de passe:",
            QLineEdit.Password
        )
        if not ok or not password:
            raise ValueError("Password required")

        # Stockage sécurisé des identifiants
        keyring.set_password(CredentialsManager.SERVICE_NAME, "username", username)
        keyring.set_password(CredentialsManager.SERVICE_NAME, "password", password)
        
        return username, password
    
    @staticmethod
    def clear_credentials():
        """Efface les identifiants stockés."""
        try:
            keyring.delete_password(CredentialsManager.SERVICE_NAME, "username")
            keyring.delete_password(CredentialsManager.SERVICE_NAME, "password")
        except keyring.errors.PasswordDeleteError:
            pass  # Ignorer si les credentials n'existent pas