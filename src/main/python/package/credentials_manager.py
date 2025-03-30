import os
import json
import base64
import platform
from pathlib import Path
from PySide6.QtWidgets import QInputDialog, QLineEdit
from typing import Tuple, Optional

class CredentialsManager:
    SERVICE_NAME = "PortailProd_SAP"
    CONFIG_DIR = Path(os.path.expanduser("~")) / ".portailprod"
    CREDS_FILE = CONFIG_DIR / "credentials.json"
    # Clé de registre Windows pour stocker les identifiants
    REG_PATH = r"Software\PortailProd\SAPCredentials"
    
    @staticmethod
    def _simple_encrypt(text):
        """Chiffrement simple des données (non sécurisé pour des données très sensibles)"""
        return base64.b64encode(text.encode()).decode()
    
    @staticmethod
    def _simple_decrypt(encrypted_text):
        """Déchiffrement simple des données"""
        try:
            return base64.b64decode(encrypted_text.encode()).decode()
        except:
            return ""
    
    @staticmethod
    def get_credentials() -> Tuple[str, str]:
        """Récupère les identifiants SAP stockés ou demande à l'utilisateur."""
        # Essayer d'abord le registre Windows (sur Windows seulement)
        if platform.system() == "Windows":
            username, password = CredentialsManager._load_credentials_from_registry()
            if username and password:
                return username, password
        
        # Sinon, essayer le fichier JSON
        username, password = CredentialsManager._load_credentials_from_file()
        
        if not username or not password:
            # Si aucun identifiant n'est trouvé, demander à l'utilisateur
            return CredentialsManager._prompt_credentials()
        
        return username, password
    
    @staticmethod
    def _load_credentials_from_file() -> Tuple[str, str]:
        """Charge les identifiants depuis le fichier JSON"""
        if not CredentialsManager.CREDS_FILE.exists():
            return "", ""
        
        try:
            with open(CredentialsManager.CREDS_FILE, "r") as f:
                data = json.load(f)
                username = CredentialsManager._simple_decrypt(data.get("username", ""))
                password = CredentialsManager._simple_decrypt(data.get("password", ""))
                return username, password
        except Exception:
            return "", ""
    
    @staticmethod
    def _load_credentials_from_registry() -> Tuple[str, str]:
        """Charge les identifiants depuis le registre Windows"""
        if platform.system() != "Windows":
            return "", ""
        
        try:
            import winreg
            reg_key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                CredentialsManager.REG_PATH,
                0,
                winreg.KEY_READ
            )
            
            username = winreg.QueryValueEx(reg_key, "Username")[0]
            encoded_password = winreg.QueryValueEx(reg_key, "Password")[0]
            
            # Convertir le format d'encodage VBS vers notre format
            password = ""
            parts = encoded_password.split(".")
            for part in parts:
                if part:  # Ignorer les parties vides
                    try:
                        password += chr(int(part))
                    except ValueError:
                        pass
            
            winreg.CloseKey(reg_key)
            return username, password
        
        except Exception as e:
            print(f"Erreur lors de la lecture du registre: {e}")
            return "", ""
    
    @staticmethod
    def _save_credentials(username, password):
        """Sauvegarde les identifiants dans un fichier et dans le registre Windows si disponible"""
        # Sauvegarder dans le fichier JSON
        CredentialsManager._save_credentials_to_file(username, password)
        
        # Sauvegarder dans le registre Windows si sur Windows
        if platform.system() == "Windows":
            CredentialsManager._save_credentials_to_registry(username, password)
    
    @staticmethod
    def _save_credentials_to_file(username, password):
        """Sauvegarde les identifiants dans un fichier JSON"""
        # Créer le répertoire si nécessaire
        CredentialsManager.CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        
        # Sauvegarder les identifiants chiffrés
        data = {
            "username": CredentialsManager._simple_encrypt(username),
            "password": CredentialsManager._simple_encrypt(password)
        }
        
        with open(CredentialsManager.CREDS_FILE, "w") as f:
            json.dump(data, f)
    
    @staticmethod
    def _save_credentials_to_registry(username, password):
        """Sauvegarde les identifiants dans le registre Windows"""
        if platform.system() != "Windows":
            return
        
        try:
            import winreg
            
            # Créer ou ouvrir la clé
            reg_key = winreg.CreateKey(
                winreg.HKEY_CURRENT_USER,
                CredentialsManager.REG_PATH
            )
            
            # Enregistrer le nom d'utilisateur
            winreg.SetValueEx(reg_key, "Username", 0, winreg.REG_SZ, username)
            
            # Encoder le mot de passe dans le format compatible avec VBS
            encoded_password = ""
            for char in password:
                encoded_password += str(ord(char)) + "."
            
            # Enregistrer le mot de passe encodé
            winreg.SetValueEx(reg_key, "Password", 0, winreg.REG_SZ, encoded_password)
            
            winreg.CloseKey(reg_key)
        except Exception as e:
            print(f"Erreur lors de l'enregistrement dans le registre: {e}")
    
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

        # Stockage des identifiants dans les deux sources
        CredentialsManager._save_credentials(username, password)
        
        return username, password
    
    @staticmethod
    def clear_credentials():
        """Efface les identifiants stockés."""
        # Effacer du fichier
        try:
            if CredentialsManager.CREDS_FILE.exists():
                CredentialsManager.CREDS_FILE.unlink()
        except Exception:
            pass
        
        # Effacer du registre Windows
        if platform.system() == "Windows":
            try:
                import winreg
                winreg.DeleteKey(
                    winreg.HKEY_CURRENT_USER,
                    CredentialsManager.REG_PATH
                )
            except Exception:
                pass