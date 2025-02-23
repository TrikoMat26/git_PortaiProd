import json
import re
from pathlib import Path

from PySide6.QtWidgets import (
    QWidget, QLineEdit, QLabel, QPushButton, QVBoxLayout,
    QHBoxLayout, QFileDialog, QMessageBox, QTextEdit
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QIcon
import os
from package.resourcesPath import AppContext

ctx = AppContext.get()
resources_dir = ctx.get_resource('Fichies_config/dummy.txt')

class ComponentSearchWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.data_by_key = {}
        self.selected_components = []
        self.setup_ui()
        self.load_json_file()

    def setup_ui(self):
        self.setWindowTitle("Recherche de Composants")
        self.setFixedSize(800, 400)
        
        # Layout principal
        main_layout = QHBoxLayout()
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # Layout gauche (recherche + boutons)
        left_layout = QVBoxLayout()
        left_layout.setSpacing(20)

        # Zone de recherche
        search_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Entrez la référence du composant...")
        self.search_input.setMinimumHeight(40)
        self.search_input.textChanged.connect(self.search_component)
        
        search_icon = QLabel()
        search_icon.setPixmap(QIcon("search_icon.png").pixmap(20, 20))
        search_layout.addWidget(search_icon)
        search_layout.addWidget(self.search_input)
        left_layout.addLayout(search_layout)

        # Zones d'affichage : nom, référence, description
        left_layout.addLayout(self.create_display_area("Nom du composant:", "component_name"))
        left_layout.addLayout(self.create_display_area("Référence:", "reference"))
        left_layout.addLayout(self.create_display_area("Description:", "description"))

        # Boutons
        button_layout = QHBoxLayout()
        self.open_file_button = QPushButton("Ouvrir le fichier")
        self.open_file_button.clicked.connect(self.open_file)
        self.reset_list_button = QPushButton("Initialiser la liste")
        self.reset_list_button.clicked.connect(self.reset_list)
        button_layout.addWidget(self.open_file_button)
        button_layout.addWidget(self.reset_list_button)
        left_layout.addLayout(button_layout)

        # Ajout du layout gauche
        main_layout.addLayout(left_layout)

        # Widget (QTextEdit) pour les composants sélectionnés
        self.selected_components_widget = QTextEdit()
        self.selected_components_widget.setReadOnly(True)
        main_layout.addWidget(self.selected_components_widget)

        self.setLayout(main_layout)

        # Style
        self.setStyleSheet("""
            QWidget {
                background-color: #ffffff;
                border-radius: 10px;
            }
            QLineEdit {
                border: 2px solid #007bff;
                border-radius: 8px;
                padding: 10px;
                font-size: 16px;
                background-color: #f9f9f9;
            }
            QLabel {
                font-size: 14px;
                color: #333;
                font-weight: bold;
            }
            QLineEdit:disabled {
                background-color: #e9ecef;
                color: #6c757d;
            }
            QPushButton {
                background-color: #007bff;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
            QTextEdit {
                border: 2px solid #007bff;
                border-radius: 8px;
                padding: 10px;
                font-size: 14px;
                background-color: #f9f9f9;
            }
        """)

    def create_display_area(self, label_text, object_name):
        layout = QVBoxLayout()

        label = QLabel(label_text)
        label.setFont(QFont("Arial", 10))
        layout.addWidget(label)

        text_field = QLineEdit()
        text_field.setObjectName(object_name)
        text_field.setReadOnly(True)
        text_field.setMinimumHeight(40)
        layout.addWidget(text_field)

        return layout

    def normalize_key(self, key):
        key = key.strip().lower()
        return re.sub(r'(\D)0+(\d+)', r'\1\2', key)

    def show_error(self, message):
        QMessageBox.critical(
            self,
            "Erreur",
            message,
            QMessageBox.Ok
        )

    def load_json_file(self):
        file_path = Path(resources_dir).parent / 'data_ref.json'
        try:
            # Utilisation de l'encodage CP1252 (Windows-1252) pour lire correctement les caractères spéciaux
            with open(file_path, 'r', encoding='cp1252') as jsonfile:
                rows = json.load(jsonfile)
                for row in rows:
                    if isinstance(row, list) and len(row) >= 2:
                        key = row[0].strip()
                        normalized_key = self.normalize_key(key)
                        value = [row[1].strip()]
                        if len(row) > 2:
                            value.append(row[2].strip())

                        if normalized_key:
                            self.data_by_key[normalized_key] = value
        except json.JSONDecodeError:
            self.show_error("Le fichier sélectionné n'est pas un fichier JSON valide.")
        except Exception as e:
            self.show_error(f"Erreur lors de la lecture du fichier : {str(e)}")

    def search_component(self):
        search_text = self.normalize_key(self.search_input.text())

        # Réinitialiser les champs
        self.findChild(QLineEdit, "component_name").clear()
        self.findChild(QLineEdit, "reference").clear()
        self.findChild(QLineEdit, "description").clear()

        if search_text in self.data_by_key:
            component_data = self.data_by_key[search_text]
            self.findChild(QLineEdit, "component_name").setText(search_text)
            self.findChild(QLineEdit, "reference").setText(component_data[0])
            if len(component_data) > 1:
                self.findChild(QLineEdit, "description").setText(component_data[1])

    def open_file(self):
        file_path = Path(resources_dir).parent / "commande_composants.txt"
        with file_path.open("a", encoding='utf-8') as file:
            for component in self.selected_components:
                file.write(component)
                file.write("\n" * 5)
                file.write("=" * 40 + "\n")
        if file_path.exists():
            os.startfile(str(file_path))

    def add_to_list(self):
        # Récupération des champs et conversion du nom en majuscules
        component_name = self.findChild(QLineEdit, "component_name").text().upper()
        reference = self.findChild(QLineEdit, "reference").text()
        description = self.findChild(QLineEdit, "description").text()

        if component_name and reference:
            component_info = (
                f"Nom du composant: {component_name}\n"
                f"Référence: {reference}\n"
                f"Description: {description}\n\n"
            )
            self.selected_components.append(component_info)
            self.update_selected_components_display()
            self.search_input.clear()
            self.findChild(QLineEdit, "component_name").clear()
            self.findChild(QLineEdit, "reference").clear()
            self.findChild(QLineEdit, "description").clear()
            self.search_input.setFocus()
        else:
            self.show_error("Veuillez sélectionner un composant valide.")

    def reset_list(self):
        self.selected_components.clear()
        self.update_selected_components_display()
        self.search_input.setFocus()

        file_path = Path(resources_dir).parent / "commande_composants.txt"
        with file_path.open("w", encoding='utf-8') as file:
            file.write("")

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key_Return, Qt.Key_Enter):
            self.add_to_list()
        else:
            super().keyPressEvent(event)

    def update_selected_components_display(self):
        self.selected_components_widget.setPlainText("\n".join(self.selected_components))


def nom_app():
    """
    Cette fonction instancie simplement le QWidget, le rend visible,
    et force l'interface à être au premier plan.
    """
    widget = ComponentSearchWidget()
    widget.show()
    widget.raise_()         # Place la fenêtre au-dessus des autres
    widget.activateWindow() # Active la fenêtre pour qu'elle reçoive le focus
    return widget
