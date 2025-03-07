import json
import re
from pathlib import Path

from PySide6.QtWidgets import (
    QWidget, QLineEdit, QLabel, QPushButton, QVBoxLayout,
    QHBoxLayout, QMessageBox, QTextEdit, QSpacerItem, QSizePolicy
)
from PySide6.QtCore import Qt, QSize, QTimer
from PySide6.QtGui import QFont, QFontMetrics
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
        # Keep original size hint, but allow for expansion
        self.setSizePolicy(QSizePolicy.Minimum,QSizePolicy.Expanding)

        # Main horizontal layout to split left and right sides
        main_layout = QHBoxLayout()
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # Left side container
        left_container = QWidget()
        left_container.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        left_layout = QVBoxLayout(left_container)
        left_layout.setSpacing(20)
        left_layout.setContentsMargins(0, 0, 0, 0)

        # Search input
        search_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Entrez la référence du composant...")
        self.search_input.setMinimumHeight(40)
        self.search_input.textChanged.connect(self.search_component)
        search_layout.addWidget(self.search_input)
        left_layout.addLayout(search_layout)

        # Display areas
        left_layout.addLayout(self.create_display_area("Nom du composant:", "component_name"))
        left_layout.addLayout(self.create_display_area("Référence:", "reference"))
        left_layout.addLayout(self.create_display_area("Description:", "description"))

        # Buttons
        button_layout = QHBoxLayout()
        self.open_file_button = QPushButton("Ouvrir le fichier")
        self.open_file_button.clicked.connect(self.open_file)
        self.reset_list_button = QPushButton("Initialiser la liste")
        self.reset_list_button.clicked.connect(self.reset_list)
        button_layout.addWidget(self.open_file_button)
        button_layout.addWidget(self.reset_list_button)
        left_layout.addLayout(button_layout)

        # Stretch to push everything to the top
        left_layout.addStretch()

        # Right side - Selected components
        self.selected_components_widget = QTextEdit()
        self.selected_components_widget.setReadOnly(True)
        self.selected_components_widget.setMinimumWidth(300)  # Ensure reasonable width for text display

        # Add containers to main layout
        main_layout.addWidget(left_container)
        main_layout.addWidget(self.selected_components_widget)

        self.setLayout(main_layout)

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
        
        # Définir un titre temporaire pour la fenêtre 
        self.setWindowTitle("Recherche de Composants")

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
        QMessageBox.critical(self, "Erreur", message, QMessageBox.Ok)

    def load_json_file(self):
        file_path = Path(resources_dir).parent / 'data_ref.json'
        try:
            with open(file_path, 'r', encoding='cp1252') as jsonfile:
                data = json.load(jsonfile)
                try:
                    row = next(row for row in data if row[0] == "0001")
                    reference = row[1].strip()
                    description = row[2].strip()
                    
                    # Construire le titre avec référence et description
                    title = f"Référence: {reference} - Description: {description}"
                    self.setWindowTitle(title)
                    
                    # Ajuster la taille de la police pour s'assurer que le titre s'affiche entièrement
                    self.adjustTitleFontSize(title)
                except StopIteration:
                    self.setWindowTitle("Clé 0001 non trouvée")

                for row in data:
                    if isinstance(row, list) and len(row) >= 2 and row[0] != "0001":
                        key = row[0].strip()
                        normalized_key = self.normalize_key(key)
                        value = [row[1].strip()]
                        if len(row) > 2:
                            value.append(row[2].strip())
                        if normalized_key:
                            self.data_by_key[normalized_key] = value

        except (json.JSONDecodeError, FileNotFoundError) as e:
            self.show_error(f"Erreur lors du chargement du fichier JSON: {e}")
        except Exception as e:
            self.show_error(f"Erreur inattendue: {e}")
        finally:
            self.update_selected_components_display()

    def adjustTitleFontSize(self, title):
        # Cette fonction est appelée après un court délai pour s'assurer que la fenêtre est construite
        QTimer.singleShot(100, lambda: self.doAdjustTitleFontSize(title))

    def doAdjustTitleFontSize(self, title):
        # S'assurer que le titre s'affiche correctement sans être tronqué
        # Cette implémentation est une solution simple car on ne peut pas directement
        # accéder à la police de la barre de titre de façon portable
        width = self.width()
        if width <= 0:
            # La fenêtre n'est pas encore complètement initialisée, réessayer plus tard
            QTimer.singleShot(100, lambda: self.doAdjustTitleFontSize(title))
            return

        # La longueur estimée du titre vs. la largeur de la fenêtre
        font = QFont("Arial", 10, QFont.Bold)
        metrics = QFontMetrics(font)
        title_width = metrics.horizontalAdvance(title)
        
        if title_width > width * 0.9:  # Si le titre est trop long (90% de la largeur de la fenêtre)
            # Adapter la taille minimale de la fenêtre pour accommoder le titre
            self.setMinimumWidth(title_width + 40)  # 40px de marge

    def search_component(self):
        search_text = self.normalize_key(self.search_input.text())
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
        try:
            with file_path.open("a", encoding='utf-8') as file:
                for component in self.selected_components:
                    file.write(component)
                    file.write("\n" * 5)
                    file.write("=" * 40 + "\n")
            if file_path.exists():
                os.startfile(str(file_path))
        except Exception as e:
            self.show_error(f"Erreur lors de l'écriture du fichier: {e}")

    def add_to_list(self):
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
        try:
            with file_path.open("w", encoding='utf-8') as file:
                file.write("")
        except Exception as e:
            self.show_error(f"Erreur lors de la réinitialisation du fichier: {e}")

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key_Return, Qt.Key_Enter):
            self.add_to_list()
        else:
            super().keyPressEvent(event)

    def update_selected_components_display(self):
        self.selected_components_widget.setPlainText("\n".join(self.selected_components))

def nom_app():
    widget = ComponentSearchWidget()
    widget.show()
    widget.raise_()
    widget.activateWindow()
    return widget
