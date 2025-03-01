import json
import re
from pathlib import Path

from PySide6.QtWidgets import (
    QWidget, QLineEdit, QLabel, QPushButton, QVBoxLayout,
    QHBoxLayout, QMessageBox, QTextEdit, QSpacerItem, QSizePolicy
)
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QFont
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
        # Keep original size hint, but allow for expansion
        self.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Expanding)

        # Main horizontal layout to split left and right sides
        main_layout = QHBoxLayout()
        main_layout.setSpacing(24)  # Increased spacing
        main_layout.setContentsMargins(24, 24, 24, 24)  # Increased margins

        # Left side container
        left_container = QWidget()
        left_container.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        left_layout = QVBoxLayout(left_container)
        left_layout.setSpacing(20)
        left_layout.setContentsMargins(0, 0, 0, 0)

        # Shadow effect for left container
        left_container.setObjectName("leftPanel")

        # Top labels for "0001" data with modernized design
        top_label_widget = QWidget()
        top_label_widget.setObjectName("headerPanel")
        top_label_layout = QHBoxLayout(top_label_widget)
        top_label_layout.setContentsMargins(16, 16, 16, 16)
        self.key0001_ref_label = QLabel()
        self.key0001_ref_label.setObjectName("headerLabel")
        self.key0001_desc_label = QLabel()
        self.key0001_desc_label.setObjectName("headerLabel")
        top_label_layout.addWidget(self.key0001_ref_label)
        top_label_layout.addWidget(self.key0001_desc_label)
        top_label_layout.addStretch(1)
        left_layout.addWidget(top_label_widget)

        # Search input
        search_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Entrez la référence du composant...")
        self.search_input.setMinimumHeight(48)  # Increased height
        self.search_input.setObjectName("searchInput")
        self.search_input.textChanged.connect(self.search_component)
        search_layout.addWidget(self.search_input)
        left_layout.addLayout(search_layout)

        # Display areas
        left_layout.addLayout(self.create_display_area("Nom du composant:", "component_name"))
        left_layout.addLayout(self.create_display_area("Référence:", "reference"))
        left_layout.addLayout(self.create_display_area("Description:", "description"))

        # Buttons
        button_layout = QHBoxLayout()
        button_layout.setSpacing(16)  # Increased spacing between buttons
        
        self.open_file_button = QPushButton("Ouvrir le fichier")
        self.open_file_button.setObjectName("primaryButton")
        self.open_file_button.setMinimumHeight(40)  # Consistent button height
        self.open_file_button.clicked.connect(self.open_file)
        
        self.reset_list_button = QPushButton("Initialiser la liste")
        self.reset_list_button.setObjectName("secondaryButton")
        self.reset_list_button.setMinimumHeight(40)  # Consistent button height
        self.reset_list_button.clicked.connect(self.reset_list)
        
        button_layout.addWidget(self.open_file_button)
        button_layout.addWidget(self.reset_list_button)
        left_layout.addLayout(button_layout)

        # Stretch to push everything to the top
        left_layout.addStretch()

        # Right side - Selected components
        self.selected_components_widget = QTextEdit()
        self.selected_components_widget.setReadOnly(True)
        self.selected_components_widget.setMinimumWidth(320)  # Slightly increased width
        self.selected_components_widget.setObjectName("selectedComponentsPanel")

        # Add containers to main layout
        main_layout.addWidget(left_container, 3)  # Give more space to left side
        main_layout.addWidget(self.selected_components_widget, 2)  # Proportional allocation

        self.setLayout(main_layout)

        # Modern stylesheet with updated color scheme and effects
        self.setStyleSheet("""
            QWidget {
                background-color: #f0f2f5;
                color: #202124;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            
            #leftPanel {
                background-color: #ffffff;
                border-radius: 12px;
                border: 1px solid #d1d5db;
                box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            }
            
            #selectedComponentsPanel {
                background-color: #ffffff;
                border-radius: 12px;
                border: 1px solid #d1d5db;
                background-color: #f8f9fa;
                box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            }
            
            #headerPanel {
                background-color: #e6effd;
                border-radius: 10px;
                margin-bottom: 16px;
                border-left: 4px solid #0d6efd;
            }
            
            #headerLabel {
                font-weight: 600;
                font-size: 15px;
                color: #0d6efd;
                text-shadow: 0 1px 0 rgba(255, 255, 255, 0.5);
            }
            
            #searchInput {
                border: 2px solid #0d6efd;
                border-radius: 8px;
                padding: 12px 16px;
                font-size: 16px;
                background-color: #ffffff;
                color: #202124;
                transition: all 0.2s ease-in-out;
            }
            
            #searchInput:focus {
                border-color: #0a58ca;
                box-shadow: 0 0 0 4px rgba(13, 110, 253, 0.25);
                outline: none;
            }
            
            QLineEdit {
                border: 2px solid #d1d5db;
                border-radius: 8px;
                padding: 10px;
                font-size: 14px;
                background-color: #f8fafc;
                color: #202124;
            }
            
            QLineEdit:disabled, QLineEdit[readOnly="true"] {
                background-color: #f1f3f5;
                color: #495057;
                border-color: #ced4da;
            }
            
            QLabel {
                font-size: 14px;
                color: #343a40;
                font-weight: 600;
                margin-bottom: 4px;
                padding-left: 2px;
            }
            
            #primaryButton {
                background-color: #0d6efd;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 12px;
                font-size: 14px;
                font-weight: 600;
                transition: all 0.2s ease;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
            }
            
            #primaryButton:hover {
                background-color: #0b5ed7;
                box-shadow: 0 4px 8px rgba(0, 0, 0, 0.15);
            }
            
            #secondaryButton {
                background-color: #6c757d;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 12px;
                font-size: 14px;
                font-weight: 600;
                transition: all 0.2s ease;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
            }
            
            #secondaryButton:hover {
                background-color: #5c636a;
                box-shadow: 0 4px 8px rgba(0, 0, 0, 0.15);
            }
            
            QTextEdit {
                border: 1px solid #d1d5db;
                border-radius: 10px;
                padding: 12px;
                font-size: 14px;
                background-color: #ffffff;
                color: #202124;
                line-height: 1.6;
            }
            
            /* Enhanced field group styling */
            QLineEdit[objectName="component_name"] {
                border-left: 4px solid #20c997;
            }
            
            QLineEdit[objectName="reference"] {
                border-left: 4px solid #fd7e14;
            }
            
            QLineEdit[objectName="description"] {
                border-left: 4px solid #6610f2;
            }
        """)

    def create_display_area(self, label_text, object_name):
        layout = QVBoxLayout()
        layout.setSpacing(6)
        
        # Crée un widget conteneur pour l'étiquette pour permettre un style distinctif
        label_container = QWidget()
        label_container.setObjectName(f"{object_name}_label_container")
        label_container.setMaximumHeight(30)
        label_container_layout = QHBoxLayout(label_container)
        label_container_layout.setContentsMargins(0, 0, 0, 0)
        
        label = QLabel(label_text)
        label.setObjectName(f"{object_name}_label")
        label.setFont(QFont("Segoe UI", 10, QFont.Weight.DemiBold))
        label_container_layout.addWidget(label)
        label_container_layout.addStretch()
        
        layout.addWidget(label_container)
        
        text_field = QLineEdit()
        text_field.setObjectName(object_name)
        text_field.setReadOnly(True)
        text_field.setMinimumHeight(42)
        layout.addWidget(text_field)
        
        # Espace vertical entre les groupes de champs
        layout.addSpacing(12)
        
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
                    self.key0001_ref_label.setText(f"Référence: {row[1].strip()}")
                    self.key0001_desc_label.setText(f"Description: {row[2].strip()}")
                except StopIteration:
                    self.key0001_ref_label.setText("Clé 0001 non trouvée")
                    self.key0001_desc_label.setText("")

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
