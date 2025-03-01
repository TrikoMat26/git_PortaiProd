import json
import re
from pathlib import Path

from PySide6.QtWidgets import (
    QWidget, QLineEdit, QLabel, QPushButton, QVBoxLayout,
    QHBoxLayout, QMessageBox, QTextEdit, QSpacerItem, QSizePolicy,
    QFrame
)
from PySide6.QtCore import Qt, QSize
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
        # Keep original size hint, but allow for expansion
        self.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Expanding)
        self.setMinimumSize(900, 700)  # Set minimum size for better layout

        # Main horizontal layout to split left and right sides
        main_layout = QHBoxLayout()
        main_layout.setSpacing(30)  # Increased spacing
        main_layout.setContentsMargins(24, 24, 24, 24)  # Increased margins

        # Left side container with card-style design
        left_container = QFrame()
        left_container.setFrameShape(QFrame.StyledPanel)
        left_container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        left_container.setObjectName("leftPanel")
        left_layout = QVBoxLayout(left_container)
        left_layout.setSpacing(20)
        left_layout.setContentsMargins(20, 20, 20, 20)

        # Application title/header with modern design
        app_title = QLabel("Gestionnaire de Composants")
        app_title.setObjectName("appTitle")
        app_title.setAlignment(Qt.AlignCenter)
        left_layout.addWidget(app_title)

        # Separator line
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        separator.setObjectName("separator")
        left_layout.addWidget(separator)
        
        # Top labels for "0001" data with card-style design
        top_label_widget = QFrame()
        top_label_widget.setObjectName("headerPanel")
        top_label_layout = QVBoxLayout(top_label_widget)
        top_label_layout.setContentsMargins(16, 16, 16, 16)
        
        header_title = QLabel("Informations de Base")
        header_title.setObjectName("sectionTitle")
        top_label_layout.addWidget(header_title)
        
        self.key0001_ref_label = QLabel()
        self.key0001_ref_label.setObjectName("infoLabel")
        self.key0001_desc_label = QLabel()
        self.key0001_desc_label.setObjectName("infoLabel")
        self.key0001_desc_label.setWordWrap(True)
        
        top_label_layout.addWidget(self.key0001_ref_label)
        top_label_layout.addWidget(self.key0001_desc_label)
        left_layout.addWidget(top_label_widget)

        # Search section with improved visual design
        search_section = QFrame()
        search_section.setObjectName("searchSection")
        search_layout = QVBoxLayout(search_section)
        search_layout.setContentsMargins(16, 16, 16, 16)
        
        search_title = QLabel("Recherche de Composant")
        search_title.setObjectName("sectionTitle")
        search_layout.addWidget(search_title)
        
        search_input_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Entrez la référence du composant...")
        self.search_input.setMinimumHeight(48)
        self.search_input.setObjectName("searchInput")
        self.search_input.textChanged.connect(self.search_component)
        
        search_button = QPushButton()
        search_button.setObjectName("searchButton")
        search_button.setMinimumHeight(48)
        search_button.setMaximumWidth(48)
        search_button.clicked.connect(self.search_component)
        
        search_input_layout.addWidget(self.search_input)
        search_input_layout.addWidget(search_button)
        search_layout.addLayout(search_input_layout)
        left_layout.addWidget(search_section)

        # Component details section with card-style design
        details_section = QFrame()
        details_section.setObjectName("detailsSection")
        details_layout = QVBoxLayout(details_section)
        details_layout.setContentsMargins(16, 16, 16, 16)
        
        details_title = QLabel("Détails du Composant")
        details_title.setObjectName("sectionTitle")
        details_layout.addWidget(details_title)
        
        # Display areas with improved styling
        details_layout.addLayout(self.create_display_area("Nom du composant:", "component_name"))
        details_layout.addLayout(self.create_display_area("Référence:", "reference"))
        details_layout.addLayout(self.create_display_area("Description:", "description"))
        
        # Add to list button
        add_button = QPushButton("Ajouter à la liste")
        add_button.setObjectName("addButton")
        add_button.setMinimumHeight(44)
        add_button.clicked.connect(self.add_to_list)
        details_layout.addWidget(add_button)
        
        left_layout.addWidget(details_section)

        # Buttons with improved layout and styling
        actions_section = QFrame()
        actions_section.setObjectName("actionsSection")
        button_layout = QHBoxLayout(actions_section)
        button_layout.setContentsMargins(16, 16, 16, 16)
        button_layout.setSpacing(16)
        
        self.open_file_button = QPushButton("Ouvrir le fichier")
        self.open_file_button.setObjectName("primaryButton")
        self.open_file_button.setMinimumHeight(44)
        self.open_file_button.clicked.connect(self.open_file)
        
        self.reset_list_button = QPushButton("Réinitialiser la liste")
        self.reset_list_button.setObjectName("secondaryButton")
        self.reset_list_button.setMinimumHeight(44)
        self.reset_list_button.clicked.connect(self.reset_list)
        
        button_layout.addWidget(self.open_file_button)
        button_layout.addWidget(self.reset_list_button)
        left_layout.addWidget(actions_section)

        # Stretch to push everything to the top
        left_layout.addStretch()

        # Right side - Selected components with card-style design
        right_container = QFrame()
        right_container.setFrameShape(QFrame.StyledPanel)
        right_container.setObjectName("rightPanel")
        right_layout = QVBoxLayout(right_container)
        right_layout.setContentsMargins(20, 20, 20, 20)
        right_layout.setSpacing(16)
        
        selected_title = QLabel("Composants Sélectionnés")
        selected_title.setObjectName("sectionTitle")
        selected_title.setAlignment(Qt.AlignCenter)
        right_layout.addWidget(selected_title)
        
        self.selected_components_widget = QTextEdit()
        self.selected_components_widget.setReadOnly(True)
        self.selected_components_widget.setObjectName("selectedComponentsPanel")
        self.selected_components_widget.setMinimumWidth(350)
        right_layout.addWidget(self.selected_components_widget)
        
        # Selected components count
        self.components_count_label = QLabel("0 composants sélectionnés")
        self.components_count_label.setObjectName("countLabel")
        self.components_count_label.setAlignment(Qt.AlignCenter)
        right_layout.addWidget(self.components_count_label)

        # Add containers to main layout
        main_layout.addWidget(left_container, 3)
        main_layout.addWidget(right_container, 2)

        self.setLayout(main_layout)

        # Modern professional stylesheet
        self.setStyleSheet("""
            QWidget {
                background-color: #f5f7fb;
                color: #2d3748;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 14px;
            }
            
            #appTitle {
                font-size: 22px;
                font-weight: bold;
                color: #2c3e50;
                padding: 10px 0;
                margin-bottom: 5px;
            }
            
            #sectionTitle {
                font-size: 16px;
                font-weight: bold;
                color: #3b82f6;
                margin-bottom: 10px;
                border-bottom: 1px solid #e2e8f0;
                padding-bottom: 8px;
            }
            
            #separator {
                background-color: #e2e8f0;
                height: 1px;
                margin: 5px 0 15px 0;
            }
            
            #leftPanel, #rightPanel, #headerPanel, #searchSection, #detailsSection, #actionsSection {
                background-color: white;
                border-radius: 12px;
                box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05), 0 1px 3px rgba(0, 0, 0, 0.1);
                margin-bottom: 15px;
            }
            
            #headerPanel {
                background-color: #ebf4ff;
                border-left: 4px solid #3b82f6;
            }
            
            #infoLabel {
                font-size: 14px;
                color: #4a5568;
                margin: 5px 0;
                padding: 3px 0;
            }
            
            #searchInput {
                border: 2px solid #cbd5e0;
                border-radius: 8px;
                padding: 10px 15px;
                font-size: 15px;
                background-color: white;
                transition: all 0.2s ease;
            }
            
            #searchInput:focus {
                border-color: #3b82f6;
                box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.3);
                outline: none;
            }
            
            #searchButton {
                background-color: #3b82f6;
                border-radius: 8px;
                border: none;
                color: white;
                font-size: 16px;
                font-weight: bold;
                qproperty-icon: url(search_icon.png);
                qproperty-iconSize: 20px 20px;
            }
            
            #searchButton:hover {
                background-color: #2563eb;
            }
            
            QLineEdit {
                border: 1px solid #cbd5e0;
                border-radius: 6px;
                padding: 10px;
                background-color: white;
                font-size: 14px;
            }
            
            QLineEdit:disabled, QLineEdit[readOnly="true"] {
                background-color: #f8fafc;
                color: #64748b;
                border-color: #e2e8f0;
            }
            
            QLabel {
                color: #4b5563;
                font-weight: 500;
            }
            
            #primaryButton, #addButton {
                background-color: #3b82f6;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 10px 20px;
                font-weight: 600;
                font-size: 14px;
                transition: all 0.2s ease;
            }
            
            #primaryButton:hover, #addButton:hover {
                background-color: #2563eb;
                box-shadow: 0 4px 6px rgba(37, 99, 235, 0.2);
            }
            
            #secondaryButton {
                background-color: #64748b;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 10px 20px;
                font-weight: 600;
                font-size: 14px;
                transition: all 0.2s ease;
            }
            
            #secondaryButton:hover {
                background-color: #475569;
                box-shadow: 0 4px 6px rgba(71, 85, 105, 0.2);
            }
            
            #selectedComponentsPanel {
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                padding: 15px;
                background-color: #f8fafc;
                font-size: 14px;
                line-height: 1.6;
            }
            
            #countLabel {
                color: #64748b;
                font-size: 13px;
                font-style: italic;
            }
            
            /* Field styling with visual indicators */
            QLineEdit[objectName="component_name"] {
                border-left: 4px solid #10b981;
            }
            
            QLineEdit[objectName="reference"] {
                border-left: 4px solid #f59e0b;
            }
            
            QLineEdit[objectName="description"] {
                border-left: 4px solid #8b5cf6;
            }
            
            /* Label styling for field groups */
            QLabel[objectName="component_name_label"] {
                color: #10b981;
                font-weight: 600;
            }
            
            QLabel[objectName="reference_label"] {
                color: #f59e0b;
                font-weight: 600;
            }
            
            QLabel[objectName="description_label"] {
                color: #8b5cf6;
                font-weight: 600;
            }
        """)

    def create_display_area(self, label_text, object_name):
        layout = QVBoxLayout()
        layout.setSpacing(6)
        
        # Label with improved styling
        label = QLabel(label_text)
        label.setObjectName(f"{object_name}_label")
        label.setFont(QFont("Segoe UI", 10, QFont.Weight.DemiBold))
        layout.addWidget(label)
        
        # Text field with improved styling
        text_field = QLineEdit()
        text_field.setObjectName(object_name)
        text_field.setReadOnly(True)
        text_field.setMinimumHeight(40)
        layout.addWidget(text_field)
        
        # Spacing between field groups
        layout.addSpacing(10)
        
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
        count = len(self.selected_components)
        self.components_count_label.setText(f"{count} composant{'s' if count != 1 else ''} sélectionné{'s' if count != 1 else ''}")

def nom_app():
    widget = ComponentSearchWidget()
    widget.show()
    widget.raise_()
    widget.activateWindow()
    return widget
