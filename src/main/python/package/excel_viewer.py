# gui/excel_viewer.py

from pathlib import Path
import json
import re
import pandas as pd

from PySide6.QtWidgets import (
    QMainWindow, QTableWidget, QTableWidgetItem, QVBoxLayout,
    QWidget, QPushButton, QLineEdit, QHeaderView, QHBoxLayout, QMessageBox,
    QFrame, QSizePolicy, QComboBox, QLabel, QDialog, QGridLayout, QApplication, QScrollArea
    
)
from PySide6.QtCore import Qt, QTimer, Signal
from PySide6.QtGui import QCursor, QKeyEvent, QWheelEvent

from package.logic import BusinessLogic
from package.resourcesPath import AppContext



# Sous-classe de QScrollArea pour un défilement "tout ou rien"
class CustomScrollArea(QScrollArea):
    def wheelEvent(self, event: QWheelEvent):
        # Si la roulette monte, on va tout en haut ; sinon, tout en bas.
        if event.angleDelta().y() > 0:
            self.verticalScrollBar().setValue(self.verticalScrollBar().minimum())
        else:
            self.verticalScrollBar().setValue(self.verticalScrollBar().maximum())
        event.accept()

    def keyPressEvent(self, event: QKeyEvent):
        if event.key() == Qt.Key_Up:
            self.verticalScrollBar().setValue(self.verticalScrollBar().minimum())
            event.accept()
        elif event.key() == Qt.Key_Down:
            self.verticalScrollBar().setValue(self.verticalScrollBar().maximum())
            event.accept()
        else:
            super().keyPressEvent(event)


def load_action_buttons_from_json(filename="action_buttons.json"):
    ctx = AppContext.get()
    resources_dir = ctx.get_resource('Fichies_config/dummy.txt')
    json_path = Path(resources_dir).parent / filename
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data.get("buttons", [])


class ContextMenuLineEdit(QLineEdit):
    """
    LineEdit personnalisé qui émet un signal doubleClicked au double-clic.
    """
    doubleClicked = Signal(QLineEdit)

    def mouseDoubleClickEvent(self, event):
        self.doubleClicked.emit(self)
        super().mouseDoubleClickEvent(event)


class FilterContainer(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFrameStyle(QFrame.NoFrame)
        layout = QHBoxLayout(self)
        layout.setSpacing(0)
        layout.setContentsMargins(0, 0, 0, 0)
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)


class FilterBox(QLineEdit):
    """
    Champ de filtre avec un combo en double-clic (inchangé).
    Permet d'afficher une liste de valeurs disponibles pour la colonne concernée.
    """
    def __init__(self, column_name, main_window=None):
        super().__init__()
        self.setPlaceholderText(f"Filtrer {column_name}")
        self.setFixedWidth(100)
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.column_name = column_name
        self.main_window = main_window
        self.combo = QComboBox(self)
        self.combo.setVisible(False)

        self.combo.setStyleSheet("""
        QComboBox {
            background-color: #1e1e1e; color: white;
            border: 1px solid #555; border-radius: 4px; padding: 5px;
        }
        QComboBox QAbstractItemView {
            background-color: #1e1e1e; color: white; selection-background-color: #404040;
        }
        """)
        self.combo.activated.connect(self.on_combo_activated)

    def mouseDoubleClickEvent(self, event):
        """
        Affiche une combo listant les valeurs uniques de la colonne filtrée.
        """
        self.clearFocus()
        if self.main_window:
            logic = self.main_window.logic
            if logic and not logic.filtered_df.empty:
                unique_values = logic.filtered_df[self.column_name].dropna().astype(str).unique()
                self.combo.clear()
                self.combo.addItems(sorted(unique_values))
                max_width = max(self.fontMetrics().boundingRect(val).width() for val in unique_values) + 30
                self.combo.setFixedWidth(max(max_width, self.width()))
                self.combo.move(0, self.height())
                self.combo.setVisible(True)
                self.combo.showPopup()

    def on_combo_activated(self, index):
        selected_value = self.combo.currentText()
        self.setText(selected_value)
        self.combo.setVisible(False)
        if self.main_window:
            self.main_window.apply_filters_delayed()
        self.setCursorPosition(0)

    def focusOutEvent(self, event):
        super().focusOutEvent(event)
        QTimer.singleShot(100, self.hide_combo)

    def hide_combo(self):
        if self.combo.isVisible() and not self.combo.view().isVisible():
            self.combo.setVisible(False)


class ActionSelectionDialog(QDialog):
    """
    Fenêtre de sélection d'actions, non-modale (restera ouverte tant
    qu'on ne la ferme pas manuellement).
    On lit la liste des actions depuis un JSON pour un aspect évolutif.
    """
    def __init__(self, action_names, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Sélection d'action")
        self.setWindowModality(Qt.NonModal)
        self.action_names = action_names
        self.parent_window = parent

        # Layout principal vertical
        main_layout = QVBoxLayout(self)

        # Bouton Fermer en premier (rouge)
        close_button = QPushButton("Fermer cette fenêtre")
        close_button.setStyleSheet("""
            QPushButton {
                background-color: red; 
                color: white; 
                font-weight: bold;
                border-radius: 5px;
                padding: 5px 15px;
            }
            QPushButton:hover {
                background-color: #aa0000;
            }
        """)
        close_button.clicked.connect(self.close)
        main_layout.addWidget(close_button)

        # --- Utilisation de CustomScrollArea pour les boutons ---
        scroll_area = CustomScrollArea()
        scroll_area.setWidgetResizable(True)  # s'adapte à la taille du contenu

        # Le conteneur qui va réellement héberger les boutons
        scroll_container = QWidget()
        scroll_layout = QVBoxLayout(scroll_container)
        scroll_layout.setSpacing(10)  # Espacement vertical entre boutons

        # Ajout des boutons dans scroll_container
        for action_name in self.action_names:
            btn = QPushButton(action_name)
            btn.clicked.connect(
                lambda checked, name=action_name: self.on_action_button_clicked(name)
            )
            scroll_layout.addWidget(btn)

        # Dire à la QScrollArea que scroll_container est son widget
        scroll_area.setWidget(scroll_container)

        # On limite la hauteur pour n'afficher que ~8 boutons
        # Supposons ~40 px de hauteur par bouton (selon votre style).
        # Ajustez selon vos besoins (ex. 8 * 45 = 360).
        scroll_area.setFixedHeight(9 * 50)

        # On ajoute la zone défilante dans le layout principal
        main_layout.addWidget(scroll_area)

        self.adjustSize()

    def on_action_button_clicked(self, action_name):
        """
        Appelé lorsqu'on clique sur un des boutons.
        On appelle la logique métier parent, MAIS on ne ferme pas la fenêtre.
        """
        if self.parent_window:
            self.parent_window.handle_menu_action(action_name)
        # Ne pas fermer la fenêtre => reste ouverte


class ExcelViewerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        # Ajouter un drapeau de premier affichage
        self.first_show = True 
        self.setWindowTitle("Excel Data Viewer")

        # -- Logique métier --
        self.logic = BusinessLogic()

        # -- Liste de boutons chargée depuis JSON --
        # (Adaptation du chemin si nécessaire)
        self.action_buttons = load_action_buttons_from_json("action_buttons.json")

        # On stocke l'instance du dialogue pour qu'il ne soit pas détruit
        self.dialog_actions = None
        self.active_input = None  # mémorise le champ sur lequel on a cliqué

        # -- Paramètres d'UI --
        self.vertical_header_width = 0
        self.row_height = 30
        self.num_rows_display = 15

        self.filter_boxes = []
        self.timer = QTimer()
        self.timer.setSingleShot(True)
        self.timer.timeout.connect(self.apply_filters_delayed)

        self.init_stylesheet()
        self.init_ui()
        # Initialisation différée après l'affichage de la fenêtre
        self.loading_label = None

    def showEvent(self, event):
        """Déclenche le chargement des données après l'affichage initial"""
        super().showEvent(event)
        if self.first_show:
            self.show_loading_message()
            QTimer.singleShot(0, self.load_data_into_ui)
            self.first_show = False

    def show_loading_message(self):
        """Affiche le message de chargement"""
        self.loading_label = QLabel("Chargement du fichier...", self)
        self.loading_label.setAlignment(Qt.AlignCenter)
        self.loading_label.setStyleSheet("""
            QLabel {
                font-size: 16px;
                color: #007bff;
                font-weight: bold;
                background-color: rgba(255, 255, 255, 200);
                padding: 20px;
                border-radius: 10px;
            }
        """)
        self.loading_label.resize(self.size())
        self.loading_label.show()

    def hide_loading_message(self):
        """Masque le message de chargement"""
        if self.loading_label:
            self.loading_label.hide()

 
    def init_stylesheet(self):
        self.setStyleSheet("""
        QMainWindow { background-color: #f0f0f0; }
        QPushButton {
            background-color: #007bff; color: white; font-size: 14px;
            font-weight: bold; border-radius: 5px; padding: 10px 20px; border: none;
        }
        QPushButton:hover { background-color: #0056b3; }
        QPushButton:pressed { background-color: #003d82; }
        QTableWidget {
            border: 1px solid #ddd; gridline-color: #eee;
            background-color: white; alternate-background-color: #f9f9f9;
            color: #333; font-size: 13px;
        }
        QHeaderView::section {
            background-color: #eee; color: #333; padding: 8px;
            font-size: 12px; font-weight: bold;
            border-bottom: 1px solid #ddd; border-right: 1px solid #ddd;
        }
        QLineEdit {
            border: 1px solid #ccc; border-radius: 4px; padding: 8px;
            font-size: 13px; background-color: white; color: #333;
        }
        QLineEdit:focus {
            border-color: #007bff; box-shadow: 0 0 5px rgba(0, 123, 255, 0.5);
        }
        QLabel { color: #666; font-size: 12px; font-weight: bold; margin-bottom: 5px; }
        QMenu { background-color: white; border: 1px solid #ddd; }
        QMenu::item { padding: 5px 20px; }
        QMenu::item:selected { background-color: #007bff; color: white; }
        """)

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        self.main_layout = QVBoxLayout(central_widget)
        self.main_layout.setSpacing(15)
        self.main_layout.setContentsMargins(20, 20, 20, 20)

        # --- Barre de saisie OF / Référence + Bouton RAZ ---
        input_wrapper = QWidget()
        input_layout = QHBoxLayout(input_wrapper)
        input_layout.setContentsMargins(0, 0, 0, 0)
        input_layout.setSpacing(20)

        # Groupe OF
        of_group = QWidget()
        of_layout = QVBoxLayout(of_group)
        of_label = QLabel("NUMÉRO OF")
        self.of_input = ContextMenuLineEdit()
        self.of_input.setFixedHeight(30)
        # On connecte la fonction open_action_buttons_dialog
        self.of_input.doubleClicked.connect(lambda: self.open_action_buttons_dialog(self.of_input))
        of_layout.addWidget(of_label)
        of_layout.addWidget(self.of_input)

        # Groupe Référence
        ref_group = QWidget()
        ref_layout = QVBoxLayout(ref_group)
        ref_label = QLabel("RÉFÉRENCE")
        self.reference_input = ContextMenuLineEdit()
        self.reference_input.setFixedHeight(30)
        # On connecte la même fonction
        self.reference_input.doubleClicked.connect(lambda: self.open_action_buttons_dialog(self.reference_input))
        ref_layout.addWidget(ref_label)
        ref_layout.addWidget(self.reference_input)

        # Bouton RAZ
        self.reset_button = QPushButton("RAZ")
        self.reset_button.setFixedHeight(30)
        self.reset_button.clicked.connect(self.reset_filters)

        # Disposition horizontale
        input_layout.addWidget(of_group, 50)
        input_layout.addWidget(ref_group, 50)
        input_layout.addWidget(self.reset_button, 20)

        self.main_layout.addWidget(input_wrapper)

        # --- Zone des filtres (ligne du haut) ---
        filter_wrapper = QWidget()
        filter_wrapper_layout = QHBoxLayout(filter_wrapper)
        self.filter_container = FilterContainer()
        filter_wrapper_layout.addWidget(self.filter_container)
        filter_wrapper_layout.addStretch()
        self.main_layout.addWidget(filter_wrapper)

        # --- Tableau ---
        self.table_widget = QTableWidget()
        self.table_widget.setAlternatingRowColors(True)
        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.table_widget.verticalHeader().setDefaultSectionSize(self.row_height)
        self.table_widget.cellDoubleClicked.connect(self.on_cell_double_clicked)
        self.table_widget.verticalHeader().sectionDoubleClicked.connect(self.on_row_header_double_clicked)
        self.main_layout.addWidget(self.table_widget)

    # ----------------------------------------------------------------
    #  Chargement des données et mise à jour du tableau
    # ----------------------------------------------------------------

    def load_data_into_ui(self):
        try:
            # Charger les données
            self.logic.load_data()
            
            # Afficher le message avec le fichier utilisé (nom seulement)
            fichier_source = Path(self.logic.current_df.attrs.get('source_file', 'inconnu')).name
            msg_box = QMessageBox(self)
            msg_box.setWindowTitle("Fichier chargé")
            msg_box.setText(f"Fichier utilisé :\n{fichier_source}")
            msg_box.setIcon(QMessageBox.Information)
            msg_box.setWindowModality(Qt.NonModal)
            msg_box.show()
            # Mise à jour de l'interface
            self.update_table_structure()
            self.populate_table()
            self.adjust_window_size()
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur lors de la lecture : {str(e)}")
        finally:
            self.hide_loading_message()

    def update_table_structure(self):
        df = self.logic.current_df
        self.table_widget.setRowCount(len(df))
        self.table_widget.setColumnCount(len(df.columns))
        self.table_widget.setHorizontalHeaderLabels(df.columns)

        # Largeurs de colonnes
        column_widths = {
            "OF": 70,
            "Référence": 130,
            "Designation": 150,
            "Client": 100,
            "Opération": 80,
            "Désignation OP": 110,
            "Qté. prévu": 80,
            "Qté. livrée": 80,
            "DateFinOF": 110
        }
        for col_idx in range(self.table_widget.columnCount()):
            col_name = df.columns[col_idx]
            if col_name in column_widths:
                self.table_widget.setColumnWidth(col_idx, column_widths[col_name])

        self.add_filter_boxes()

        # Calcul de la largeur totale du tableau
        table_width = self.table_widget.verticalHeader().width()
        for col in range(self.table_widget.columnCount()):
            table_width += self.table_widget.columnWidth(col)
        total_width = table_width + 40  # marges
        self.resize(total_width, 600)

        self.vertical_header_width = self.table_widget.verticalHeader().width()
        self.table_widget.verticalHeader().setFixedWidth(self.vertical_header_width)
        self.adjust_filter_position()
        self.adjust_filter_boxes_width()

    def add_filter_boxes(self):
        for box in self.filter_boxes:
            box.deleteLater()
        self.filter_boxes.clear()
        self.filter_container.layout().takeAt(0)

        for col_name in self.logic.current_df.columns:
            filter_box = FilterBox(col_name, main_window=self)
            filter_box.textChanged.connect(self.start_filtering_timer)
            self.filter_container.layout().addWidget(filter_box)
            self.filter_boxes.append(filter_box)

    def populate_table(self):
        df = self.logic.filtered_df
        self.table_widget.setRowCount(len(df))

        for row_index, (_, row_data) in enumerate(df.iterrows()):
            for col_index, cell_data in enumerate(row_data):
                item = QTableWidgetItem(str(cell_data))
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)  # Rendre non éditable

                column_name = df.columns[col_index]
                if column_name == "DateFinOF":
                    item.setTextAlignment(Qt.AlignCenter)
                else:
                    item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                    item.setData(Qt.TextWrapAnywhere, True)

                self.table_widget.setItem(row_index, col_index, item)

        self.table_widget.resizeRowsToContents()

    # ----------------------------------------------------------------
    #  Filtres
    # ----------------------------------------------------------------

    def start_filtering_timer(self):
        self.timer.start(300)  # 300ms

    def apply_filters_delayed(self):
        filter_texts = [box.text() for box in self.filter_boxes]
        try:
            self.logic.filter_data(filter_texts)
            self.populate_table()
        except Exception as e:
            QMessageBox.warning(self, "Erreur", f"Erreur de filtrage : {e}")

    def reset_filters(self):
        for filter_box in self.filter_boxes:
            filter_box.blockSignals(True)
            filter_box.clear()
            filter_box.blockSignals(False)
        self.of_input.clear()
        self.reference_input.clear()

        self.logic.reset_filters()
        self.populate_table()

    # ----------------------------------------------------------------
    #  Sélection dans le tableau
    # ----------------------------------------------------------------

    def on_row_header_double_clicked(self, logical_index):
        """
        Double clic sur l'entête de ligne => remplit champ OF/REF.
        """
        try:
            df = self.logic.filtered_df
            if not df.empty and all(col in df.columns for col in ["OF", "Référence"]):
                of_val = str(df.iloc[logical_index]["OF"])
                ref_val = str(df.iloc[logical_index]["Référence"])
                self.of_input.setText(of_val)
                self.reference_input.setText(ref_val)
        except Exception as e:
            QMessageBox.warning(self, "Erreur", f"Erreur lors de la copie : {str(e)}")

    def on_cell_double_clicked(self, row, column):
        """
        Double-clic sur une cellule => applique le texte au filtre correspondant.
        """
        try:
            cell_value = self.table_widget.item(row, column).text()
            if column < len(self.filter_boxes):
                filter_box = self.filter_boxes[column]
                filter_box.setText(cell_value)
                filter_box.setCursorPosition(0)
                self.apply_filters_delayed()
        except Exception as e:
            QMessageBox.warning(self, "Erreur", f"Erreur lors de la sélection : {str(e)}")

    # ----------------------------------------------------------------
    #  Nouvelle fenêtre non-modale de sélection d'actions
    # ----------------------------------------------------------------

    # gui/excel_viewer.py (modifié)

    def open_action_buttons_dialog(self, source_input):
        """
        Ouvre (ou ré-affiche) la boîte de dialogue contenant tous les boutons d'action,
        positionnée à droite de la dernière colonne de données + 50px plus haut.
        """
        from PySide6.QtGui import QGuiApplication
        from PySide6.QtCore import QPoint

        self.active_input = source_input

        if not self.dialog_actions:
            self.dialog_actions = ActionSelectionDialog(self.action_buttons, parent=self)

        self.dialog_actions.show()

        # Calcul position horizontale
        total_columns_width = sum(
            self.table_widget.columnWidth(i) 
            for i in range(self.table_widget.columnCount())
        )
        x_in_table = self.vertical_header_width + total_columns_width
        
        # Ajustement vertical : hauteur header - 50px
        y_in_table = self.table_widget.horizontalHeader().height() - 50  # Modification ici

        # Conversion en coordonnées globales
        global_pos = self.table_widget.mapToGlobal(QPoint(x_in_table, y_in_table))

        # Gestion multi-écran
        screen = QGuiApplication.screenAt(global_pos) or QGuiApplication.primaryScreen()
        if screen:
            screen_geo = screen.availableGeometry()
            dialog_size = self.dialog_actions.size()

            # Correction horizontale
            if (global_pos.x() + dialog_size.width()) > screen_geo.right():
                global_pos.setX(screen_geo.right() - dialog_size.width())
            
            # Correction verticale (avec limite supérieure)
            min_y = screen_geo.top()
            new_y = max(min_y, global_pos.y())  # Bloque en haut de l'écran
            if (new_y + dialog_size.height()) > screen_geo.bottom():
                new_y = screen_geo.bottom() - dialog_size.height()
            global_pos.setY(new_y)

        self.dialog_actions.move(global_pos)
        self.dialog_actions.raise_()
        self.dialog_actions.activateWindow()




    def handle_menu_action(self, action_name):
        """
        Appelle la logique pour exécuter l’action correspondante.
        """
        if action_name in ["Note infos", "COOIS sur OF"]:
            input_value = self.of_input.text().strip()
    
        else:
            input_value = self.active_input.text().strip()

        if not input_value:
            # Si rien saisi, on tente de prendre le presse-papiers
            clipboard = QApplication.clipboard()
            input_value = clipboard.text().strip()

        if not input_value:
            QMessageBox.critical(self, "Erreur", "Aucune valeur disponible !")
            return

        if len(input_value) > 40:
            QMessageBox.critical(self, "Erreur", "La valeur ne doit pas excéder 40 caractères.")
            return
        
        # On délègue la logique métier
        try:
            self.logic.execute_action(action_name, input_value)
        except Exception as e:
            QMessageBox.critical(self, "Erreur", str(e))

    # ----------------------------------------------------------------
    #  Ajustements layout
    # ----------------------------------------------------------------

    def adjust_filter_boxes_width(self):
        column_widths = {
            "OF": 70,
            "Référence": 130,
            "Designation": 150,
            "Client": 100,
            "Opération": 80,
            "Désignation OP": 110,
            "Qté. prévu": 80,
            "Qté. livrée": 80,
            "DateFinOF": 110
        }
        for i, filter_box in enumerate(self.filter_boxes):
            col_name = self.logic.current_df.columns[i]
            if col_name in column_widths:
                filter_box.setFixedWidth(column_widths[col_name])

    def adjust_filter_position(self):
        self.filter_container.setContentsMargins(self.vertical_header_width - 10, 0, 0, 0)

    def adjust_window_size(self):
        total_height = (self.row_height * self.num_rows_display)
        total_height += self.table_widget.horizontalHeader().height()
        total_height += 200  # marge pour la zone filtres, etc.
        self.setFixedHeight(total_height)


if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    window = ExcelViewerApp()
    window.show()
    sys.exit(app.exec())
