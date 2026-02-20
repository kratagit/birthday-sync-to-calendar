import sys
from PyQt5 import uic
from PyQt5.QtWidgets import (
    QMainWindow, QMessageBox, QTableWidgetItem, 
    QInputDialog, QStyledItemDelegate, QHeaderView,
    QStyleOptionComboBox, QStyle
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
from datetime import datetime

# Importy naszych modułów
from utils import resource_path
from data_manager import DataManager
from google_sync import GoogleCalendarSync

class CenteredItemDelegate(QStyledItemDelegate):
    def initStyleOption(self, option, index):
        super().initStyleOption(option, index)
        option.displayAlignment = Qt.AlignCenter

class BirthdayApp(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Ładowanie UI
        try:
            uic.loadUi(resource_path("mainwindow.ui"), self)
        except Exception as e:
            QMessageBox.critical(self, "Błąd", f"Nie można załadować pliku interfejsu: {e}")
            sys.exit(1)

        try:
            self.setWindowIcon(QIcon(resource_path("app_icon.ico")))
        except Exception:
            pass

        # Inicjalizacja menedżerów
        self.data_manager = DataManager()
        self.google_sync = GoogleCalendarSync(self)

        # Konfiguracja UI
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.horizontalLayout_labels.setStretch(0, 1)
        self.horizontalLayout_labels.setStretch(1, 1)
        self.horizontalLayout_labels.setStretch(2, 1)
        self.horizontalLayout_combos.setStretch(0, 1)
        self.horizontalLayout_combos.setStretch(1, 1)
        self.horizontalLayout_combos.setStretch(2, 1)
        self.setup_combos()

        # Podpięcie przycisków
        self.add_button.clicked.connect(self.add_person_action)
        self.delete_button.clicked.connect(self.delete_person_action)
        self.sort_by_nearest_button.clicked.connect(self.sort_nearest_action)
        self.sort_chronologically_button.clicked.connect(self.sort_chrono_action)
        self.export_button.clicked.connect(self.export_action)

        self.update_table()

    def setup_combos(self):
        current_date = datetime.now()
        
        # Pomocnicza funkcja do konfiguracji combo
        def configure_combo(combo, items, default, validator_func):
            combo.addItems(items)
            combo.setCurrentText(str(default))
            combo.lineEdit().setAlignment(Qt.AlignCenter)
            combo.setItemDelegate(CenteredItemDelegate(combo))
            combo.currentTextChanged.connect(validator_func)

        self.configure_date_validators(current_date)

    def balance_combo_text_with_dropdown(self, combo):
        line_edit = combo.lineEdit()
        if line_edit is None:
            return

        option = QStyleOptionComboBox()
        combo.initStyleOption(option)
        arrow_rect = combo.style().subControlRect(
            QStyle.CC_ComboBox,
            option,
            QStyle.SC_ComboBoxArrow,
            combo
        )

        left_margin = max(0, arrow_rect.width())
        line_edit.setTextMargins(left_margin, 0, 0, 0)

    def configure_date_validators(self, current_date):
        # Dzień
        self.day_combo.addItems([str(d) for d in range(1, 32)])
        self.day_combo.setCurrentText(str(current_date.day))
        self.day_combo.lineEdit().setAlignment(Qt.AlignCenter)
        self.day_combo.setItemDelegate(CenteredItemDelegate(self.day_combo))
        self.balance_combo_text_with_dropdown(self.day_combo)
        self.day_combo.currentTextChanged.connect(
            lambda: self.validate_input(self.day_combo, range(1, 32), current_date.day)
        )
        # Miesiąc
        self.month_combo.addItems([str(m) for m in range(1, 13)])
        self.month_combo.setCurrentText(str(current_date.month))
        self.month_combo.lineEdit().setAlignment(Qt.AlignCenter)
        self.month_combo.setItemDelegate(CenteredItemDelegate(self.month_combo))
        self.balance_combo_text_with_dropdown(self.month_combo)
        self.month_combo.currentTextChanged.connect(
            lambda: self.validate_input(self.month_combo, range(1, 13), current_date.month)
        )
        # Rok
        self.year_combo.addItems([str(y) for y in range(1900, current_date.year + 1)])
        self.year_combo.setCurrentText(str(current_date.year))
        self.year_combo.lineEdit().setAlignment(Qt.AlignCenter)
        self.year_combo.setItemDelegate(CenteredItemDelegate(self.year_combo))
        self.balance_combo_text_with_dropdown(self.year_combo)
        self.year_combo.currentTextChanged.connect(
            lambda: self.validate_input(self.year_combo, range(1900, current_date.year + 1), current_date.year, min_length=4)
        )

    def validate_input(self, combo_box, valid_range, default_value, min_length=1):
        current_text = combo_box.currentText()
        if not current_text or len(current_text) < min_length: return
        try:
            value = int(current_text)
            if value not in valid_range: raise ValueError
        except (ValueError, TypeError):
            # Cicha walidacja lub reset
            pass

    def update_table(self):
        data = self.data_manager.get_data()
        self.table.setRowCount(len(data))
        for row, person in enumerate(data):
            self.table.setItem(row, 0, QTableWidgetItem(person["name"]))
            self.table.setItem(row, 1, QTableWidgetItem(person["date"]))
            self.table.setItem(row, 2, QTableWidgetItem(person["age"]))

    # --- Akcje (Eventy) ---

    def add_person_action(self):
        name = self.name_input.text().strip()
        day = self.day_combo.currentText()
        month = self.month_combo.currentText()
        year = self.year_combo.currentText()
        
        if name and day and month and year:
            date_str = f"{year}-{int(month):02d}-{int(day):02d}"
            try:
                datetime.strptime(date_str, "%Y-%m-%d") # Walidacja daty
                self.data_manager.add_person(name, date_str)
                self.update_table()
                self.name_input.clear()
            except ValueError:
                QMessageBox.warning(self, "Błąd", "Nieprawidłowa data.")
        else:
            QMessageBox.warning(self, "Błąd", "Uzupełnij wszystkie pola.")

    def delete_person_action(self):
        data = self.data_manager.get_data()
        if not data:
            QMessageBox.information(self, "Info", "Lista jest pusta.")
            return
        
        options = [f"{i + 1}: {p['name']}" for i, p in enumerate(data)]
        item, ok = QInputDialog.getItem(self, "Usuń", "Wybierz osobę:", options, editable=False)
        if ok and item:
            index = options.index(item)
            self.data_manager.remove_person(index)
            self.update_table()

    def sort_nearest_action(self):
        self.data_manager.sort_by_birthday()
        self.update_table()

    def sort_chrono_action(self):
        self.data_manager.sort_chronologically()
        self.update_table()

    def export_action(self):
        self.google_sync.export_events(self.data_manager.get_data())