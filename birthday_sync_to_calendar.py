import sys
import json
import os
from pathlib import Path
from PyQt5 import uic
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QMessageBox, QTableWidgetItem, 
    QInputDialog, QStyledItemDelegate, QProgressDialog, QHeaderView
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
from datetime import datetime
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# Zakres uprawnień dla Google Calendar API
SCOPES = ['https://www.googleapis.com/auth/calendar']

# --- Funkcje pomocnicze ---
def resource_path(relative_path):
    """ Zwraca poprawną ścieżkę do zasobów (dla trybu deweloperskiego i .exe). """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_app_data_dir():
    """ Zwraca ścieżkę do folderu danych aplikacji w Dokumentach użytkownika. """
    documents_path = Path.home() / "Documents" / "BirthdayApp"
    documents_path.mkdir(parents=True, exist_ok=True)
    return documents_path

class CenteredItemDelegate(QStyledItemDelegate):
    """ Delegat do wyśrodkowywania tekstu w QComboBox. """
    def initStyleOption(self, option, index):
        super().initStyleOption(option, index)
        option.displayAlignment = Qt.AlignCenter

# --- Główna klasa aplikacji ---
class BirthdayApp(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Ładowanie interfejsu z pliku .ui
        try:
            uic.loadUi(resource_path("mainwindow.ui"), self)
        except Exception as e:
            QMessageBox.critical(self, "Błąd", f"Nie można załadować pliku interfejsu mainwindow.ui: {e}")
            sys.exit(1)

        # Ustawienia okna (ikona itp.)
        try:
            self.setWindowIcon(QIcon(resource_path("app_icon.ico")))
        except Exception:
            pass # Ignoruj brak ikony

        # Konfiguracja tabeli (rozciąganie kolumn - reszta ustawiona w .ui)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        # Inicjalizacja Comboboxów (wypełnienie danymi)
        self.setup_combos()

        # Podpięcie przycisków do funkcji
        self.add_button.clicked.connect(self.add_person)
        self.delete_button.clicked.connect(self.delete_person)
        self.sort_by_nearest_button.clicked.connect(self.sort_by_birthday)
        self.sort_chronologically_button.clicked.connect(self.sort_chronologically)
        self.export_button.clicked.connect(self.export_to_google_calendar)

        # Dane
        self.data_file_path = get_app_data_dir() / "data.json"
        self.load_data()

        # Przelicz wiek i odśwież tabelę
        self.recalculate_ages()
        self.update_table()

    def setup_combos(self):
        """ Wypełnia Comboboxy dynamicznymi danymi (dni, miesiące, lata). """
        current_date = datetime.now()

        # Dzień
        self.day_combo.addItems([str(d) for d in range(1, 32)])
        self.day_combo.setCurrentText(str(current_date.day))
        self.day_combo.lineEdit().setAlignment(Qt.AlignCenter)
        self.day_combo.setItemDelegate(CenteredItemDelegate(self.day_combo))
        self.day_combo.currentTextChanged.connect(
             lambda: self.validate_input(self.day_combo, range(1, 32), current_date.day)
        )

        # Miesiąc
        self.month_combo.addItems([str(m) for m in range(1, 13)])
        self.month_combo.setCurrentText(str(current_date.month))
        self.month_combo.lineEdit().setAlignment(Qt.AlignCenter)
        self.month_combo.setItemDelegate(CenteredItemDelegate(self.month_combo))
        self.month_combo.currentTextChanged.connect(
            lambda: self.validate_input(self.month_combo, range(1, 13), current_date.month)
        )

        # Rok
        self.year_combo.addItems([str(y) for y in range(1900, current_date.year + 1)])
        self.year_combo.setCurrentText(str(current_date.year))
        self.year_combo.lineEdit().setAlignment(Qt.AlignCenter)
        self.year_combo.setItemDelegate(CenteredItemDelegate(self.year_combo))
        self.year_combo.currentTextChanged.connect(
            lambda: self.validate_input(self.year_combo, range(1900, current_date.year + 1), current_date.year, min_length=4)
        )

    # --- Metody obsługi danych (bez zmian logicznych) ---
    def load_data(self):
        try:
            with open(self.data_file_path, "r", encoding='utf-8') as file:
                self.data = json.load(file)
        except (FileNotFoundError, json.JSONDecodeError):
            self.data = []
            self.save_data()

    def save_data(self):
        with open(self.data_file_path, "w", encoding='utf-8') as file:
            json.dump(self.data, file, indent=4, ensure_ascii=False)

    def calculate_age(self, birth_date_str):
        birth_date = datetime.strptime(birth_date_str, "%Y-%m-%d")
        today = datetime.today()
        age = today.year - birth_date.year - ((today.month, today.day) < (birth_date.month, birth_date.day))
        return str(age)

    def recalculate_ages(self):
        changed = False
        for person in self.data:
            new_age = self.calculate_age(person["date"])
            if person.get("age") != new_age:
                person["age"] = new_age
                changed = True
        if changed:
            self.save_data()

    def update_table(self):
        self.table.setRowCount(len(self.data))
        for row, person in enumerate(self.data):
            self.table.setItem(row, 0, QTableWidgetItem(person["name"]))
            self.table.setItem(row, 1, QTableWidgetItem(person["date"]))
            self.table.setItem(row, 2, QTableWidgetItem(person["age"]))

    # --- Metody akcji ---
    def add_person(self):
        name = self.name_input.text().strip()
        day = self.day_combo.currentText()
        month = self.month_combo.currentText()
        year = self.year_combo.currentText()
        if name and day and month and year:
            date_str = f"{year}-{int(month):02d}-{int(day):02d}"
            try:
                datetime.strptime(date_str, "%Y-%m-%d")
                age = self.calculate_age(date_str)
                self.data.append({"name": name, "date": date_str, "age": age})
                self.save_data()
                self.update_table()
                self.name_input.clear()
            except ValueError:
                QMessageBox.warning(self, "Błąd", "Nieprawidłowa data (np. 31 luty).")
        else:
            QMessageBox.warning(self, "Błąd", "Wszystkie pola muszą być uzupełnione!")

    def delete_person(self):
        if not self.data:
            QMessageBox.information(self, "Informacja", "Lista jest pusta.")
            return
        options = [f"{i + 1}: {p['name']}" for i, p in enumerate(self.data)]
        item, ok = QInputDialog.getItem(self, "Usuń osobę", "Wybierz osobę do usunięcia:", options, editable=False)
        if ok and item:
            index_to_delete = options.index(item)
            del self.data[index_to_delete]
            self.save_data()
            self.update_table()

    def sort_by_birthday(self):
        today = datetime.today().date()
        def days_until(birth_date_str):
            birth_date = datetime.strptime(birth_date_str, "%Y-%m-%d").date()
            next_birthday = birth_date.replace(year=today.year)
            if next_birthday < today:
                next_birthday = next_birthday.replace(year=today.year + 1)
            return (next_birthday - today).days
        self.data.sort(key=lambda person: days_until(person["date"]))
        self.update_table()

    def sort_chronologically(self):
        self.data.sort(key=lambda person: person["date"][5:])
        self.update_table()

    def validate_input(self, combo_box, valid_range, default_value, min_length=1):
        current_text = combo_box.currentText()
        if not current_text: return
        if len(current_text) < min_length: return
        try:
            value = int(current_text)
            if value not in valid_range: raise ValueError
        except (ValueError, TypeError):
            QMessageBox.warning(self, "Błąd", f"Wartość '{current_text}' jest nieprawidłowa.")
            combo_box.setCurrentText(str(default_value))

    # --- Metoda eksportu do Google Calendar ---
    def export_to_google_calendar(self):
        reply = QMessageBox.question(self, "Potwierdzenie", "Czy na pewno chcesz wyeksportować dane do Google Calendar?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.No: return

        progress = QProgressDialog("Przygotowanie do eksportu...", "Anuluj", 0, len(self.data) + 3, self)
        progress.setWindowTitle("Eksportowanie")
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)
        progress.setValue(0)

        creds = None
        try:
            token_path = get_app_data_dir() / "token.json"
            credentials_path = resource_path("credentials.json")

            progress.setLabelText("Autoryzacja Google...")
            progress.setValue(1)
            QApplication.processEvents()

            if os.path.exists(token_path):
                creds = Credentials.from_authorized_user_file(token_path, SCOPES)
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    if not os.path.exists(credentials_path):
                        raise FileNotFoundError("Brak pliku credentials.json!")
                    flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
                    creds = flow.run_local_server(port=0)
                with open(token_path, 'w') as token_file:
                    token_file.write(creds.to_json())

            service = build('calendar', 'v3', credentials=creds)

            progress.setLabelText("Sprawdzanie kalendarza 'Birthdays'...")
            progress.setValue(2)
            QApplication.processEvents()

            calendar_id = None
            calendar_list = service.calendarList().list().execute().get('items', [])
            for calendar in calendar_list:
                if calendar['summary'] == 'Birthdays':
                    calendar_id = calendar['id']
                    break
            if not calendar_id:
                new_calendar = {'summary': 'Birthdays', 'timeZone': 'Europe/Warsaw'}
                created_calendar = service.calendars().insert(body=new_calendar).execute()
                calendar_id = created_calendar['id']

            calendar_body = {'defaultReminders': [{'method': 'popup', 'minutes': 12 * 60}]}
            service.calendars().patch(calendarId=calendar_id, body=calendar_body).execute()
            
            existing_events = service.events().list(calendarId=calendar_id, singleEvents=True).execute().get('items', [])
            
            progress.setLabelText("Eksportowanie urodzin...")
            progress.setValue(3)
            QApplication.processEvents()

            for i, person in enumerate(self.data):
                progress.setValue(3 + i)
                if progress.wasCanceled(): break
                
                name = person["name"]
                birth_date = person["date"]
                year = birth_date.split('-')[0]
                summary = f'Urodziny: {name} ({year})'

                event_exists = any(
                    event.get('summary') == summary and
                    (event.get('start', {}).get('date') == birth_date or event.get('start', {}).get('dateTime', '').startswith(birth_date))
                    for event in existing_events
                )

                if event_exists: continue

                event_body = {
                    'summary': summary,
                    'start': {'date': birth_date},
                    'end': {'date': birth_date},
                    'recurrence': ['RRULE:FREQ=YEARLY'],
                    'colorId': '6',
                }
                service.events().insert(calendarId=calendar_id, body=event_body).execute()

            progress.setValue(len(self.data) + 3)
            QMessageBox.information(self, "Sukces", "Urodziny zostały pomyślnie zsynchronizowane z kalendarzem 'Birthdays'!")

        except FileNotFoundError as e:
            QMessageBox.critical(self, "Błąd pliku", f"Nie znaleziono wymaganego pliku: {e}")
        except Exception as e:
            if progress: progress.close()
            QMessageBox.critical(self, "Błąd", f"Wystąpił błąd podczas eksportu:\n{str(e)}")

# --- Uruchomienie aplikacji ---
if __name__ == "__main__":
    lock_file_path = os.path.join(os.getcwd(), "birthday_app.lock")
    if os.path.exists(lock_file_path):
        QMessageBox.warning(None, "Aplikacja już działa", "Jedna instancja aplikacji jest już uruchomiona.")
        sys.exit()
    try:
        with open(lock_file_path, "w") as f:
            f.write(str(os.getpid()))
        
        app = QApplication(sys.argv)
        window = BirthdayApp()
        window.show()
        sys.exit(app.exec_())
    finally:
        if os.path.exists(lock_file_path):
            os.remove(lock_file_path)