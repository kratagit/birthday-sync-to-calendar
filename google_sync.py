import os
from datetime import datetime, timedelta
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from PyQt5.QtWidgets import QMessageBox, QProgressDialog, QApplication
from PyQt5.QtCore import Qt
from utils import resource_path, get_app_data_dir

SCOPES = ['https://www.googleapis.com/auth/calendar']

class GoogleCalendarSync:
    def __init__(self, parent_window):
        self.parent = parent_window

    def export_events(self, data_list):
        # 1. Potwierdzenie użytkownika
        reply = QMessageBox.question(
            self.parent, 
            "Potwierdzenie", 
            "Czy na pewno chcesz wyeksportować dane do Google Calendar?", 
            QMessageBox.Yes | QMessageBox.No, 
            QMessageBox.No
        )
        if reply == QMessageBox.No:
            return

        # 2. Pasek postępu
        progress = QProgressDialog("Przygotowanie do eksportu...", "Anuluj", 0, len(data_list) + 3, self.parent)
        progress.setWindowTitle("Eksportowanie")
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)
        progress.setValue(0)

        creds = None
        try:
            # 3. Autoryzacja (ścieżki do plików)
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

            # 4. Sprawdzanie/Tworzenie kalendarza
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

            # 5. Pobranie istniejących wydarzeń (żeby nie dublować)
            existing_events = service.events().list(calendarId=calendar_id, singleEvents=True).execute().get('items', [])
            
            progress.setLabelText("Eksportowanie urodzin...")
            progress.setValue(3)
            QApplication.processEvents()

            # 6. Pętla dodawania wydarzeń
            for i, person in enumerate(data_list):
                progress.setValue(3 + i)
                if progress.wasCanceled():
                    break
                
                name = person["name"]
                birth_date = person["date"] # np. "2000-05-20"
                year = birth_date.split('-')[0]
                
                # Obliczanie daty końca (start + 1 dzień) dla wydarzenia całodniowego
                dt_obj = datetime.strptime(birth_date, '%Y-%m-%d')
                end_date_obj = dt_obj + timedelta(days=1)
                end_date = end_date_obj.strftime('%Y-%m-%d')

                # Tytuł
                summary = f'Urodziny: {name} {year}'

                # Sprawdzenie czy wydarzenie już istnieje
                event_exists = any(
                    event.get('summary') == summary and
                    (event.get('start', {}).get('date') == birth_date)
                    for event in existing_events
                )

                if event_exists:
                    continue

                # Definicja wydarzenia
                event_body = {
                    'summary': summary,
                    'start': {'date': birth_date}, # Cały dzień
                    'end': {'date': end_date},     # Wymagane dla całodniowych (następny dzień)
                    'recurrence': ['RRULE:FREQ=YEARLY'],
                    'colorId': '6', # Kolor pomarańczowy/dyniowy
                    'reminders': {
                        'useDefault': True 
                    }
                }
                service.events().insert(calendarId=calendar_id, body=event_body).execute()

            # 7. Zakończenie
            progress.setValue(len(data_list) + 3)
            QMessageBox.information(self.parent, "Sukces", "Urodziny zostały pomyślnie zsynchronizowane z kalendarzem 'Birthdays'!")

        except FileNotFoundError as e:
            QMessageBox.critical(self.parent, "Błąd pliku", f"Nie znaleziono wymaganego pliku: {e}")
        except Exception as e:
            if progress: progress.close()
            QMessageBox.critical(self.parent, "Błąd", f"Wystąpił błąd podczas eksportu:\n{str(e)}")