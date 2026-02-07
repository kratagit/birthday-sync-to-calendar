import json
from datetime import datetime
from utils import get_app_data_dir

class DataManager:
    def __init__(self):
        self.data_file_path = get_app_data_dir() / "data.json"
        self.data = []
        self.load_data()
        self.recalculate_ages()

    def load_data(self):
        """ Wczytuje dane z pliku. """
        try:
            with open(self.data_file_path, "r", encoding='utf-8') as file:
                self.data = json.load(file)
        except (FileNotFoundError, json.JSONDecodeError):
            self.data = []
            self.save_data()

    def save_data(self):
        """ Zapisuje dane do pliku. """
        with open(self.data_file_path, "w", encoding='utf-8') as file:
            json.dump(self.data, file, indent=4, ensure_ascii=False)

    def calculate_age(self, birth_date_str):
        """ Oblicza wiek. """
        try:
            birth_date = datetime.strptime(birth_date_str, "%Y-%m-%d")
            today = datetime.today()
            age = today.year - birth_date.year - ((today.month, today.day) < (birth_date.month, birth_date.day))
            return str(age)
        except ValueError:
            return "0"

    def recalculate_ages(self):
        """ Aktualizuje wiek dla wszystkich osób. """
        changed = False
        for person in self.data:
            new_age = self.calculate_age(person["date"])
            if person.get("age") != new_age:
                person["age"] = new_age
                changed = True
        if changed:
            self.save_data()

    def add_person(self, name, date_str):
        """ Dodaje osobę i zwraca True jeśli się udało. """
        age = self.calculate_age(date_str)
        self.data.append({"name": name, "date": date_str, "age": age})
        self.save_data()

    def remove_person(self, index):
        """ Usuwa osobę po indeksie. """
        if 0 <= index < len(self.data):
            del self.data[index]
            self.save_data()

    def get_data(self):
        return self.data
    
    def sort_by_birthday(self):
        today = datetime.today().date()
        def days_until(birth_date_str):
            birth_date = datetime.strptime(birth_date_str, "%Y-%m-%d").date()
            next_birthday = birth_date.replace(year=today.year)
            if next_birthday < today:
                next_birthday = next_birthday.replace(year=today.year + 1)
            return (next_birthday - today).days
        self.data.sort(key=lambda person: days_until(person["date"]))

    def sort_chronologically(self):
        self.data.sort(key=lambda person: person["date"][5:])