import sys
import os
from pathlib import Path

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