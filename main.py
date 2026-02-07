import sys
import os
from PyQt5.QtWidgets import QApplication, QMessageBox
from gui import BirthdayApp

if __name__ == "__main__":
    # Blokada wielokrotnego uruchomienia (mutex na pliku)
    lock_file_path = os.path.join(os.getcwd(), "birthday_app.lock")
    
    if os.path.exists(lock_file_path):
        app_temp = QApplication(sys.argv)
        QMessageBox.warning(None, "Aplikacja działa", "Aplikacja jest już uruchomiona.")
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
            try:
                os.remove(lock_file_path)
            except:
                pass