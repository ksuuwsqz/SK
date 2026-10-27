"""
Точка входа приложения.
Запускает главное окно PyQt6.
"""
import sys

from PyQt6.QtWidgets import QApplication

from app.main_window import MainWindow


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName("Отчёты СК РФ по Челябинской области")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()