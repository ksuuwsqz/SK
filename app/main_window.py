"""
Главное окно приложения: боковая навигация и переключение страниц.
"""
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QButtonGroup, QFrame, QHBoxLayout, QLabel, QMainWindow,
    QPushButton, QStackedWidget, QVBoxLayout, QWidget, QMessageBox,
)

from app.pages.check_page import CheckPage
from app.pages.home_page import HomePage
from app.pages.report_page import ReportPage
from app.pages.summary_page import SummaryPage
from app.pages.upload_page import UploadPage
from app.services.dynamics import build_dynamics_text, enrich_dynamics
from app.theme import APP_STYLESHEET, COLOR_SIDEBAR

# Имена разделов навигации
NAV_ITEMS = [
    "Главная",
    "Загрузка отчётов",
    "Проверка данных",
    "Сводка и динамика",
    "Докладная",
]


class MainWindow(QMainWindow):
    """Главное окно приложения."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Отчёты СК РФ по Челябинской области")
        self.resize(1280, 800)
        self.setStyleSheet(APP_STYLESHEET)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QHBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # --- Боковая панель ---
        sidebar = QFrame()
        sidebar.setObjectName("Sidebar")
        sidebar.setFixedWidth(260)
        sidebar.setStyleSheet(f"background-color: {COLOR_SIDEBAR}; border: none;")
        side_layout = QVBoxLayout(sidebar)
        side_layout.setContentsMargins(0, 24, 0, 24)
        side_layout.setSpacing(4)

        logo = QLabel("СК РФ\nпо Челябинской\nобласти")
        logo.setStyleSheet(
            "color: white; font-weight: 700; font-size: 12pt; "
            "padding: 0 12px 18px 12px;"
        )
        logo.setWordWrap(True)
        side_layout.addWidget(logo)

        self.nav_group = QButtonGroup(self)
        self.nav_group.setExclusive(True)
        for idx, name in enumerate(NAV_ITEMS):
            btn = QPushButton(name)
            btn.setObjectName("NavButton")
            btn.setCheckable(True)
            if idx == 0:
                btn.setChecked(True)
            btn.clicked.connect(lambda _checked, i=idx: self._switch_page(i))
            side_layout.addWidget(btn)
            self.nav_group.addButton(btn, idx)
        side_layout.addStretch()

        layout.addWidget(sidebar)

        # --- Стек страниц ---
        self.stack = QStackedWidget()
        self.home_page = HomePage()
        self.upload_page = UploadPage()
        self.check_page = CheckPage()
        self.summary_page = SummaryPage()
        self.report_page = ReportPage()

        for page in (self.home_page, self.upload_page, self.check_page,
                     self.summary_page, self.report_page):
            self.stack.addWidget(page)

        layout.addWidget(self.stack, 1)

        # --- Связи ---
        self.upload_page.process_requested.connect(self._process_data)

        # Статус-бар
        self.statusBar().showMessage("Готово. Ожидание загрузки файлов...")

    # ---------- Переключение страниц ----------
    def _switch_page(self, idx: int) -> None:
        self.stack.setCurrentIndex(idx)

    # ---------- Обработка данных ----------
    def process_files(self, reports: list) -> None:
        """Обработка файлов, переданных из страницы загрузки."""
        if not reports:
            QMessageBox.warning(
                self, "Нет файлов",
                "Сначала выберите файл и нажмите «Добавить файл», чтобы он появился в таблице.",
            )
            return

        try:
            from app.services.excel_loader import load_excel_sheets
            from app.services.validator import validate_report

            all_errors = []
            for r in reports:
                sheets = load_excel_sheets(r.file_path)
                all_errors.extend(validate_report(r, sheets))
        except Exception as exc:
            QMessageBox.critical(
                self, "Ошибка обработки",
                f"Не удалось прочитать файлы:\n{exc}",
            )
            return

        self.check_page.set_errors(all_errors)

        # Передаем файлы для докладной
        period_name = reports[0].period if reports else "—"
        self.report_page.set_context([], period_name, reports)

        self.statusBar().showMessage(
            f"Обработано файлов: {len(reports)}, найдено ошибок: {len(all_errors)}."
        )
        QMessageBox.information(
            self, "Обработка завершена",
            f"Успешно обработано файлов: {len(reports)}\n"
            f"Найдено формальных ошибок: {len(all_errors)}\n\n"
            "Теперь вы можете перейти в раздел «Докладная» "
            "и нажать «Сформировать Word».",
        )