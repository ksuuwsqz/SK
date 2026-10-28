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
        self.upload_page = UploadPage(self)
        self.check_page = CheckPage()
        self.summary_page = SummaryPage()
        self.report_page = ReportPage()

        for page in (self.home_page, self.upload_page, self.check_page,
                     self.summary_page, self.report_page):
            self.stack.addWidget(page)

        layout.addWidget(self.stack, 1)

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
            from app.services.excel_to_word import Report
            from app.data.models import SummaryRow
            from app.services.dynamics import enrich_dynamics

            all_errors = []
            summary_rows = []

            for r in reports:
                # 1. Проверка на формальные ошибки
                sheets = load_excel_sheets(r.file_path)
                all_errors.extend(validate_report(r, sheets))

                # 2. Формирование сводки
                try:
                    current_report = Report(r.file_path)
                    current_vals = current_report.values()

                    # Пытаемся найти предыдущий период для сравнения
                    # Ищем файл с таким же именем, но на 1 месяц назад (например, 2602 -> 2601)
                    import re
                    import os

                    # Получаем имя текущего файла без пути
                    filename = os.path.basename(r.file_path)
                    match = re.search(r"_(\d{2})(\d{2})\.xls", filename, re.IGNORECASE)

                    prev_vals = None
                    if match:
                        year = int(match.group(1))
                        month = int(match.group(2))

                        # Вычисляем прошлый месяц
                        if month > 1:
                            prev_month = month - 1
                            prev_year = year
                        else:
                            prev_month = 12
                            prev_year = year - 1

                        # Ищем файл в той же папке
                        prev_filename = f"{filename[:match.start(1)]}{prev_year:02d}{prev_month:02d}.XLS"
                        prev_path = os.path.join(os.path.dirname(r.file_path), prev_filename)

                        if os.path.exists(prev_path):
                            prev_report = Report(prev_path)
                            prev_vals = prev_report.values()

                    # Формируем строки для сводки
                    for k, _, _, title in [
                        ("messages", "", "", "Поступило сообщений"),
                        ("cancelled", "", "", "Отменено постановлений о возбуждении УД"),
                        ("in_progress", "", "", "Находилось в производстве УД"),
                        ("to_court", "", "", "Направлено уголовных дел в суд"),
                        ("terminated", "", "", "Прекращено уголовных дел"),
                    ]:
                        cur_val = current_vals[k]["value"]
                        prev_val = prev_vals[k]["value"] if prev_vals else None

                        summary_rows.append(SummaryRow(
                            district=r.district,
                            indicator=title,
                            current_value=float(cur_val) if cur_val is not None else None,
                            previous_value=float(prev_val) if prev_val is not None else None
                        ))
                except Exception as e:
                    print("Ошибка при подготовке сводки:", e)

            # Обновляем страницы
            self.check_page.set_errors(all_errors)

            # Передаем данные в сводку
            summary_rows = enrich_dynamics(summary_rows)
            self.summary_page.set_rows(summary_rows)

            # Передаем файлы для докладной
            period_name = reports[0].period if reports else "—"
            self.report_page.set_context([], period_name, reports)

            self.statusBar().showMessage(
                f"Обработано файлов: {len(reports)}, найдено ошибок: {len(all_errors)}.",
            )
            QMessageBox.information(
                self, "Обработка завершена",
                f"Успешно обработано файлов: {len(reports)}\n"
                f"Найдено формальных ошибок: {len(all_errors)}\n\n"
                "Данные переданы в разделы «Сводка» и «Докладная».",
            )
        except Exception as exc:
            QMessageBox.critical(
                self, "Ошибка обработки",
                f"Не удалось прочитать файлы:\n{exc}",
            )