"""
Страница загрузки Excel-отчётов и привязки к району и периоду.
"""
from datetime import datetime

from app.data.districts import get_all_districts, get_district_name
from PyQt6.QtCore import pyqtSignal
from PyQt6.QtWidgets import (
    QComboBox, QFileDialog, QHBoxLayout, QLabel, QMessageBox,
    QPushButton, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget,
)

from app.data.models import UploadedReport
from app.services.excel_loader import make_uploaded_report
from app.theme import COLOR_MUTED, COLOR_SIDEBAR


class UploadPage(QWidget):
    """Страница загрузки файлов."""

    # Сигнал: пользователь нажал «Запустить обработку»
    process_requested = pyqtSignal()

    def __init__(self, main_window, parent=None):
        super().__init__(parent)
        self.main_window = main_window

        self._selected_file: str | None = None
        self._reports: list[UploadedReport] = []
        self.reports_list = []

        root = QVBoxLayout(self)
        root.setContentsMargins(24, 24, 24, 24)
        root.setSpacing(12)

        title = QLabel("Загрузка отчётов")
        title.setObjectName("PageTitle")
        root.addWidget(title)

        hint = QLabel(
            "Выберите Excel-файл (.xlsx / .xls), укажите район и период, "
            "затем нажмите «Добавить файл»."
        )
        hint.setStyleSheet(f"color: {COLOR_MUTED};")
        hint.setWordWrap(True)
        root.addWidget(hint)

        # Панель управления
        controls = QHBoxLayout()
        controls.setSpacing(8)

        self.file_label = QLabel("Файл не выбран")
        self.file_label.setStyleSheet(f"color: {COLOR_MUTED}; padding: 4px;")

        self.choose_btn = QPushButton("Выбрать Excel-файл")
        self.choose_btn.clicked.connect(self._on_choose_file)

        self.district_combo = QComboBox()
        self.district_combo.addItems(get_all_districts())

        self.period_combo = QComboBox()
        self.period_combo.setEditable(True)
        self.period_combo = QComboBox()
        self.period_combo.setEditable(True)
        self.period_combo.addItems([
            "1 квартал 2024 г.",
            "2 квартал 2024 г.",
            "3 квартал 2024 г.",
            "4 квартал 2024 г.",
            "1 квартал 2025 г.",
            "2 квартал 2025 г.",
            "3 квартал 2025 г.",
            "4 квартал 2025 г.",
        ])
        self.period_combo.setMinimumWidth(170)

        self.add_btn = QPushButton("Добавить файл")
        self.add_btn.setObjectName("Secondary")
        self.add_btn.clicked.connect(self._on_add_file)

        self.process_btn = QPushButton("Запустить обработку")
        self.process_btn.clicked.connect(self._on_process)

        controls.addWidget(self.choose_btn)
        controls.addWidget(self.file_label, 1)
        controls.addWidget(QLabel("Район:"))
        controls.addWidget(self.district_combo)
        controls.addWidget(QLabel("Период:"))
        controls.addWidget(self.period_combo)
        controls.addWidget(self.add_btn)
        controls.addWidget(self.process_btn)
        root.addLayout(controls)

        # Таблица загруженных файлов
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(
            ["Район", "Период", "Файл", "Листы"]
        )
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setColumnWidth(0, 220)
        self.table.setColumnWidth(1, 140)
        self.table.setColumnWidth(2, 320)
        root.addWidget(self.table, 1)

        hint = QLabel(
            "Выберите Excel-файл (.xlsx / .xls), укажите район и период, "
            "затем нажмите «Добавить файл»."
        )

    # ---------- Обработчики ----------

    def _on_choose_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Выберите Excel-отчёт", "",
            "Excel-файлы (*.xlsx *.xls);;Все файлы (*.*)",
        )
        if not path:
            return
        self._selected_file = path
        filename = path.split("/")[-1]
        self.file_label.setText(filename)
        self.file_label.setStyleSheet(f"color: {COLOR_SIDEBAR}; padding: 4px;")

        # Автоматическое определение района по коду в начале имени файла (например, 101_230_...)
        import re
        match = re.match(r"^(\d{3})_", filename)
        if match:
            code = match.group(1)
            district_name = get_district_name(code)
            index = self.district_combo.findText(district_name)
            if index >= 0:
                self.district_combo.setCurrentIndex(index)

    def _on_add_file(self) -> None:
        if not self._selected_file:
            QMessageBox.warning(
                self, "Файл не выбран",
                "Сначала выберите Excel-файл кнопкой «Выбрать Excel-файл».",
            )
            return

        district = self.district_combo.currentText().strip()
        period = self.period_combo.currentText().strip()
        if not district or not period:
            QMessageBox.warning(
                self, "Не заданы параметры",
                "Укажите район и период.",
            )
            return

        if "_230_" not in self._selected_file:
            QMessageBox.warning(
                self, "Неверный файл",
                "Выбранный файл не похож на отчёт формы № 230.\n"
                "Убедитесь, что в названии файла есть «_230_» (например, 1175_230_2601.XLS)."
            )
            return

        import re
        # Автоопределение периода
        # Автоопределение периода
        match_period = re.search(r"_(\d{2})(\d{2})\.", self._selected_file)
        if match_period:
            year = 2000 + int(match_period.group(1))
            month = int(match_period.group(2))
            quarter = (month - 1) // 3 + 1
            period = f"{quarter} квартал {year} г."
            index = self.period_combo.findText(period)
            if index >= 0:
                self.period_combo.setCurrentIndex(index)
            else:
                # Если такого периода нет в списке, добавляем его
                self.period_combo.addItem(period)
                self.period_combo.setCurrentIndex(self.period_combo.count() - 1)
        try:
            report = make_uploaded_report(district, period, self._selected_file)
        except Exception as exc:
            QMessageBox.critical(
                self, "Ошибка чтения файла",
                f"Не удалось открыть файл:\n{exc}",
            )
            return

        # Если в списке есть демо-файлы, очищаем их перед добавлением реального
        if self._reports and self._reports[0].is_demo:
            self._reports.clear()
            self.table.setRowCount(0)
        self._append_row(report)
        self._selected_file = None
        self.file_label.setText("Файл не выбран")
        self.file_label.setStyleSheet(f"color: {COLOR_MUTED}; padding: 4px;")
        QMessageBox.information(self, "Готово", "Файл добавлен в список.")

    def _append_row(self, report) -> None:
        row = self.table.rowCount()
        self.table.insertRow(row)
        self.table.setItem(row, 0, QTableWidgetItem(report.district))
        self.table.setItem(row, 1, QTableWidgetItem(report.period))
        self.table.setItem(row, 2, QTableWidgetItem(report.file_path))
        self.table.setItem(row, 3, QTableWidgetItem(", ".join(report.sheets)))

    def _on_process(self) -> None:
        # Собираем данные прямо из таблицы на экране
        reports = []
        for row in range(self.table.rowCount()):
            district = self.table.item(row, 0).text()
            period = self.table.item(row, 1).text()
            file_path = self.table.item(row, 2).text()

            # Создаем объект на лету
            from app.data.models import UploadedReport
            rep = UploadedReport(
                district=district,
                period=period,
                file_path=file_path,
                sheets=[],  # Листы нам сейчас не критичны
                is_demo=False
            )
            reports.append(rep)

        # Передаем в главное окно
        if hasattr(self.main_window, "process_files"):
            self.main_window.process_files(reports)
    def get_reports(self) -> list:
        """Возвращает список загруженных файлов."""
        return list(self._reports)