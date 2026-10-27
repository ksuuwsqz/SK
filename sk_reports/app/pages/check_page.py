"""
Страница проверки данных: таблица ошибок с фильтрами и подсветкой.
"""
from PyQt6.QtGui import QBrush, QColor
from PyQt6.QtWidgets import (
    QComboBox, QHBoxLayout, QLabel, QPushButton, QTableWidget,
    QTableWidgetItem, QVBoxLayout, QWidget,
)

from app.data.models import ValidationError
from app.theme import COLOR_ERR, COLOR_MUTED, COLOR_OK


class CheckPage(QWidget):
    """Страница проверки Excel-отчётов на формальные ошибки."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._errors: list[ValidationError] = []

        root = QVBoxLayout(self)
        root.setContentsMargins(24, 24, 24, 24)
        root.setSpacing(12)

        title = QLabel("Проверка данных")
        title.setObjectName("PageTitle")
        root.addWidget(title)

        hint = QLabel(
            "Ниже перечислены найденные формальные ошибки и предупреждения."
        )
        hint.setStyleSheet(f"color: {COLOR_MUTED};")
        hint.setWordWrap(True)
        root.addWidget(hint)

        # Панель фильтров
        filters = QHBoxLayout()
        filters.setSpacing(8)

        self.district_filter = QComboBox()
        self.district_filter.addItem("Все районы")
        self.district_filter.currentIndexChanged.connect(self._apply_filters)

        self.severity_filter = QComboBox()
        self.severity_filter.addItems(["Все типы", "error", "warning"])
        self.severity_filter.currentIndexChanged.connect(self._apply_filters)

        refresh_btn = QPushButton("Обновить список")
        refresh_btn.setObjectName("Secondary")
        refresh_btn.clicked.connect(self._refresh)

        filters.addWidget(QLabel("Район:"))
        filters.addWidget(self.district_filter)
        filters.addWidget(QLabel("Тип ошибки:"))
        filters.addWidget(self.severity_filter)
        filters.addStretch()
        filters.addWidget(refresh_btn)
        root.addLayout(filters)

        # Таблица ошибок
        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(
            ["Район", "Файл", "Лист", "Поле", "Описание ошибки", "Статус"]
        )
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setColumnWidth(0, 180)
        self.table.setColumnWidth(1, 200)
        self.table.setColumnWidth(2, 120)
        self.table.setColumnWidth(3, 160)
        self.table.setColumnWidth(4, 360)
        root.addWidget(self.table, 1)

    # ---------- Публичное API ----------

    def set_errors(self, errors: list[ValidationError]) -> None:
        """Устанавливает список ошибок и обновляет таблицу."""
        self._errors = list(errors)

        # Заполняем фильтр районов
        current = self.district_filter.currentText()
        self.district_filter.blockSignals(True)
        self.district_filter.clear()
        self.district_filter.addItem("Все районы")
        districts = sorted({e.district for e in self._errors})
        self.district_filter.addItems(districts)
        idx = self.district_filter.findText(current)
        if idx >= 0:
            self.district_filter.setCurrentIndex(idx)
        self.district_filter.blockSignals(False)

        self._refresh()

    # ---------- Внутренняя логика ----------

    def _refresh(self) -> None:
        self._apply_filters()

    def _apply_filters(self) -> None:
        district = self.district_filter.currentText()
        severity = self.severity_filter.currentText()

        filtered = [
            e for e in self._errors
            if (district == "Все районы" or e.district == district)
            and (severity == "Все типы" or e.severity == severity)
        ]
        self._fill_table(filtered)

    def _fill_table(self, errors: list[ValidationError]) -> None:
        self.table.setRowCount(0)
        for e in errors:
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.table.setItem(row, 0, QTableWidgetItem(e.district))
            self.table.setItem(row, 1, QTableWidgetItem(e.file_name))
            self.table.setItem(row, 2, QTableWidgetItem(e.sheet))
            self.table.setItem(row, 3, QTableWidgetItem(e.field))
            self.table.setItem(row, 4, QTableWidgetItem(e.description))
            status = "Ошибка" if e.severity == "error" else "Предупреждение"
            self.table.setItem(row, 5, QTableWidgetItem(status))

            # Цветовая подсветка строки по типу ошибки
            color = QColor("#FBE9E7") if e.severity == "error" else QColor("#FFF6E0")
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if item:
                    item.setBackground(QBrush(color))