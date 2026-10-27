"""
Страница сводки и динамики показателей по районам и области.
"""
from PyQt6.QtGui import QBrush, QColor
from PyQt6.QtWidgets import (
    QLabel, QPushButton, QTableWidget, QTableWidgetItem,
    QVBoxLayout, QWidget,
)

from app.data.models import SummaryRow
from app.theme import COLOR_MUTED


class SummaryPage(QWidget):
    """Сводная таблица показателей с расчётом динамики."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._rows: list[SummaryRow] = []

        root = QVBoxLayout(self)
        root.setContentsMargins(24, 24, 24, 24)
        root.setSpacing(12)

        title = QLabel("Сводка и динамика")
        title.setObjectName("PageTitle")
        root.addWidget(title)

        hint = QLabel(
            "Сравнение показателей текущего и предыдущего периодов. "
            "Рост отмечен зелёным, снижение — красным, без изменений — серым."
        )
        hint.setStyleSheet(f"color: {COLOR_MUTED};")
        hint.setWordWrap(True)
        root.addWidget(hint)

        refresh_btn = QPushButton("Пересчитать сводку")
        refresh_btn.setObjectName("Secondary")
        refresh_btn.clicked.connect(self._refresh)
        root.addWidget(refresh_btn)

        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels(
            ["Район", "Показатель", "Текущий период",
             "Предыдущий период", "Разница", "Изм., %", "Тенденция"]
        )
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setColumnWidth(0, 240)
        self.table.setColumnWidth(1, 280)
        self.table.setColumnWidth(2, 140)
        self.table.setColumnWidth(3, 160)
        self.table.setColumnWidth(4, 100)
        self.table.setColumnWidth(5, 100)
        root.addWidget(self.table, 1)

    # ---------- Публичное API ----------

    def set_rows(self, rows: list[SummaryRow]) -> None:
        self._rows = list(rows)
        self._refresh()

    # ---------- Заполнение таблицы ----------

    def _refresh(self) -> None:
        self.table.setRowCount(0)
        for r in self._rows:
            row = self.table.rowCount()
            self.table.insertRow(row)

            self.table.setItem(row, 0, QTableWidgetItem(r.district))
            self.table.setItem(row, 1, QTableWidgetItem(r.indicator))
            self.table.setItem(row, 2, QTableWidgetItem(_fmt(r.current_value)))
            self.table.setItem(row, 3, QTableWidgetItem(_fmt(r.previous_value)))
            self.table.setItem(row, 4, QTableWidgetItem(_fmt(r.abs_diff)))
            self.table.setItem(row, 5, QTableWidgetItem(
                "—" if r.pct_diff is None else f"{r.pct_diff:+.1f} %"
            ))

            arrow = {"рост": "▲ рост", "снижение": "▼ снижение",
                     "без изменений": "= без изменений"}.get(r.direction, "—")
            self.table.setItem(row, 6, QTableWidgetItem(arrow))

            # Подсветка по тенденции
            if r.direction == "рост":
                color = QColor("#E9F6EC")
            elif r.direction == "снижение":
                color = QColor("#FBE9E7")
            else:
                color = QColor("#F3F4F6")
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if item:
                    item.setBackground(QBrush(color))


def _fmt(value) -> str:
    """Форматирует число или возвращает «—»."""
    if value is None:
        return "—"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)