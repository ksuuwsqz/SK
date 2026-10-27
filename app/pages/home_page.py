"""
Главная страница: краткие KPI и сводка по динамике.
"""
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QFrame, QGridLayout, QLabel, QVBoxLayout, QWidget,
)

from app.theme import COLOR_MUTED, COLOR_SIDEBAR


class Card(QFrame):
    """Карточка с заголовком и значением."""
    def __init__(self, title: str, value: str = "—", parent=None):
        super().__init__(parent)
        self.setObjectName("Card")
        self.setMinimumHeight(96)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 12, 16, 12)

        self.title_label = QLabel(title)
        self.title_label.setObjectName("CardTitle")
        self.value_label = QLabel(value)
        self.value_label.setObjectName("CardValue")

        layout.addWidget(self.title_label)
        layout.addWidget(self.value_label)
        layout.addStretch()

    def set_value(self, value: str) -> None:
        self.value_label.setText(str(value))


class HomePage(QWidget):
    """Главная страница с KPI и краткой сводкой."""

    def __init__(self, parent=None):
        super().__init__(parent)

        root = QVBoxLayout(self)
        root.setContentsMargins(24, 24, 24, 24)
        root.setSpacing(16)

        title = QLabel("Главная")
        title.setObjectName("PageTitle")
        root.addWidget(title)

        subtitle = QLabel(
            "Проект: «Формирование автоматизированных отчётов "
            "для Следственного комитета РФ по Челябинской области». "
            "Данные на экране — ТЕСТОВЫЕ."
        )
        subtitle.setStyleSheet(f"color: {COLOR_MUTED};")
        subtitle.setWordWrap(True)
        root.addWidget(subtitle)

        # Сетка карточек
        grid = QGridLayout()
        grid.setSpacing(12)

        self.card_period = Card("Выбранный период")
        self.card_reports = Card("Загружено отчётов")
        self.card_districts = Card("Количество районов")
        self.card_errors = Card("Найдено ошибок")

        grid.addWidget(self.card_period, 0, 0)
        grid.addWidget(self.card_reports, 0, 1)
        grid.addWidget(self.card_districts, 0, 2)
        grid.addWidget(self.card_errors, 0, 3)
        root.addLayout(grid)

        # Сводка динамики
        summary_title = QLabel("Краткая сводка по динамике")
        summary_title.setStyleSheet(
            f"font-size: 13pt; font-weight: 600; color: {COLOR_SIDEBAR};"
        )
        root.addWidget(summary_title)

        self.summary_label = QLabel("Данные ещё не рассчитаны.")
        self.summary_label.setWordWrap(True)
        self.summary_label.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.summary_label.setStyleSheet(
            f"background: white; border: 1px solid #E1E5EC; "
            f"border-radius: 10px; padding: 14px; color: {COLOR_MUTED};"
        )
        root.addWidget(self.summary_label, 1)

    def update_state(self, period: str, reports_count: int, districts_count: int,
                     errors_count: int, dynamics_text: str) -> None:
        """Обновляет KPI-карточки и сводку."""
        self.card_period.set_value(period or "—")
        self.card_reports.set_value(str(reports_count))
        self.card_districts.set_value(str(districts_count))
        self.card_errors.set_value(str(errors_count))
        self.summary_label.setText(dynamics_text or "Данные ещё не рассчитаны.")