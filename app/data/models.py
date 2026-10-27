"""
Модели данных приложения.
Все поля — простые, чтобы легко было потом заменить
демо-данные на реальные Excel-отчёты.
"""
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class UploadedReport:
    """Загруженный Excel-отчёт, привязанный к району и периоду."""
    district: str                # район
    period: str                  # например, "2024-Q1"
    file_path: str               # путь к файлу (или "<демо>")
    sheets: list[str] = field(default_factory=list)
    is_demo: bool = False        # True для тестовых данных


@dataclass
class ValidationError:
    """Одна найденная формальная ошибка/предупреждение."""
    district: str
    file_name: str
    sheet: str
    field: str
    description: str
    severity: str = "error"      # "error" | "warning"


@dataclass
class SummaryRow:
    """Строка сводной таблицы по району."""
    district: str
    indicator: str
    current_value: Optional[float]
    previous_value: Optional[float]
    abs_diff: Optional[float] = None
    pct_diff: Optional[float] = None
    direction: str = ""          # "рост" | "снижение" | "без изменений" | "—"