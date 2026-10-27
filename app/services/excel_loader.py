"""
Базовая загрузка Excel-файлов (.xlsx / .xls) через pandas.
Структура таблиц пока не фиксируется жёстко — это задел
для последующей настройки под реальные шаблоны.
"""
from pathlib import Path
from typing import Optional

import pandas as pd

from app.data.models import UploadedReport


def load_excel_sheets(file_path: str) -> dict[str, pd.DataFrame]:
    """
    Читает все листы Excel-файла.
    Возвращает словарь {имя_листа: DataFrame}.
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"Файл не найден: {file_path}")

    suffix = path.suffix.lower()
    if suffix not in (".xlsx", ".xls"):
        raise ValueError("Поддерживаются только файлы .xlsx и .xls")

    # sheet_name=None — читаем все листы
    engine = "openpyxl" if suffix == ".xlsx" else "xlrd"
    sheets = pd.read_excel(path, sheet_name=None, engine=engine, header=None)
    return sheets


def make_uploaded_report(district: str, period: str, file_path: str) -> UploadedReport:
    """
    Создаёт объект UploadedReport после лёгкой проверки файла.
    Не бросает исключение, если файл удалось открыть.
    """
    sheets = load_excel_sheets(file_path)
    return UploadedReport(
        district=district,
        period=period,
        file_path=str(file_path),
        sheets=list(sheets.keys()),
        is_demo=False,
    )


def try_load_sheets_safe(file_path: str) -> Optional[dict[str, pd.DataFrame]]:
    """Безопасная загрузка: возвращает None при ошибке."""
    try:
        return load_excel_sheets(file_path)
    except Exception:
        return None