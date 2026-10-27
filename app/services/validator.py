"""
Базовая проверка Excel-отчётов на формальные ошибки.

Правила вынесены в конфигурацию REQUIRED_FIELDS / NUMERIC_FIELDS,
чтобы позже легко было заменить их на реальные требования.
"""
from typing import Optional

import pandas as pd

from app.data.models import UploadedReport, ValidationError

# Правила-заготовки под будущую настройку.
# Ключ — имя листа (None = любой лист), значение — список обязательных полей.
REQUIRED_FIELDS: dict[Optional[str], list[str]] = {
    None: ["Район", "Период", "Показатель", "Значение"],
}

# Поля, которые должны быть числовыми
NUMERIC_FIELDS: list[str] = ["Значение"]


def validate_report(report: UploadedReport, sheets: dict[str, pd.DataFrame]) -> list[ValidationError]:
    """
    Проверяет один отчёт и возвращает список найденных ошибок.
    Логика базовая; списки полей можно менять в REQUIRED_FIELDS.
    """
    errors: list[ValidationError] = []
    file_name = report.file_path.split("/")[-1] if report.file_path else "<демо>"

    for sheet_name, df in sheets.items():
        # Ищем строку заголовков — берём первую непустую
        header_row_idx = _find_header_row(df)
        if header_row_idx is None:
            errors.append(ValidationError(
                district=report.district,
                file_name=file_name,
                sheet=sheet_name,
                field="—",
                description="Не найдена строка заголовков (тест-правило).",
                severity="error",
            ))
            continue

        header = [str(x).strip() if pd.notna(x) else "" for x in df.iloc[header_row_idx].tolist()]
        body = df.iloc[header_row_idx + 1:].reset_index(drop=True)

        # 1) Обязательные поля
        required = REQUIRED_FIELDS.get(sheet_name, []) + REQUIRED_FIELDS.get(None, [])
        for field_name in required:
            if field_name not in header:
                errors.append(ValidationError(
                    district=report.district,
                    file_name=file_name,
                    sheet=sheet_name,
                    field=field_name,
                    description=f"Отсутствует обязательное поле «{field_name}» (тест-правило).",
                    severity="error",
                ))

        # 2) Пустые значения
        for col_name in required:
            if col_name not in header:
                continue
            col_idx = header.index(col_name)
            empty_rows = body.index[body.iloc[:, col_idx].isna()].tolist()
            if empty_rows:
                errors.append(ValidationError(
                    district=report.district,
                    file_name=file_name,
                    sheet=sheet_name,
                    field=col_name,
                    description=f"Пустые значения в поле «{col_name}» "
                                f"(строк: {len(empty_rows)}) (тест-правило).",
                    severity="warning",
                ))

        # 3) Числовые поля
        for col_name in NUMERIC_FIELDS:
            if col_name not in header:
                continue
            col_idx = header.index(col_name)
            series = body.iloc[:, col_idx]
            bad = series[series.notna() & pd.to_numeric(series, errors="coerce").isna()]
            if not bad.empty:
                errors.append(ValidationError(
                    district=report.district,
                    file_name=file_name,
                    sheet=sheet_name,
                    field=col_name,
                    description=f"Некорректные числовые значения в поле «{col_name}» "
                                f"(строк: {len(bad)}) (тест-правило).",
                    severity="error",
                ))
    return errors


def _find_header_row(df: pd.DataFrame) -> Optional[int]:
    """Ищет первую строку, где есть хотя бы 2 непустые ячейки."""
    for idx in range(min(len(df), 10)):
        row = df.iloc[idx]
        non_empty = sum(1 for v in row.tolist() if pd.notna(v) and str(v).strip() != "")
        if non_empty >= 2:
            return idx
    return None


def validate_demo() -> list[ValidationError]:
    """Возвращает заранее подготовленные демо-ошибки."""
    from app.data.demo_data import demo_validation_errors
    return demo_validation_errors()