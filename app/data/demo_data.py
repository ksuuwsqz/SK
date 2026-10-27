"""
ДЕМОНСТРАЦИОННЫЕ ДАННЫЕ. ТОЛЬКО ДЛЯ ТЕСТИРОВАНИЯ ИНТЕРФЕЙСА.
Все цифры — придуманные, тестовые; они не отражают реальную
статистику Следственного комитета.
"""
from app.data.models import UploadedReport, ValidationError, SummaryRow

# Список условных районов (только для демо)
DEMO_DISTRICTS = [
    "Тестовый район № 1",
    "Тестовый район № 2",
    "Тестовый район № 3",
    "Тестовый район № 4",
]


def demo_uploaded_reports() -> list[UploadedReport]:
    """Пакет демо-отчётов, привязанных к районам за один период."""
    return [
        UploadedReport(
            district=d,
            period="ДЕМО-период",
            file_path="<демо-данные>",
            sheets=["Сводка", "Приложение 1"],
            is_demo=True,
        )
        for d in DEMO_DISTRICTS
    ]


def demo_validation_errors() -> list[ValidationError]:
    """Примеры тестовых ошибок для раздела «Проверка данных»."""
    return [
        ValidationError(
            district=DEMO_DISTRICTS[0],
            file_name="<демо-данные>",
            sheet="Сводка",
            field="Кол-во сообщений",
            description="Пустое значение в обязательном поле (ТЕСТ).",
            severity="error",
        ),
        ValidationError(
            district=DEMO_DISTRICTS[1],
            file_name="<демо-данные>",
            sheet="Сводка",
            field="Сроки рассмотрения",
            description="Некорректное числовое значение: «н/д» (ТЕСТ).",
            severity="error",
        ),
        ValidationError(
            district=DEMO_DISTRICTS[2],
            file_name="<демо-данные>",
            sheet="Приложение 1",
            field="Итого",
            description="Не найдено обязательное поле «Итого» (ТЕСТ).",
            severity="warning",
        ),
    ]


def demo_summary_rows() -> list[SummaryRow]:
    """
    Демонстрационные показатели по районам и по области.
    Все значения — ТЕСТОВЫЕ.
    """
    indicators = [
        ("Количество поступивших сообщений", 120, 100),
        ("Принято процессуальных решений", 80, 90),
    ]
    rows: list[SummaryRow] = []
    for idx, district in enumerate(DEMO_DISTRICTS):
        for name, cur_base, prev_base in indicators:
            cur = cur_base + idx * 7
            prev = prev_base + idx * 5
            rows.append(_make_row(district, name, cur, prev))

    # Итог по области
    for name, cur_base, prev_base in indicators:
        total_cur = sum(cur_base + i * 7 for i in range(len(DEMO_DISTRICTS)))
        total_prev = sum(prev_base + i * 5 for i in range(len(DEMO_DISTRICTS)))
        rows.append(_make_row("Челябинская область (итого)", name, total_cur, total_prev))

    return rows


def _make_row(district: str, indicator: str, cur: float, prev: float) -> SummaryRow:
    """Вспомогательный расчёт разницы для демо-строки."""
    abs_diff = cur - prev
    if prev == 0:
        pct_diff = None
    else:
        pct_diff = round((cur - prev) / prev * 100, 1)
    if abs_diff > 0:
        direction = "рост"
    elif abs_diff < 0:
        direction = "снижение"
    else:
        direction = "без изменений"
    return SummaryRow(
        district=district,
        indicator=indicator,
        current_value=float(cur),
        previous_value=float(prev),
        abs_diff=abs_diff,
        pct_diff=pct_diff,
        direction=direction,
    )