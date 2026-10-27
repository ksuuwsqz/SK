"""
Расчёт динамики показателей между текущим и предыдущим периодами.
"""
from app.data.models import SummaryRow


def enrich_dynamics(rows: list[SummaryRow]) -> list[SummaryRow]:
    """
    Заполняет abs_diff, pct_diff и direction для строк.
    Если процент посчитать нельзя (prev = 0) — ставим None.
    """
    for r in rows:
        if r.current_value is None or r.previous_value is None:
            r.abs_diff = None
            r.pct_diff = None
            r.direction = "—"
            continue
        r.abs_diff = r.current_value - r.previous_value
        if r.previous_value == 0:
            r.pct_diff = None
        else:
            r.pct_diff = round((r.abs_diff / r.previous_value) * 100, 1)

        if r.abs_diff > 0:
            r.direction = "рост"
        elif r.abs_diff < 0:
            r.direction = "снижение"
        else:
            r.direction = "без изменений"
    return rows


def build_dynamics_text(rows: list[SummaryRow]) -> str:
    """
    Краткая сводка динамики только по имеющимся строкам.
    Без выводов, которых нет в данных.
    """
    if not rows:
        return "Данные для сводки по динамике отсутствуют."

    growing = [r for r in rows if r.direction == "рост"]
    falling = [r for r in rows if r.direction == "снижение"]
    flat = [r for r in rows if r.direction == "без изменений"]

    parts = []
    parts.append(f"Всего показателей в сводке: {len(rows)}.")
    parts.append(f"Показателей с ростом: {len(growing)}.")
    parts.append(f"Показателей со снижением: {len(falling)}.")
    parts.append(f"Показателей без изменений: {len(flat)}.")
    return " ".join(parts)