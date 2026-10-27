"""
Страница черновика докладной: формирование, редактирование, экспорт в .docx.
"""
import os
from pathlib import Path

from PyQt6.QtWidgets import (
    QFileDialog, QHBoxLayout, QLabel, QMessageBox, QPushButton,
    QTextEdit, QVBoxLayout, QWidget,
)

from app.data.models import SummaryRow
from app.services.dynamics import build_dynamics_text
from app.theme import COLOR_MUTED

# Импортируем логику из вашего скрипта
try:
    from app.services.excel_to_word import main as generate_word
except ImportError:
    generate_word = None


class ReportPage(QWidget):
    """Формирование и экспорт черновика докладной."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._rows: list[SummaryRow] = []
        self._period: str = "—"
        self._reports = []  # Сюда будем сохранять загруженные отчеты

        root = QVBoxLayout(self)
        root.setContentsMargins(24, 24, 24, 24)
        root.setSpacing(12)

        title = QLabel("Докладная")
        title.setObjectName("PageTitle")
        root.addWidget(title)

        hint = QLabel(
            "Если вы загрузили реальные файлы (текущий, предыдущий и АППГ), "
            "нажмите «Сформировать Word (Реальные данные)». "
            "Если файлов нет — сформируется тестовый черновик."
        )
        hint.setStyleSheet(f"color: {COLOR_MUTED};")
        hint.setWordWrap(True)
        root.addWidget(hint)

        buttons = QHBoxLayout()

        self.generate_real_btn = QPushButton("📄 Сформировать Word (Реальные данные)")
        self.generate_real_btn.clicked.connect(self._generate_real_word)

        self.generate_test_btn = QPushButton("Сформировать тестовый текст")
        self.generate_test_btn.setObjectName("Secondary")
        self.generate_test_btn.clicked.connect(self._generate_test_draft)

        buttons.addWidget(self.generate_real_btn)
        buttons.addWidget(self.generate_test_btn)
        buttons.addStretch()
        root.addLayout(buttons)

        self.editor = QTextEdit()
        self.editor.setPlaceholderText(
            "Здесь появится автоматически сформированный черновик. "
            "Нажмите «Сформировать тестовый текст» или «Сформировать Word»."
        )
        root.addWidget(self.editor, 1)

    # ---------- Публичное API ----------

    def set_context(self, rows: list[SummaryRow], period: str, reports: list = None) -> None:
        """Передаёт данные для формирования черновика."""
        self._rows = list(rows)
        self._period = period
        if reports is not None:
            self._reports = reports

    # ---------- Генерация Word из реальных Excel ----------

    def _generate_real_word(self) -> None:
        if generate_word is None:
            QMessageBox.critical(self, "Ошибка", "Файл excel_to_word.py не найден в папке services.")
            return

        if not self._reports:
            QMessageBox.warning(
                self, "Нет файлов",
                "Сначала загрузите Excel-файлы в разделе «Загрузка отчётов»."
            )
            return

        # Ищем файлы по периодам (предполагаем, что пользователь загрузил их в нужном порядке)
        # В базовой версии берем первый файл как текущий, остальные ищем в той же папке
        current_file = None
        for r in self._reports:
            if os.path.exists(r.file_path):
                current_file = r.file_path
                break

        if not current_file:
            QMessageBox.warning(self, "Ошибка", "Не удалось найти реальный Excel-файл среди загруженных.")
            return

        # Спрашиваем, куда сохранить Word
        save_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить докладную в Word", "Доклад_аварийное_жильё.docx",
            "Документы Word (*.docx)",
        )
        if not save_path:
            return

        # Формируем аргументы для вызова вашего скрипта
        # (В простейшем случае ищем prev и appg в той же папке, где лежит текущий файл)
        search_dir = os.path.dirname(current_file)

        args = [
            "--current", current_file,
            "--search", search_dir,
            "--out", save_path
        ]

        try:
            # Запускаем ваш скрипт
            result_code = generate_word(args)
            if result_code == 0:
                QMessageBox.information(
                    self, "Успех",
                    f"Документ успешно сформирован и сохранён:\n{Path(save_path).resolve()}"
                )
                # Показываем путь в текстовом поле для информации
                self.editor.setPlainText(f"Документ сохранён по пути:\n{save_path}")
            else:
                QMessageBox.critical(self, "Ошибка", "Скрипт завершился с ошибкой. Проверьте консоль.")
        except Exception as exc:
            QMessageBox.critical(
                self, "Ошибка экспорта",
                f"Не удалось сформировать Word:\n{exc}",
            )

    # ---------- Генерация тестового текста ----------

    def _generate_test_draft(self) -> None:
        try:
            self.editor.setPlainText(self._build_draft_text())
        except Exception as exc:
            QMessageBox.critical(
                self, "Ошибка формирования",
                f"Не удалось сформировать черновик:\n{exc}",
            )

    def _build_draft_text(self) -> str:
        if not self._rows:
            return (
                "ЧЕРНОВИК ДОКЛАДНОЙ ЗАПИСКИ\n\n"
                "ТЕСТОВЫЕ ДАННЫЕ.\n\n"
                "Сводная статистика за период не сформирована: "
                "нет загруженных данных."
            )

        districts = sorted({r.district for r in self._rows})
        total_rows = [r for r in self._rows if "итого" in r.district.lower()]
        dyn_text = build_dynamics_text(self._rows)

        lines = []
        lines.append("ЧЕРНОВИК ДОКЛАДНОЙ ЗАПИСКИ")
        lines.append("")
        lines.append(f"Период: {self._period}")
        lines.append("Источник данных: ТЕСТОВЫЕ демонстрационные данные.")
        lines.append("")
        lines.append("1. Общие сведения")
        lines.append(
            f"В сводке участвуют данные по {len(districts)} районам(у). "
            f"Всего строк показателей: {len(self._rows)}."
        )
        lines.append("")
        lines.append("2. Краткая характеристика динамики")
        lines.append(dyn_text)
        lines.append("")

        if total_rows:
            lines.append("3. Итоговые показатели по области (тест)")
            for r in total_rows:
                pct = "—" if r.pct_diff is None else f"{r.pct_diff:+.1f} %"
                lines.append(
                    f"— {r.indicator}: текущий период {r.current_value}, "
                    f"предыдущий {r.previous_value}, "
                    f"разница {r.abs_diff}, изменение {pct} ({r.direction})."
                )
            lines.append("")

        lines.append("Примечание: формулировки построены только по доступным "
                     "тестовым числам; выводы о причинах не делаются.")
        return "\n".join(lines)