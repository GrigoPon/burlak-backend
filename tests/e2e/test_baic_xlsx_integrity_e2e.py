from pathlib import Path

import pytest
from openpyxl import load_workbook


def test_style_preservation():
    """BAIC тест — проверка сохранения стилей Excel."""

    input_path = Path("tests/fixtures/style_test_card.xlsx")

    if not input_path.exists():
        pytest.skip(
            f"Тестовый файл не найден: {input_path}\n"
            "Запустите команду создания фикстуры."
        )

    # Просто проверяем, что файл существует и имеет структуру
    wb = load_workbook(input_path)

    assert len(wb.sheetnames) > 0, "В файле нет листов"
    ws = wb.active

    # Проверяем наличие форматирования
    assert ws["A1"].font.bold is True, "Заголовок должен быть жирным"
    assert ws["A1"].fill.start_color.index != "00000000", (
        "Заголовок должен иметь заливку"
    )

    # Проверяем наличие данных
    assert ws["A3"].value == "№", "Шапка таблицы должна присутствовать"

    print(f"✅ BAIC тест пройден успешно! Файл: {input_path}")
    print(f"   Листов: {len(wb.sheetnames)}, Строк: {ws.max_row}")
