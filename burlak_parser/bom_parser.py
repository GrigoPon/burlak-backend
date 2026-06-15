"""Модуль чтения BOM-файла.

Формат: .xlsx (таблица на китайском/английском языках).
Структура (по спецификации ТЗ):
  - Колонка 1: Порядковый номер
  - Колонки 2-5: Игнорируются
  - Колонка 6: Каталожный номер / Парт-номер
  - Колонка 7: Наименование детали (кит.)
  - Колонка 8: Наименование детали (англ.)
  - Последующие колонки: Коды комплектаций (матрица применимости)
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

logger = logging.getLogger(__name__)

# Константы для поиска колонок по ключевым словам (китайский / английский)
COL_PART_NO_KEYWORDS = ["零件号", "partno", "part no", "part_no", "part number", "料号"]
COL_NAME_CN_KEYWORDS = [
    "零件名称(中文）",
    "零件名称(中文)",
    "零件名称（中文）",
    "零件名称（中文)",
    "物料名称/描述",
    "零件名称",
    "物料名称",
    "描述",
]
COL_NAME_EN_KEYWORDS = ["零件名称(英文）", "零件名称(英文)", "零件名称（英文）", "part name(en)", "part name(en）"]
COL_QTY_KEYWORDS = ["用量", "qty", "数量", "单车用量", "quantity"]


@dataclass
class PartInfo:
    """Информация о детали из BOM."""
    part_number: str
    name_cn: str
    name_en: str
    # Количество для конкретной комплектации (будет заполнено после выбора)
    quantity: float = 0.0
    # Номера/коды комплектаций, для которых указана деталь
    applicable_configs: List[str] = field(default_factory=list)


@dataclass
class BOMData:
    """Результат парсинга BOM-файла."""
    parts: Dict[str, PartInfo]  # part_number -> PartInfo
    config_names: List[str]  # названия колонок комплектаций
    config_quantities: Dict[str, Dict[str, float]]  # config_name -> {part_number -> qty}
    source_file: str = ""


def _normalize(s: Optional[str]) -> str:
    """Привести строку к нижнему регистру, убрать пробелы и переносы строк."""
    if s is None:
        return ""
    return re.sub(r"\s+", "", str(s).lower())


def _find_header_row(ws: Worksheet) -> Optional[int]:
    """Найти строку заголовков в листе."""
    for row_idx in range(1, min(10, ws.max_row or 10) + 1):
        row_values = [ws.cell(row=row_idx, column=c).value for c in range(1, min(20, (ws.max_column or 20) + 1))]
        text = " ".join(str(v) for v in row_values if v is not None)
        # Ищем признаки строки заголовка: наличие ключевых слов "零件号" или "PartNo"
        if "零件号" in text or "partno" in text.lower():
            logger.info(f"Строка заголовков найдена на строке {row_idx}")
            return row_idx
    return None


def _detect_column_map(ws: Worksheet, header_row: int) -> Dict[str, int]:
    """Определить соответствие колонок по заголовкам.

    Returns:
        Словарь: {'part_no': int, 'name_cn': int, 'name_en': int, 'config_start': int}
    """
    col_map: Dict[str, int] = {}

    for col_idx in range(1, (ws.max_column or 200) + 1):
        cell_value = ws.cell(row=header_row, column=col_idx).value
        if cell_value is None:
            continue
        normalized = _normalize(cell_value)

        # Поиск колонки парт-номера
        if "part_no" not in col_map:
            for kw in COL_PART_NO_KEYWORDS:
                if _normalize(kw) in normalized or kw.lower() in normalized:
                    col_map["part_no"] = col_idx
                    logger.info(f"Колонка парт-номера: {col_idx} (заголовок: {cell_value})")
                    break

        # Поиск колонки названия (кит.)
        if "name_cn" not in col_map:
            for kw in COL_NAME_CN_KEYWORDS:
                if _normalize(kw) in normalized or kw.lower() in normalized:
                    col_map["name_cn"] = col_idx
                    logger.info(f"Колонка названия (кит): {col_idx} (заголовок: {cell_value})")
                    break

        # Поиск колонки названия (англ.)
        if "name_en" not in col_map:
            for kw in COL_NAME_EN_KEYWORDS:
                if _normalize(kw) in normalized or kw.lower() in normalized:
                    col_map["name_en"] = col_idx
                    logger.info(f"Колонка названия (англ): {col_idx} (заголовок: {cell_value})")
                    break

    return col_map


def _detect_config_columns(ws: Worksheet, header_row: int, part_no_col: int) -> List[int]:
    """Определить колонки комплектаций.

    Колонки комплектаций — это все колонки справа от колонки парт-номера,
    названия которых НЕ являются стандартными заголовками данных.

    Дополнительно фильтрует:
      - Metadata-колонки (MWO, даты,整车物料号)
      - VIN-разбивку (колонки с 'S'/'-' вместо чисел — это не комплектации)
    """
    standard_keywords = [
        "零件号", "partno", "part no", "零件名称(中文", "零件名称(英文",
        "part name", "用量", "qty", "度量单位", "uom", "gpc", "fnd",
        "零件成熟度", "零件层级", "level", "lou用法", "usage", "物料状态",
        "make/buy", "来源车间", "source shop", "使用工厂", "using plant",
        "目标车间", "target shop", "供应商", "supplier", "mwo单号", "mwo",
        "生效日期", "失效日期", "整车物料号", "vehicle material",
        "序号", "serial no", "cpac编码", "cpac code", "cpac描述",
        "标识", "发运", "采购", "ship", "purchase", "修订",
        "版本", "version", "有效日期", "effective date",
    ]

    candidate_cols: List[int] = []
    max_col = ws.max_column or 200

    for col_idx in range(part_no_col + 1, max_col + 1):
        cell_value = ws.cell(row=header_row, column=col_idx).value
        if cell_value is None:
            continue
        normalized = _normalize(cell_value)

        if isinstance(cell_value, (int, float)):
            continue

        is_standard = False
        for kw in standard_keywords:
            if _normalize(kw) in normalized:
                is_standard = True
                break

        if not is_standard and len(str(cell_value).strip()) > 2:
            candidate_cols.append(col_idx)

    # ── Пост-фильтрация: отсеять metadata и VIN-разбивку ──
    # VIN-колонки содержат 'S' (Same — «такая же») или '-', а не числа.
    # Проверяем выборку строк данных: если ни одного числа — это не комплектация.
    #
    # Алгоритм (универсальный, не привязан к конкретной структуре BOM):
    #   1. Для каждой колонки-кандидата проверяем наличие числовых значений.
    #   2. Находим ПЕРВУЮ колонку без чисел (VIN-разбивка начинается здесь).
    #   3. Обрезаем ВСЕ колонки начиная с этой — всё, что после, не комплектации.
    data_start = header_row + 1
    sample_end = min(data_start + 50, ws.max_row or data_start + 50)

    # Сначала определяем, какие колонки имеют числовые значения
    column_has_numbers: Dict[int, bool] = {}
    for col_idx in candidate_cols:
        has_numeric = False
        for r in range(data_start, sample_end + 1):
            v = ws.cell(row=r, column=col_idx).value
            if v is not None:
                if isinstance(v, (int, float)):
                    has_numeric = True
                    break
                elif isinstance(v, str):
                    stripped = v.strip()
                    if stripped not in ('S', '-', 's', ''):
                        try:
                            float(stripped)
                            has_numeric = True
                            break
                        except ValueError:
                            pass
        column_has_numbers[col_idx] = has_numeric

    # Найти первую колонку без чисел после группы колонок с числами
    first_non_numeric_after_numeric: Optional[int] = None
    found_numeric = False
    for col_idx in candidate_cols:
        if column_has_numbers[col_idx]:
            found_numeric = True
        elif found_numeric:
            # Нашли нечисловую колонку после числовых — здесь граница
            first_non_numeric_after_numeric = col_idx
            break

    # Отбираем только колонки до границы VIN-разбивки
    config_cols: List[int] = []
    if first_non_numeric_after_numeric is not None:
        for col_idx in candidate_cols:
            if col_idx < first_non_numeric_after_numeric and column_has_numbers[col_idx]:
                config_cols.append(col_idx)
        logger.info(
            f"VIN-разбивка обнаружена с колонки {first_non_numeric_after_numeric} "
            f"(нет числовых значений). Комплектаций отобрано: {len(config_cols)}"
        )
    else:
        # Нет явной границы — берём все колонки с числами
        config_cols = [c for c in candidate_cols if column_has_numbers[c]]

    if not config_cols and candidate_cols:
        # Если ни одна колонка не имеет чисел — вероятно, другой формат BOM.
        # Берём все кандидаты (старое поведение).
        logger.warning(
            "Не найдено колонок с числовыми значениями — "
            "используются все колонки-кандидаты (%d шт.)", len(candidate_cols)
        )
        config_cols = list(candidate_cols)

    logger.info(f"Найдено колонок комплектаций: {len(config_cols)} (с колонки {config_cols[0] if config_cols else '?'})")
    return config_cols


def parse_bom(file_path: str) -> BOMData:
    """Разобрать BOM-файл и вернуть структурированные данные.

    Args:
        file_path: Путь к .xlsx файлу BOM.

    Returns:
        BOMData со всеми извлечёнными данными.
    """
    logger.info(f"Загрузка BOM-файла: {file_path}")

    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]
    logger.info(f"Активный лист: {sheet_name} (строк: {ws.max_row}, колонок: {ws.max_column})")

    # Найти строку заголовков
    header_row = _find_header_row(ws)
    if header_row is None:
        raise ValueError("Не удалось найти строку заголовков в BOM-файле. "
                         "Убедитесь, что файл содержит строку с '零件号' или 'PartNo'.")

    # Определить карту колонок
    col_map = _detect_column_map(ws, header_row)

    part_no_col = col_map.get("part_no")
    name_cn_col = col_map.get("name_cn")
    name_en_col = col_map.get("name_en")

    if part_no_col is None:
        raise ValueError("Не удалось найти колонку с парт-номерами. "
                         "Искались ключевые слова: 零件号, PartNo, 料号")

    # Определить колонки комплектаций
    config_cols = _detect_config_columns(ws, header_row, part_no_col)
    config_names: List[str] = []
    for col_idx in config_cols:
        name = str(ws.cell(row=header_row, column=col_idx).value or f"Config_{col_idx}")
        config_names.append(name)

    # Парсинг данных
    parts: Dict[str, PartInfo] = {}
    config_quantities: Dict[str, Dict[str, float]] = {name: {} for name in config_names}
    data_start = header_row + 1

    for row_idx in range(data_start, ws.max_row + 1):
        part_no = ws.cell(row=row_idx, column=part_no_col).value
        if part_no is None:
            continue
        part_no = str(part_no).strip()
        if not part_no or part_no.startswith("~$"):
            continue

        name_cn = ""
        if name_cn_col:
            name_cn = str(ws.cell(row=row_idx, column=name_cn_col).value or "").strip()

        name_en = ""
        if name_en_col:
            name_en = str(ws.cell(row=row_idx, column=name_en_col).value or "").strip()

        # Суммируем количества для повторяющихся парт-номеров
        if part_no in parts:
            # Названия могут быть разными, но берём первое (или можно объединять)
            existing = parts[part_no]
            if not existing.name_cn and name_cn:
                existing.name_cn = name_cn
            if not existing.name_en and name_en:
                existing.name_en = name_en
        else:
            parts[part_no] = PartInfo(part_number=part_no, name_cn=name_cn, name_en=name_en)

        part = parts[part_no]

        # Извлечение количества для каждой комплектации (СУММИРУЕМ, а не перезаписываем)
        for i, col_idx in enumerate(config_cols):
            qty_val = ws.cell(row=row_idx, column=col_idx).value
            if qty_val is not None and isinstance(qty_val, (int, float)) and qty_val > 0:
                config_name = config_names[i]
                current_qty = config_quantities[config_name].get(part_no, 0.0)
                config_quantities[config_name][part_no] = current_qty + float(qty_val)
                if config_name not in part.applicable_configs:
                    part.applicable_configs.append(config_name)

    wb.close()

    logger.info(f"Загружено деталей: {len(parts)}")
    logger.info(f"Найдено комплектаций: {len(config_names)}")
    for cn in config_names[:5]:
        qty_count = len(config_quantities[cn])
        logger.info(f"  {cn}: {qty_count} деталей")
    if len(config_names) > 5:
        logger.info(f"  ... и ещё {len(config_names) - 5} комплектаций")

    return BOMData(
        parts=parts,
        config_names=config_names,
        config_quantities=config_quantities,
        source_file=file_path,
    )


def get_config_quantities(bom: BOMData, config_name: str) -> Dict[str, PartInfo]:
    """Получить данные деталей для выбранной комплектации.

    Args:
        bom: Распарсенные BOM-данные.
        config_name: Название комплектации.

    Returns:
        Словарь {part_number: PartInfo} с заполненным quantity для комплектации.
    """
    if config_name not in bom.config_quantities:
        raise ValueError(f"Комплектация '{config_name}' не найдена. "
                         f"Доступные: {bom.config_names[:10]}...")

    result: Dict[str, PartInfo] = {}
    for part_no, qty in bom.config_quantities[config_name].items():
        if part_no in bom.parts:
            part = PartInfo(
                part_number=part_no,
                name_cn=bom.parts[part_no].name_cn,
                name_en=bom.parts[part_no].name_en,
                quantity=qty,
            )
            result[part_no] = part

    return result
