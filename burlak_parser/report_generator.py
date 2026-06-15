"""Модуль генерации выходных артефактов.

По итогам работы система отдаёт пользователю 2 артефакта:
  1. Excel-файл «Отчёт о расхождениях» с колонками:
     Каталожный номер | Название (кит/англ) | Кол-во по BOM | Кол-во по Картам |
     Номера операционных карт | Тип ошибки.
  2. ZIP-архив с очищенными и разделёнными (по одному листу) .xlsx файлами
     операционных карт.
"""

from __future__ import annotations

import logging
import os
import zipfile
from typing import Dict, List, Optional

import xlsxwriter

from burlak_parser.bom_parser import PartInfo
from burlak_parser.comparator import ComparisonResult, DiscrepancyType

logger = logging.getLogger(__name__)


def generate_discrepancy_report(comparison: ComparisonResult,
                                 output_path: str,
                                 bom_parts: Optional[Dict[str, PartInfo]] = None) -> str:
    """Сгенерировать Excel-отчёт о расхождениях.

    Args:
        comparison: Результат сверки.
        output_path: Путь для сохранения .xlsx файла.
        bom_parts: Детали BOM для записи на лист "Все детали BOM".

    Returns:
        Путь к созданному файлу.
    """
    workbook = xlsxwriter.Workbook(output_path)

    # Форматы
    header_fmt = workbook.add_format({
        'bold': True,
        'bg_color': '#4472C4',
        'font_color': 'white',
        'border': 1,
        'text_wrap': True,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11,
    })
    cell_fmt = workbook.add_format({
        'border': 1,
        'text_wrap': True,
        'valign': 'vcenter',
        'font_size': 10,
    })
    cell_center_fmt = workbook.add_format({
        'border': 1,
        'text_wrap': True,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 10,
    })
    cell_num_fmt = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'num_format': '0.00',
        'font_size': 10,
    })
    bom_only_fmt = workbook.add_format({
        'border': 1,
        'bg_color': '#FFF2CC',  # Жёлтый
        'text_wrap': True,
        'valign': 'vcenter',
        'font_size': 10,
    })
    cards_only_fmt = workbook.add_format({
        'border': 1,
        'bg_color': '#D9E2F3',  # Голубой
        'text_wrap': True,
        'valign': 'vcenter',
        'font_size': 10,
    })
    qty_mismatch_fmt = workbook.add_format({
        'border': 1,
        'bg_color': '#FCE4EC',  # Розовый
        'text_wrap': True,
        'valign': 'vcenter',
        'font_size': 10,
    })
    fuzzy_fmt = workbook.add_format({
        'border': 1,
        'bg_color': '#E8F5E9',  # Зелёный (найдено предполагаемое совпадение)
        'text_wrap': True,
        'valign': 'vcenter',
        'font_size': 10,
    })
    bom_part_fmt = workbook.add_format({
        'border': 1,
        'text_wrap': True,
        'valign': 'vcenter',
        'font_size': 10,
    })

    # === Лист 1: Сводка ===
    ws_summary = workbook.add_worksheet('Сводка')
    ws_summary.set_tab_color('#4472C4')
    ws_summary.set_column('A:A', 25)
    ws_summary.set_column('B:B', 20)
    ws_summary.set_column('C:C', 50)

    title_fmt = workbook.add_format({
        'bold': True,
        'font_size': 14,
        'font_color': '#1F3864',
    })
    label_fmt = workbook.add_format({
        'bold': True,
        'font_size': 11,
    })
    value_fmt = workbook.add_format({
        'font_size': 11,
    })

    ws_summary.merge_range('A1:C1', 'Отчёт о расхождениях BOM и операционных карт', title_fmt)
    ws_summary.write('A3', 'Параметр', label_fmt)
    ws_summary.write('B3', 'Значение', label_fmt)

    summary_data = [
        ('Комплектация', comparison.bom_config_name),
        ('Деталей в BOM', str(comparison.total_bom_parts)),
        ('Деталей в картах', str(comparison.total_cards_parts)),
        ('Совпало', str(comparison.matched_parts)),
        ('Всего расхождений', str(len(comparison.discrepancies))),
        ('  Только в BOM', str(sum(1 for d in comparison.discrepancies
                                    if d.discrepancy_type == DiscrepancyType.ONLY_IN_BOM))),
        ('  Только в картах', str(sum(1 for d in comparison.discrepancies
                                        if d.discrepancy_type == DiscrepancyType.ONLY_IN_CARDS))),
        ('  Конфликт количества', str(sum(1 for d in comparison.discrepancies
                                            if d.discrepancy_type == DiscrepancyType.QUANTITY_MISMATCH))),
    ]
    for i, (label, value) in enumerate(summary_data, 4):
        ws_summary.write(i, 0, label, label_fmt)
        ws_summary.write(i, 1, value, value_fmt)

    # === Лист 2: Расхождения ===
    ws = workbook.add_worksheet('Расхождения')
    ws.set_tab_color('#C00000')

    # Колонки
    headers = [
        'Каталожный номер',
        'Название (кит.)',
        'Название (англ.)',
        'Кол-во по BOM',
        'Кол-во по картам',
        'Номера операционных карт',
        'Fuzzy-совпадение',
        'Тип ошибки',
    ]
    col_widths = [22, 30, 30, 14, 14, 45, 22, 22]

    for col_idx, (header, width) in enumerate(zip(headers, col_widths)):
        ws.set_column(col_idx, col_idx, width)
        ws.write(0, col_idx, header, header_fmt)

    # Устанавливаем высоту заголовка
    ws.set_row(0, 30)

    # Данные
    for row_idx, disc in enumerate(comparison.discrepancies, 1):
        # Выбираем формат в зависимости от типа ошибки
        if disc.discrepancy_type in (DiscrepancyType.FUZZY_IN_BOM, DiscrepancyType.FUZZY_IN_CARDS):
            fmt = fuzzy_fmt
        elif disc.discrepancy_type == DiscrepancyType.ONLY_IN_BOM:
            fmt = bom_only_fmt
        elif disc.discrepancy_type == DiscrepancyType.ONLY_IN_CARDS:
            fmt = cards_only_fmt
        else:
            fmt = qty_mismatch_fmt

        ws.write(row_idx, 0, disc.part_number, fmt)
        ws.write(row_idx, 1, disc.name_cn, fmt)
        ws.write(row_idx, 2, disc.name_en, fmt)
        ws.write(row_idx, 3, disc.qty_bom, cell_num_fmt)
        ws.write(row_idx, 4, disc.qty_cards, cell_num_fmt)
        ws.write(row_idx, 5, ', '.join(disc.card_numbers), fmt)
        ws.write(row_idx, 6, disc.fuzzy_match, fmt)
        ws.write(row_idx, 7, disc.discrepancy_type, fmt)

    # === Лист 3: Все детали BOM ===
    ws_bom = workbook.add_worksheet('Все детали BOM')
    ws_bom.set_tab_color('#548235')

    bom_headers = [
        'Каталожный номер',
        'Название (кит.)',
        'Название (англ.)',
        'Количество',
    ]
    bom_widths = [22, 30, 30, 14]
    for col_idx, (header, width) in enumerate(zip(bom_headers, bom_widths)):
        ws_bom.set_column(col_idx, col_idx, width)
        ws_bom.write(0, col_idx, header, header_fmt)

    # Записываем данные деталей BOM, если они переданы
    if bom_parts:
        for row_idx, (part_no, part) in enumerate(sorted(bom_parts.items()), 1):
            ws_bom.write(row_idx, 0, part_no, bom_part_fmt)
            ws_bom.write(row_idx, 1, part.name_cn, bom_part_fmt)
            ws_bom.write(row_idx, 2, part.name_en, bom_part_fmt)
            ws_bom.write(row_idx, 3, part.quantity, cell_num_fmt)

    workbook.close()

    logger.info(f"Отчёт сохранён: {output_path}")
    return output_path


def create_split_cards_archive(split_files_dir: str, output_path: str) -> str:
    """Создать ZIP-архив с разделёнными операционными картами.

    Args:
        split_files_dir: Директория с разделёнными файлами карт.
        output_path: Путь для сохранения ZIP-файла.

    Returns:
        Путь к созданному ZIP-архиву.
    """
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(split_files_dir):
            for fn in files:
                file_path = os.path.join(root, fn)
                arcname = os.path.relpath(file_path, split_files_dir)
                zf.write(file_path, arcname)

    logger.info(f"ZIP-архив создан: {output_path} ({os.path.getsize(output_path) / 1024:.1f} KB)")
    return output_path
