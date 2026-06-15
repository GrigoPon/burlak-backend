"""Модуль генерации выходных артефактов.

По итогам работы система отдаёт:
  1. Excel-файл «discrepancies.xlsx» — 5 листов Enterprise-уровня:
     - «Сводка»: общая статистика + таблица по комплектациям.
     - «Расхождения»: детальный список всех несоответствий с автофильтром.
     - «Неточное совпадение номеров»: fuzzy matches.
     - «Все детали BOM»: полный перечень деталей спецификации.
     - «Ошибки файлов»: повреждённые файлы (если есть).
  2. ZIP-архив с разделёнными .xlsx файлами операционных карт.

Класс Reporter — обёртка для использования в FastAPI/серверной архитектуре.
"""

from __future__ import annotations

import logging
import os
import zipfile
from typing import Dict, List, Optional

import xlsxwriter

from burlak_parser.bom_parser import BOMData
from burlak_parser.comparator import (
    ConfigComparisonResult,
    Discrepancy,
    DiscrepancyType,
    MultiConfigComparisonResult,
)
from burlak_parser.card_parser import CardsData

logger = logging.getLogger(__name__)

MAX_CONFIGS_IN_MATRIX = 50


def generate_discrepancy_report(
    result: MultiConfigComparisonResult,
    output_path: str,
    bom: Optional[BOMData] = None,
    cards_data: Optional[CardsData] = None,
) -> str:
    """Сгенерировать Excel-отчёт Enterprise-уровня для ВСЕХ комплектаций.

    Структура:
      1. Сводка — общая статистика и таблица по комплектациям.
      2. Расхождения — полный список с автофильтром и цветовой индикацией.
      3. Неточное совпадение номеров — fuzzy matches.
      4. Все детали BOM — полный перечень деталей.
      5. Ошибки файлов — повреждённые файлы (если есть).
    """
    workbook = xlsxwriter.Workbook(output_path)

    # ── Общие форматы ──
    header_fmt = workbook.add_format({
        'bold': True, 'bg_color': '#4472C4', 'font_color': 'white',
        'border': 1, 'text_wrap': True, 'align': 'center',
        'valign': 'vcenter', 'font_size': 11,
    })
    title_fmt = workbook.add_format({
        'bold': True, 'font_size': 14, 'font_color': '#1F3864',
    })
    label_fmt = workbook.add_format({
        'bold': True, 'font_size': 11,
    })
    value_fmt = workbook.add_format({
        'font_size': 11,
    })
    cell_fmt = workbook.add_format({
        'border': 1, 'text_wrap': True, 'valign': 'vcenter', 'font_size': 10,
    })
    cell_center_fmt = workbook.add_format({
        'border': 1, 'text_wrap': True, 'align': 'center',
        'valign': 'vcenter', 'font_size': 10,
    })
    cell_num_fmt = workbook.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter',
        'num_format': '0.00', 'font_size': 10,
    })
    # Цветовые форматы по типам несоответствий
    qty_mismatch_fmt = workbook.add_format({
        'border': 1, 'bg_color': '#FCE4EC', 'text_wrap': True,
        'valign': 'vcenter', 'font_size': 10,
    })
    bom_only_fmt = workbook.add_format({
        'border': 1, 'bg_color': '#FFF2CC', 'text_wrap': True,
        'valign': 'vcenter', 'font_size': 10,
    })
    cards_only_fmt = workbook.add_format({
        'border': 1, 'bg_color': '#D9E2F3', 'text_wrap': True,
        'valign': 'vcenter', 'font_size': 10,
    })
    fuzzy_fmt = workbook.add_format({
        'border': 1, 'bg_color': '#E2EFDA', 'text_wrap': True,
        'valign': 'vcenter', 'font_size': 10,
    })

    # ══════════════════════════════════════════════════════════════════════════
    # Лист 1: СВОДКА
    # ══════════════════════════════════════════════════════════════════════════
    ws_summary = workbook.add_worksheet('Сводка')
    ws_summary.set_tab_color('#4472C4')
    ws_summary.hide_gridlines(2)
    ws_summary.set_column('A:A', 40)
    ws_summary.set_column('B:B', 18)

    # Заголовок
    ws_summary.merge_range('A1:B1', 'ОТЧЁТ ПРОВЕРКИ КОМПЛЕКТАЦИЙ', title_fmt)
    ws_summary.set_row(1, 24)

    # Статистика
    total_bom_only = sum(1 for d in result.all_discrepancies if d.discrepancy_type == DiscrepancyType.ONLY_IN_BOM)
    total_cards_only = sum(1 for d in result.all_discrepancies if d.discrepancy_type == DiscrepancyType.ONLY_IN_CARDS)
    total_qty_mismatch = sum(1 for d in result.all_discrepancies if d.discrepancy_type == DiscrepancyType.QUANTITY_MISMATCH)
    total_fuzzy = sum(1 for d in result.all_discrepancies if d.discrepancy_type == DiscrepancyType.FUZZY_MATCH)

    stats = [
        ('Проверено комплектаций', str(result.total_configs)),
        ('Деталей в спецификации', str(result.total_bom_unique_parts)),
        ('Деталей в инструкциях', str(result.total_cards_unique_parts)),
        ('', ''),
        ('ВСЕГО НЕСООТВЕТСТВИЙ', str(len(result.all_discrepancies))),
        ('  Разное количество', str(total_qty_mismatch)),
        ('  Есть в спецификации, нет в инструкциях', str(total_bom_only)),
        ('  Есть в инструкциях, нет в спецификации', str(total_cards_only)),
    ]
    if total_fuzzy:
        stats.append(('  Разный формат номера', str(total_fuzzy)))
    if cards_data:
        corrupted_count = len(cards_data.corrupted_files) if cards_data.corrupted_files else 0
        stats.append(('', ''))
        stats.append(('Обработано файлов инструкций', str(cards_data.total_cards_processed)))
        stats.append(('  Служебных файлов пропущено', str(cards_data.service_files_skipped)))
        if corrupted_count:
            stats.append(('  Повреждённых файлов', str(corrupted_count)))

    row = 3
    for label, value in stats:
        if label.startswith('  '):
            ws_summary.write(row, 0, '  ' + label.strip(), label_fmt)
        elif label == '':
            row += 1
            continue
        else:
            ws_summary.write(row, 0, label, label_fmt)
        ws_summary.write(row, 1, value, value_fmt)
        row += 1

    # Таблица по комплектациям
    gap_row = row + 1
    ws_summary.merge_range(gap_row, 0, gap_row, 7,
                           'СВОДКА ПО КОМПЛЕКТАЦИЯМ', title_fmt)
    ws_summary.set_row(gap_row, 22)

    config_header_row = gap_row + 1
    config_headers = ['Комплектация', 'Деталей в спец.', 'Деталей в инстр.',
                      'Совпало', 'Несоотв.', 'Только в спец.', 'Только в инстр.', 'Разное кол-во']
    config_widths = [55, 14, 14, 10, 10, 14, 14, 14]
    for ci, w in enumerate(config_widths):
        ws_summary.set_column(ci, ci, w)
    for ci, h in enumerate(config_headers):
        ws_summary.write(config_header_row, ci, h, header_fmt)

    for ri, cr in enumerate(result.config_results, config_header_row + 1):
        short = cr.config_name if len(cr.config_name) <= 52 else cr.config_name[:49] + "..."
        ws_summary.write(ri, 0, short, cell_fmt)
        ws_summary.write(ri, 1, cr.total_bom_parts, cell_center_fmt)
        ws_summary.write(ri, 2, cr.total_cards_parts, cell_center_fmt)
        ws_summary.write(ri, 3, cr.matched_parts, cell_center_fmt)
        ws_summary.write(ri, 4, len(cr.discrepancies), cell_center_fmt)
        ws_summary.write(ri, 5, sum(1 for d in cr.discrepancies if d.discrepancy_type == DiscrepancyType.ONLY_IN_BOM), cell_center_fmt)
        ws_summary.write(ri, 6, sum(1 for d in cr.discrepancies if d.discrepancy_type == DiscrepancyType.ONLY_IN_CARDS), cell_center_fmt)
        ws_summary.write(ri, 7, sum(1 for d in cr.discrepancies if d.discrepancy_type == DiscrepancyType.QUANTITY_MISMATCH), cell_center_fmt)

    # Заморозка и автофильтр
    ws_summary.freeze_panes(config_header_row + 1, 0)
    ws_summary.autofilter(config_header_row, 0, config_header_row + len(result.config_results), 7)

    # ══════════════════════════════════════════════════════════════════════════
    # Лист 2: РАСХОЖДЕНИЯ (основной)
    # ══════════════════════════════════════════════════════════════════════════
    ws = workbook.add_worksheet('Расхождения')
    ws.set_tab_color('#C00000')
    ws.freeze_panes(1, 0)

    disc_headers = [
        'Каталожный номер', 'Название (кит.)', 'Название (англ.)',
        'Комплектация', 'Кол-во в спецификации', 'Кол-во в инструкциях',
        'Номера инструкций', 'Тип несоответствия',
    ]
    disc_widths = [22, 30, 30, 35, 14, 14, 45, 30]

    for ci, (h, w) in enumerate(zip(disc_headers, disc_widths)):
        ws.set_column(ci, ci, w)
        ws.write(0, ci, h, header_fmt)
    ws.set_row(0, 30)

    # Автофильтр на весь диапазон
    if result.all_discrepancies:
        ws.autofilter(0, 0, len(result.all_discrepancies), 7)

    for ri, disc in enumerate(result.all_discrepancies, 1):
        if disc.discrepancy_type == DiscrepancyType.QUANTITY_MISMATCH:
            fmt = qty_mismatch_fmt
        elif disc.discrepancy_type == DiscrepancyType.ONLY_IN_BOM:
            fmt = bom_only_fmt
        elif disc.discrepancy_type == DiscrepancyType.ONLY_IN_CARDS:
            fmt = cards_only_fmt
        elif disc.discrepancy_type == DiscrepancyType.FUZZY_MATCH:
            fmt = fuzzy_fmt
        else:
            fmt = cell_fmt

        ws.write(ri, 0, disc.part_number, fmt)
        ws.write(ri, 1, disc.name_cn, fmt)
        ws.write(ri, 2, disc.name_en, fmt)
        ws.write(ri, 3, disc.config_name[:70] if disc.config_name else "", fmt)
        ws.write(ri, 4, disc.qty_bom, cell_num_fmt)
        ws.write(ri, 5, disc.qty_cards, cell_num_fmt)
        ws.write(ri, 6, ', '.join(disc.card_numbers[:5]), fmt)
        ws.write(ri, 7, disc.discrepancy_type, fmt)

    # ══════════════════════════════════════════════════════════════════════════
    # Лист 3: НЕТОЧНОЕ СОВПАДЕНИЕ НОМЕРОВ
    # ══════════════════════════════════════════════════════════════════════════
    fuzzy_discs = [d for d in result.all_discrepancies if d.discrepancy_type == DiscrepancyType.FUZZY_MATCH]
    if fuzzy_discs:
        ws_fuzzy = workbook.add_worksheet('Неточное совпадение номеров')
        ws_fuzzy.set_tab_color('#548235')
        ws_fuzzy.freeze_panes(1, 0)

        fuzzy_headers = ['Номер в инструкциях', 'Номер в спецификации',
                         'Кол-во в спец.', 'Кол-во в инстр.', 'Комплектация']
        fuzzy_widths = [25, 25, 14, 14, 40]
        for ci, (h, w) in enumerate(zip(fuzzy_headers, fuzzy_widths)):
            ws_fuzzy.set_column(ci, ci, w)
            ws_fuzzy.write(0, ci, h, header_fmt)

        ws_fuzzy.autofilter(0, 0, len(fuzzy_discs), 4)
        for ri, disc in enumerate(fuzzy_discs, 1):
            ws_fuzzy.write(ri, 0, disc.part_number, fuzzy_fmt)
            ws_fuzzy.write(ri, 1, disc.fuzzy_matched_to, fuzzy_fmt)
            ws_fuzzy.write(ri, 2, disc.qty_bom, cell_num_fmt)
            ws_fuzzy.write(ri, 3, disc.qty_cards, cell_num_fmt)
            ws_fuzzy.write(ri, 4, disc.config_name[:70] if disc.config_name else "", fuzzy_fmt)

    # ══════════════════════════════════════════════════════════════════════════
    # Лист 4: ВСЕ ДЕТАЛИ BOM
    # ══════════════════════════════════════════════════════════════════════════
    if bom and bom.parts:
        ws_bom = workbook.add_worksheet('Все детали BOM')
        ws_bom.set_tab_color('#2F5496')
        ws_bom.freeze_panes(1, 0)

        bom_headers = ['Каталожный номер', 'Название (кит.)', 'Название (англ.)']
        bom_widths = [22, 35, 35]
        if bom.config_names:
            # Показываем количества по каждой комплектации
            for cn in bom.config_names[:MAX_CONFIGS_IN_MATRIX]:
                short = cn if len(cn) <= 25 else cn[:22] + "..."
                bom_headers.append(short)
                bom_widths.append(10)

        for ci, (h, w) in enumerate(zip(bom_headers, bom_widths)):
            ws_bom.set_column(ci, ci, w)
            ws_bom.write(0, ci, h, header_fmt)
        ws_bom.set_row(0, 30)

        sorted_parts = sorted(bom.parts.items())
        ws_bom.autofilter(0, 0, len(sorted_parts), len(bom_headers) - 1)

        for ri, (pn, part) in enumerate(sorted_parts, 1):
            ws_bom.write(ri, 0, pn, cell_fmt)
            ws_bom.write(ri, 1, part.name_cn, cell_fmt)
            ws_bom.write(ri, 2, part.name_en, cell_fmt)
            for ci, cn in enumerate(bom.config_names[:MAX_CONFIGS_IN_MATRIX], 3):
                qty = bom.config_quantities[cn].get(pn, 0.0)
                ws_bom.write(ri, ci, qty if qty > 0 else "", cell_num_fmt)

    # ══════════════════════════════════════════════════════════════════════════
    # Лист 5: ОШИБКИ ФАЙЛОВ
    # ══════════════════════════════════════════════════════════════════════════
    if cards_data and cards_data.corrupted_files:
        corrupted = cards_data.corrupted_files
        ws_corrupt = workbook.add_worksheet('Ошибки файлов')
        ws_corrupt.set_tab_color('#C00000')
        ws_corrupt.freeze_panes(1, 0)
        ws_corrupt.set_column(0, 0, 90)
        ws_corrupt.write(0, 0, 'Путь к повреждённому файлу', header_fmt)
        ws_corrupt.set_row(0, 30)
        if corrupted:
            ws_corrupt.autofilter(0, 0, len(corrupted), 0)
        for ri, fpath in enumerate(corrupted, 1):
            ws_corrupt.write(ri, 0, fpath, cell_fmt)

    workbook.close()
    logger.info("Отчёт сохранён: %s", output_path)
    return output_path


def generate_legacy_report(
    comparison,
    output_path: str,
    bom_parts: Optional[Dict[str, "PartInfo"]] = None,
) -> str:
    """Сгенерировать Excel-отчёт для одной комплектации (старый формат)."""
    from burlak_parser.bom_parser import PartInfo

    workbook = xlsxwriter.Workbook(output_path)

    header_fmt = workbook.add_format({
        'bold': True, 'bg_color': '#4472C4', 'font_color': 'white',
        'border': 1, 'text_wrap': True, 'align': 'center',
        'valign': 'vcenter', 'font_size': 11,
    })
    bom_only_fmt = workbook.add_format({
        'border': 1, 'bg_color': '#FFF2CC', 'text_wrap': True,
        'valign': 'vcenter', 'font_size': 10,
    })
    cards_only_fmt = workbook.add_format({
        'border': 1, 'bg_color': '#D9E2F3', 'text_wrap': True,
        'valign': 'vcenter', 'font_size': 10,
    })
    qty_mismatch_fmt = workbook.add_format({
        'border': 1, 'bg_color': '#FCE4EC', 'text_wrap': True,
        'valign': 'vcenter', 'font_size': 10,
    })
    cell_num_fmt = workbook.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter',
        'num_format': '0.00', 'font_size': 10,
    })
    cell_fmt = workbook.add_format({
        'border': 1, 'text_wrap': True, 'valign': 'vcenter', 'font_size': 10,
    })

    ws = workbook.add_worksheet('Расхождения')
    ws.freeze_panes(1, 0)
    headers = ['Каталожный номер', 'Название (кит.)', 'Название (англ.)',
               'Кол-во по BOM', 'Кол-во по картам', 'Номера операционных карт', 'Тип ошибки']
    widths = [22, 30, 30, 14, 14, 45, 20]
    for ci, (h, w) in enumerate(zip(headers, widths)):
        ws.set_column(ci, ci, w)
        ws.write(0, ci, h, header_fmt)
    ws.set_row(0, 30)

    for ri, disc in enumerate(comparison.discrepancies, 1):
        if disc.discrepancy_type == DiscrepancyType.ONLY_IN_BOM:
            fmt = bom_only_fmt
        elif disc.discrepancy_type == DiscrepancyType.ONLY_IN_CARDS:
            fmt = cards_only_fmt
        else:
            fmt = qty_mismatch_fmt
        ws.write(ri, 0, disc.part_number, fmt)
        ws.write(ri, 1, disc.name_cn, fmt)
        ws.write(ri, 2, disc.name_en, fmt)
        ws.write(ri, 3, disc.qty_bom, cell_num_fmt)
        ws.write(ri, 4, disc.qty_cards, cell_num_fmt)
        ws.write(ri, 5, ', '.join(disc.card_numbers), fmt)
        ws.write(ri, 6, disc.discrepancy_type, fmt)

    if bom_parts:
        ws_bom = workbook.add_worksheet('Все детали BOM')
        ws_bom.freeze_panes(1, 0)
        bom_headers = ['Каталожный номер', 'Название (кит.)', 'Название (англ.)', 'Количество']
        bom_widths = [22, 30, 30, 14]
        for ci, (h, w) in enumerate(zip(bom_headers, bom_widths)):
            ws_bom.set_column(ci, ci, w)
            ws_bom.write(0, ci, h, header_fmt)
        for ri, (pn, part) in enumerate(sorted(bom_parts.items()), 1):
            ws_bom.write(ri, 0, pn, cell_fmt)
            ws_bom.write(ri, 1, part.name_cn, cell_fmt)
            ws_bom.write(ri, 2, part.name_en, cell_fmt)
            ws_bom.write(ri, 3, part.quantity, cell_num_fmt)

    workbook.close()
    return output_path


def create_split_cards_archive(split_files_dir: str, output_path: str) -> str:
    """Создать ZIP-архив с разделёнными операционными картами."""
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(split_files_dir):
            for fn in files:
                if fn.startswith("~$"):
                    continue
                file_path = os.path.join(root, fn)
                if not os.path.exists(file_path):
                    continue
                try:
                    arcname = os.path.relpath(file_path, split_files_dir)
                    zf.write(file_path, arcname)
                except (FileNotFoundError, PermissionError):
                    logger.debug("Пропуск недоступного файла: %s", fn)

    size_kb = os.path.getsize(output_path) / 1024 if os.path.exists(output_path) else 0
    logger.info("ZIP-архив создан: %s (%.1f KB)", output_path, size_kb)
    return output_path


class Reporter:
    """Сервис генерации отчётов Enterprise-уровня."""

    def generate(
        self,
        result: MultiConfigComparisonResult,
        output_dir: str,
        bom: Optional[BOMData] = None,
        cards_data: Optional[CardsData] = None,
    ) -> Dict[str, str]:
        """Сгенерировать все отчёты.

        Returns:
            Словарь {описание: путь_к_файлу}.
        """
        os.makedirs(output_dir, exist_ok=True)
        outputs: Dict[str, str] = {}

        # Excel-отчёт
        excel_path = os.path.join(output_dir, "discrepancies.xlsx")
        generate_discrepancy_report(result, excel_path, bom=bom, cards_data=cards_data)
        outputs["excel_report"] = excel_path

        # Текстовый отчёт
        from burlak_parser.comparator import format_discrepancy_report
        text = format_discrepancy_report(result)
        txt_path = os.path.join(output_dir, "report.txt")
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(text)
        outputs["text_report"] = txt_path

        return outputs
