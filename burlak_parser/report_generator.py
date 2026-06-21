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
    Discrepancy,
    DiscrepancyType,
    MultiConfigComparisonResult,
)
from burlak_parser.card_parser import CardsData

logger = logging.getLogger(__name__)

# Без лимита — показываем ВСЕ комплектации в матрице
# (T1L BOM: 78, может быть больше)

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
    # Лист 1: СВОДКА — профессиональный дашборд
    # ══════════════════════════════════════════════════════════════════════════
    ws_summary = workbook.add_worksheet('Сводка')
    ws_summary.set_tab_color('#1F3864')
    ws_summary.hide_gridlines(2)

    # Заголовок-шапка
    title_fmt_big = workbook.add_format({
        'bold': True, 'font_size': 18, 'font_color': '#1F3864',
        'align': 'center', 'valign': 'vcenter',
    })
    subtitle_fmt = workbook.add_format({
        'font_size': 10, 'font_color': '#666666',
        'align': 'center', 'valign': 'vcenter',
    })
    # Карточки с метриками
    card_title_fmt = workbook.add_format({
        'bold': True, 'font_size': 11, 'font_color': '#FFFFFF',
        'bg_color': '#2F5496', 'align': 'center', 'valign': 'vcenter',
        'border': 0,
    })
    card_value_fmt = workbook.add_format({
        'bold': True, 'font_size': 24, 'font_color': '#1F3864',
        'bg_color': '#D6E4F0', 'align': 'center', 'valign': 'vcenter',
        'border': 0,
    })
    card_sub_fmt = workbook.add_format({
        'font_size': 9, 'font_color': '#595959',
        'bg_color': '#D6E4F0', 'align': 'center', 'valign': 'vcenter',
        'border': 0,
    })

    # Ширины колонок для дашборда
    for c in range(8):
        ws_summary.set_column(c, c, 22)

    # Заголовок
    ws_summary.merge_range('A1:H1', 'ОТЧЁТ СВЕРКИ BOM И ОПЕРАЦИОННЫХ КАРТ', title_fmt_big)
    ws_summary.set_row(0, 30)
    ws_summary.merge_range('A2:H2', f'Проверено {result.total_configs} комплектаций | Деталей в BOM: {result.total_bom_unique_parts} | В картах: {result.total_cards_unique_parts}', subtitle_fmt)
    ws_summary.set_row(1, 18)

    # Карточки метрик (строка 3-5)
    total = len(result.all_discrepancies)
    total_qty_mismatch = sum(1 for d in result.all_discrepancies if d.discrepancy_type == DiscrepancyType.QUANTITY_MISMATCH)
    total_bom_only = sum(1 for d in result.all_discrepancies if d.discrepancy_type == DiscrepancyType.ONLY_IN_BOM)
    total_cards_only = sum(1 for d in result.all_discrepancies if d.discrepancy_type == DiscrepancyType.ONLY_IN_CARDS)

    # Карточка 1: Всего несоответствий
    ws_summary.merge_range('A3:B3', 'ВСЕГО НЕСООТВЕТСТВИЙ', card_title_fmt)
    ws_summary.merge_range('A4:B4', str(total), card_value_fmt)
    ws_summary.merge_range('A5:B5', 'по всем комплектациям', card_sub_fmt)
    ws_summary.set_row(2, 22)
    ws_summary.set_row(3, 40)
    ws_summary.set_row(4, 16)

    # Карточка 2: Разное количество
    qty_color = '#C00000' if total_qty_mismatch > 0 else '#548235'
    card_title_qty = workbook.add_format({
        'bold': True, 'font_size': 11, 'font_color': '#FFFFFF',
        'bg_color': qty_color, 'align': 'center', 'valign': 'vcenter', 'border': 0,
    })
    card_value_qty = workbook.add_format({
        'bold': True, 'font_size': 24, 'font_color': '#1F3864',
        'bg_color': '#FCE4EC', 'align': 'center', 'valign': 'vcenter', 'border': 0,
    })
    card_sub_qty = workbook.add_format({
        'font_size': 9, 'font_color': '#595959',
        'bg_color': '#FCE4EC', 'align': 'center', 'valign': 'vcenter', 'border': 0,
    })
    ws_summary.merge_range('C3:D3', 'РАЗНОЕ КОЛИЧЕСТВО', card_title_qty)
    ws_summary.merge_range('C4:D4', str(total_qty_mismatch), card_value_qty)
    ws_summary.merge_range('C5:D5', 'конфликтов количества', card_sub_qty)

    # Карточка 3: Есть в BOM, нет в картах
    ws_summary.merge_range('E3:F3', 'В BOM, НЕТ В КАРТАХ', card_title_fmt)
    ws_summary.merge_range('E4:F4', str(total_bom_only), card_value_fmt)
    ws_summary.merge_range('E5:F5', 'отсутствуют в картах', card_sub_fmt)

    # Карточка 4: Есть в картах, нет в BOM
    ws_summary.merge_range('G3:H3', 'В КАРТАХ, НЕТ В BOM', card_title_fmt)
    ws_summary.merge_range('G4:H4', str(total_cards_only), card_value_fmt)
    ws_summary.merge_range('G5:H5', 'отсутствуют в BOM', card_sub_fmt)

    # Строка с информацией о файлах
    info_row = 6
    if cards_data:
        corrupted_count = len(cards_data.corrupted_files) if cards_data.corrupted_files else 0
        info_text = f'Обработано файлов карт: {cards_data.total_cards_processed}  |  Служебных пропущено: {cards_data.service_files_skipped}'
        if corrupted_count:
            info_text += f'  |  Повреждённых: {corrupted_count}'
        ws_summary.merge_range(info_row, 0, info_row, 7, info_text, subtitle_fmt)
        info_row += 1

    # Таблица по комплектациям
    table_title_fmt = workbook.add_format({
        'bold': True, 'font_size': 13, 'font_color': '#1F3864',
        'align': 'left', 'valign': 'vcenter',
    })
    ws_summary.merge_range(info_row + 1, 0, info_row + 1, 7, 'СВОДКА ПО КОМПЛЕКТАЦИЯМ', table_title_fmt)
    ws_summary.set_row(info_row + 1, 22)

    config_header_row = info_row + 2
    config_headers = ['Комплектация', 'Деталей в BOM', 'Деталей в картах',
                      'Совпало', 'Несоотв.', 'Только в BOM', 'Только в картах', 'Разное кол-во']
    for ci, h in enumerate(config_headers):
        ws_summary.write(config_header_row, ci, h, header_fmt)
    ws_summary.set_row(config_header_row, 30)

    # Чередование строк
    alt_row_fmt = workbook.add_format({
        'border': 1, 'text_wrap': True, 'valign': 'vcenter', 'font_size': 10,
        'bg_color': '#F2F2F2',
    })

    for ri, cr in enumerate(result.config_results, config_header_row + 1):
        short = cr.config_name if len(cr.config_name) <= 52 else cr.config_name[:49] + "..."
        fmt = alt_row_fmt if (ri - config_header_row) % 2 == 0 else cell_fmt
        ws_summary.write(ri, 0, short, fmt)
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
        'Комплектация', 'Кол-во в BOM', 'Кол-во в картах',
        'Номера операционных карт', 'Тип несоответствия',
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

        fuzzy_headers = ['Номер в картах', 'Номер в BOM',
                         'Кол-во в BOM', 'Кол-во в картах', 'Комплектация']
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
            for cn in bom.config_names:
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
            # Используем оригинальный формат номера из BOM (с тире и т.д.)
            original_no = part.part_number if part.part_number else pn
            ws_bom.write(ri, 0, original_no, cell_fmt)
            ws_bom.write(ri, 1, part.name_cn, cell_fmt)
            ws_bom.write(ri, 2, part.name_en, cell_fmt)
            for ci, cn in enumerate(bom.config_names, 3):
                qty = bom.config_quantities[cn].get(pn, 0.0)
                ws_bom.write(ri, ci, qty if qty > 0 else "", cell_num_fmt)

    # ══════════════════════════════════════════════════════════════════════════
    # Лист 5: ОШИБКИ ФАЙЛОВ / ПОВРЕЖДЁННЫЕ ФАЙЛЫ
    # ══════════════════════════════════════════════════════════════════════════
    # Объединяем все повреждённые файлы (из парсинга + из разделения)
    all_corrupted_detailed: list = []
    if cards_data:
        # Parse-phase errors (detailed)
        parse_corrupted = getattr(cards_data, 'corrupted_files_detailed', None)
        if parse_corrupted:
            all_corrupted_detailed.extend(parse_corrupted)
        # Legacy corrupted_files (simple list of paths, no details)
        if cards_data.corrupted_files:
            known_files = {e.get('file', '') for e in all_corrupted_detailed}
            for fpath in cards_data.corrupted_files:
                fname = os.path.basename(fpath)
                if fname not in known_files:
                    all_corrupted_detailed.append({
                        'file': fname,
                        'folder': os.path.dirname(fpath),
                        'error': '',
                        'phase': 'split',
                    })

    if all_corrupted_detailed:
        ws_corrupt = workbook.add_worksheet('Поврежденные файлы')
        ws_corrupt.set_tab_color('#C00000')
        ws_corrupt.freeze_panes(1, 0)
        ws_corrupt.set_column(0, 0, 50)
        ws_corrupt.set_column(1, 1, 50)
        ws_corrupt.set_column(2, 2, 70)
        ws_corrupt.set_column(3, 3, 12)
        headers_corrupt = ['Имя файла', 'Расположение', 'Описание ошибки', 'Фаза']
        for ci, h in enumerate(headers_corrupt):
            ws_corrupt.write(0, ci, h, header_fmt)
        ws_corrupt.set_row(0, 30)
        ws_corrupt.autofilter(0, 0, len(all_corrupted_detailed), 3)
        for ri, entry in enumerate(all_corrupted_detailed, 1):
            ws_corrupt.write(ri, 0, entry.get('file', entry.get('file_name', '')), cell_fmt)
            ws_corrupt.write(ri, 1, entry.get('folder', ''), cell_fmt)
            ws_corrupt.write(ri, 2, entry.get('error', ''), cell_fmt)
            ws_corrupt.write(ri, 3, entry.get('phase', ''), cell_fmt)

    workbook.close()
    logger.info("Отчёт сохранён: %s", output_path)
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
