"""Модуль чтения операционных карт (ОК).

Формат: Множество файлов .xlsx (распределённых по папкам или загружаемых архивом).
Специфика обработки:
  - Каждый файл может содержать несколько листов. Один лист = одна операция.
  - Пустые листы или листы без номера карты — игнорируются.
  - Извлекаются: [Парт-номер запчасти] и [Необходимое количество].
  - Если парт-номер переносится на следующую строку (символ «-» на конце) —
    система склеивает строки.
  - Повторяющиеся детали в одной карте или в разных картах — суммируются.
  - Поддерживаются .xlsx (openpyxl) и .xls (xlrd).
"""

from __future__ import annotations

import logging
import os
import re
import tempfile
import zipfile
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from tqdm import tqdm

logger = logging.getLogger(__name__)


# ─── Универсальный загрузчик Excel (.xlsx + .xls) ────────────────────────────

class ExcelReader:
    """Универсальный читатель Excel-файлов.
    Поддерживает .xlsx (openpyxl) и .xls (xlrd).
    Использует openpyxl для .xlsx, xlrd для .xls.
    """

    def __init__(self, file_path: str):
        self.file_path = file_path
        self._wb: Any = None
        self._engine: str = ""  # 'openpyxl' или 'xlrd'
        self._sheet_names: List[str] = []
        self._sheets: Dict[str, Any] = {}
        self._load()

    def _load(self) -> None:
        """Загрузить файл. Выбор движка по расширению."""
        ext = os.path.splitext(self.file_path)[1].lower()

        if ext == ".xls":
            # .xls → сразу xlrd (openpyxl может зависнуть на .xls!)
            self._load_via_xlrd()
        else:
            # .xlsx (и всё остальное) → openpyxl, при ошибке — xlrd
            try:
                wb = openpyxl.load_workbook(self.file_path, data_only=True)
                self._engine = "openpyxl"
                self._wb = wb
                self._sheet_names = wb.sheetnames
                for sn in self._sheet_names:
                    self._sheets[sn] = wb[sn]
                logger.debug(
                    "Файл %s загружен через openpyxl",
                    os.path.basename(self.file_path),
                )
                return
            except Exception:
                pass
            # Если openpyxl не справился, пробуем xlrd
            self._load_via_xlrd()

    def _load_via_xlrd(self) -> None:
        """Загрузить файл через xlrd."""
        try:
            import xlrd
        except ImportError:
            raise ImportError(
                "Для чтения .xls файлов требуется xlrd. Установите: pip install xlrd"
            )

        try:
            wb = xlrd.open_workbook(self.file_path)
            self._engine = "xlrd"
            self._wb = wb
            self._sheet_names = wb.sheet_names()
            for sn in self._sheet_names:
                self._sheets[sn] = wb.sheet_by_name(sn)
            logger.debug(
                "Файл %s загружен через xlrd",
                os.path.basename(self.file_path),
            )
        except Exception as e:
            raise ValueError(
                f"Не удалось открыть Excel-файл {self.file_path}: {e}"
            )

    @property
    def sheet_names(self) -> List[str]:
        return self._sheet_names

    def get_sheet(self, name: str) -> "ExcelSheet":
        if name not in self._sheets:
            raise KeyError(f"Лист '{name}' не найден")
        return ExcelSheet(self._sheets[name], self._engine)

    def close(self) -> None:
        if self._engine == "openpyxl" and self._wb is not None:
            self._wb.close()


class ExcelSheet:
    """Обёртка над листом Excel для единого API openpyxl / xlrd."""

    def __init__(self, ws: Any, engine: str):
        self._ws = ws
        self._engine = engine

    @property
    def max_row(self) -> int:
        if self._engine == "openpyxl":
            return self._ws.max_row or 0
        else:  # xlrd
            return self._ws.nrows

    @property
    def max_column(self) -> int:
        if self._engine == "openpyxl":
            return self._ws.max_column or 0
        else:  # xlrd
            return self._ws.ncols

    def cell_value(self, row: int, column: int) -> Any:
        """Получить значение ячейки (row и column — 1-индексированные)."""
        try:
            if self._engine == "openpyxl":
                return self._ws.cell(row=row, column=column).value
            else:  # xlrd
                val = self._ws.cell_value(row - 1, column - 1)
                # xlrd возвращает пустую строку для пустых ячеек, приводим к None
                if val == "" or val is None:
                    return None
                # xlrd возвращает float для чисел, преобразуем в int если целое
                if isinstance(val, float) and val == int(val):
                    return int(val)
                return val
        except Exception:
            return None

    def iter_rows_values(self, min_row: int, max_row: int, max_col: int) -> List[List[Any]]:
        """Получить диапазон значений как список строк."""
        result: List[List[Any]] = []
        for r in range(min_row, max_row + 1):
            row_vals: List[Any] = []
            for c in range(1, max_col + 1):
                row_vals.append(self.cell_value(r, c))
            result.append(row_vals)
        return result


# Шаблоны для поиска заголовков таблицы с деталями в операционных картах
PART_TABLE_HEADER_KEYWORDS = [
    "物料清单", "bom", "零件清单", "物料列表",
    "序", "号", "料号", "零件号", "part no", "partno",
    "物料名称", "零件名称", "描述",
    "规格", "单车用量", "用量", "qty", "数量",
    "适用版型",
]

# Шаблоны для поиска заголовка таблицы (столбцов) — должны быть в одной строке
TABLE_COLUMN_PATTERNS = [
    # 序号 | 料号 | 物料名称/描述 | 规格 | 单车用量 | 适用版型
    [r"序.*号", r"料号|零件号|part\s*no", r"物料名称|零件名称|描述", r"规格", r"单车用量|用量|qty|数量"],
    # Более короткие таблицы
    [r"序.*号", r"料号|零件号|part\s*no", r"物料名称|零件名称|描述", r"单车用量|用量|qty"],
    [r"料号|零件号|part\s*no", r"物料名称|零件名称", r"数量|用量|qty"],
    [r"序.*号", r"零件号|part\s*no", r"零件名称", r"数量|用量"],
]


@dataclass
class CardSheetInfo:
    """Информация об одном листе операционной карты."""
    card_number: str  # Номер карты (из документа / имени файла)
    sheet_name: str   # Имя листа
    operation_name: str = ""  # Название операции
    is_valid: bool = False    # Содержит ли лист полезные данные


@dataclass
class CardPart:
    """Деталь, найденная в операционной карте."""
    part_number: str
    quantity: float
    source_card: str  # Номер карты
    source_sheet: str  # Имя листа


@dataclass
class CardParseResult:
    """Результат парсинга одной операционной карты."""
    card_number: str
    file_path: str
    sheets: List[CardSheetInfo]
    parts: List[CardPart]
    # Агрегированные детали по всем листам карты
    aggregated_parts: Dict[str, float]  # part_number -> total_qty


@dataclass
class CardsData:
    """Результат парсинга всех операционных карт."""
    all_parts: Dict[str, float]  # part_number -> суммарное количество по ВСЕМ картам
    part_sources: Dict[str, List[Tuple[str, str, float]]]  # part_number -> [(card_number, sheet_name, qty)]
    card_results: List[CardParseResult]
    total_cards_processed: int = 0
    total_sheets_processed: int = 0
    total_sheets_skipped: int = 0


def _extract_card_number(file_path: str, ws: "ExcelSheet") -> str:
    """Извлечь номер карты из файла.

    Сначала проверяет содержимое листа на наличие номера документа,
    затем использует имя файла.
    """
    max_row = min(10, ws.max_row or 10)
    max_col = min(10, ws.max_column or 10)
    for row_idx in range(1, max_row + 1):
        for col_idx in range(1, max_col + 1):
            val = ws.cell_value(row_idx, col_idx)
            if val is not None:
                text = str(val).strip()
                match = re.search(r"(SQRT[\w-]+)", text)
                if match:
                    return match.group(1)

    basename = os.path.basename(file_path)
    name = os.path.splitext(basename)[0]
    return name


def _find_part_table(ws: "ExcelSheet") -> Optional[Tuple[int, int, int, int]]:
    """Найти таблицу с деталями в листе.

    Ищет строку заголовка, содержащую "料号" или "零件号" (парт-номер).
    Если также найдена колонка количества (单车用量/用量/数量/qty) — использует её.
    Если колонка количества не найдена — quantity по умолчанию = 1.

    Returns:
        (header_row, part_no_col, qty_col, name_col) или None, если таблица не найдена.
        qty_col может быть 0, если колонка количества не найдена.
    """
    max_row = ws.max_row or 200
    max_col = ws.max_column or 100

    for row_idx in range(1, max_row + 1):
        row_values: List[str] = []
        for col_idx in range(1, max_col + 1):
            val = ws.cell_value(row_idx, col_idx)
            row_values.append(str(val).strip().lower() if val is not None else "")

        if not any(row_values):
            continue

        has_part_no = any(
            kw in v
            for v in row_values
            for kw in ["料号", "零件号", "件号", "物料编码", "partno", "part no"]
        )
        if not has_part_no:
            continue

        part_no_col: Optional[int] = None
        qty_col: Optional[int] = None
        name_col: Optional[int] = None

        for col_idx, val in enumerate(row_values, 1):
            if any(kw in val for kw in ["料号", "零件号", "件号", "物料编码", "partno", "part no"]):
                part_no_col = col_idx
            elif any(kw in val for kw in ["用量", "数量", "qty", "单车用量"]):
                qty_col = col_idx
            elif any(kw in val for kw in ["物料名称", "零件名称", "描述", "物料描述", "name"]):
                name_col = col_idx

        if part_no_col is not None:
            logger.debug(
                "Таблица деталей найдена: строка %d, part_no=%s, qty=%s, name=%s",
                row_idx, part_no_col, qty_col, name_col,
            )
            return (row_idx, part_no_col, qty_col or 0, name_col or 0)

    return None


def _extract_part_number(text: str) -> str:
    """Очистить и нормализовать парт-номер.

    Удаляет пробелы, переносы строк, лишние символы.
    """
    cleaned = re.sub(r"[\s\n\r\t]+", "", text.strip())
    # Удаляем символы-разделители на конце (признак переноса)
    cleaned = cleaned.rstrip("-—–")
    return cleaned


def _merge_multiline_part_numbers(rows: List[Tuple[int, str, float, str, int]]) -> List[Tuple[str, float, str, int]]:
    """Склеить парт-номера, перенесённые на следующую строку.

    Если строка заканчивается на '-', значит парт-номер продолжается на следующей строке.

    Args:
        rows: Список (row_idx, raw_part_no, qty, name, source_col) из сырых данных.

    Returns:
        Список (part_no, qty, name, first_row_idx) после склейки.
    """
    merged: List[Tuple[str, float, str, int]] = []
    buffer = ""
    buffer_qty: Optional[float] = None
    buffer_name = ""
    buffer_row = 0
    last_was_continued = False

    for row_idx, raw_part_no, qty, name, _ in rows:
        if last_was_continued:
            # Продолжение предыдущего номера
            buffer += _extract_part_number(raw_part_no)
            last_was_continued = False
        elif raw_part_no.rstrip().endswith("-") or raw_part_no.rstrip().endswith("—") or raw_part_no.rstrip().endswith("–"):
            # Начало переноса
            buffer = _extract_part_number(raw_part_no.rstrip("-—–"))
            buffer_qty = qty
            buffer_name = name
            buffer_row = row_idx
            last_was_continued = True
            continue
        else:
            buffer = _extract_part_number(raw_part_no)
            buffer_qty = qty
            buffer_name = name
            buffer_row = row_idx

        # Если у нас есть данные, добавляем их
        if buffer and buffer_qty is not None:
            merged.append((buffer, buffer_qty, buffer_name, buffer_row))
            buffer = ""
            buffer_qty = None
            buffer_name = ""

    # Добавляем оставшийся буфер, если есть
    if buffer and buffer_qty is not None:
        merged.append((buffer, buffer_qty, buffer_name, buffer_row))

    return merged


def parse_card_file(file_path: str) -> CardParseResult:
    """Разобрать один файл операционной карты.

    Args:
        file_path: Путь к .xlsx или .xls файлу.

    Returns:
        CardParseResult с данными всех непустых листов.
    """
    basename = os.path.basename(file_path)
    logger.debug("Обработка файла: %s", basename)

    reader = ExcelReader(file_path)
    card_parts: List[CardPart] = []
    sheets_info: List[CardSheetInfo] = []
    aggregated: Dict[str, float] = {}
    card_number = ""

    for sheet_name in reader.sheet_names:
        ws = reader.get_sheet(sheet_name)
        max_row = ws.max_row
        max_col = ws.max_column

        # Пропускаем пустые листы
        if max_row == 0 or max_col == 0:
            sheets_info.append(CardSheetInfo(
                card_number=card_number or basename,
                sheet_name=sheet_name,
                is_valid=False,
            ))
            continue

        # Проверяем, есть ли вообще какие-то данные (первые 5 строк)
        has_data = False
        for r in range(1, min(max_row, 5) + 1):
            for c in range(1, min(max_col, 10) + 1):
                if ws.cell_value(r, c) is not None:
                    has_data = True
                    break
            if has_data:
                break

        if not has_data:
            sheets_info.append(CardSheetInfo(
                card_number=card_number or basename,
                sheet_name=sheet_name,
                is_valid=False,
            ))
            continue

        # Извлекаем номер карты из первого листа
        if not card_number:
            card_number = _extract_card_number(file_path, ws)

        # Ищем таблицу с деталями
        table_info = _find_part_table(ws)
        if table_info is None:
            sheets_info.append(CardSheetInfo(
                card_number=card_number or basename,
                sheet_name=sheet_name,
                operation_name="Лист без таблицы деталей",
                is_valid=False,
            ))
            continue

        header_row, part_no_col, qty_col, name_col = table_info

        # Извлекаем название операции
        operation_name = ""
        for r in range(1, min(header_row, 15)):
            for c in range(1, min(max_col + 1, 10)):
                val = ws.cell_value(r, c)
                if val is None:
                    continue
                text = str(val).strip()
                if "作业要素" in text:
                    for check_c in range(c + 1, min(c + 3, max_col + 1)):
                        next_val = ws.cell_value(r, check_c)
                        if next_val and len(str(next_val).strip()) > 1 and "作业要素" not in str(next_val):
                            operation_name = str(next_val).strip()
                            break
                    if not operation_name:
                        next_val = ws.cell_value(r + 1, c)
                        if next_val and len(str(next_val).strip()) > 1:
                            operation_name = str(next_val).strip()
                elif operation_name == "" and len(text) > 3 and \
                     not any(kw in text for kw in ["作业指导书", "文件编号", "工具/夹具",
                                                    "版本", "发行时间", "关键点", "车间",
                                                    "序号", "变更记录", "物料清单",
                                                    "说明性符号", "编制", "校对"]):
                    if any("\u4e00" <= ch <= "\u9fff" for ch in text):
                        operation_name = text

        # Собираем строки таблицы с деталями (каждую строку try/except)
        raw_rows: List[Tuple[int, str, float, str, int]] = []
        max_data_row = min(max_row, header_row + 500)

        for row_idx in range(header_row + 1, max_data_row + 1):
            try:
                raw_part_no = ws.cell_value(row_idx, part_no_col)
                if raw_part_no is None:
                    all_empty = True
                    for c in range(1, min(max_col + 1, 20)):
                        if ws.cell_value(row_idx, c) is not None:
                            all_empty = False
                            break
                    if all_empty:
                        continue
                    continue

                raw_part_no_str = str(raw_part_no).strip()
                if not raw_part_no_str:
                    continue

                # Пропускаем служебные строки
                skip_keywords = ["物料清单", "变更记录", "编制", "校对", "审核", "批准", "说明性符号",
                                 "工具", "夹具", "文件编号", "文件版次", "无"]
                if any(kw in raw_part_no_str.lower() for kw in skip_keywords):
                    continue

                # Извлекаем количество
                if qty_col > 0:
                    raw_qty = ws.cell_value(row_idx, qty_col)
                    try:
                        qty = float(raw_qty) if raw_qty is not None else 1.0
                    except (ValueError, TypeError):
                        qty = 1.0
                else:
                    qty = 1.0

                # Извлекаем название
                name = ""
                if name_col > 0:
                    name_val = ws.cell_value(row_idx, name_col)
                    name = str(name_val).strip() if name_val is not None else ""

                raw_rows.append((row_idx, raw_part_no_str, qty, name, part_no_col))

            except Exception:
                # Если строка битая — просто пропускаем её
                logger.debug("Ошибка при обработке строки %d в %s, пропускаем", row_idx, basename)
                continue

        # Склеиваем перенесённые парт-номера
        merged_parts = _merge_multiline_part_numbers(raw_rows)

        # Добавляем в результаты
        for part_no, qty, name, _ in merged_parts:
            card_parts.append(CardPart(
                part_number=part_no,
                quantity=qty,
                source_card=card_number or basename,
                source_sheet=sheet_name,
            ))
            if part_no in aggregated:
                aggregated[part_no] += qty
            else:
                aggregated[part_no] = qty

        sheets_info.append(CardSheetInfo(
            card_number=card_number or basename,
            sheet_name=sheet_name,
            operation_name=operation_name,
            is_valid=len(merged_parts) > 0,
        ))

    reader.close()

    return CardParseResult(
        card_number=card_number or basename,
        file_path=file_path,
        sheets=sheets_info,
        parts=card_parts,
        aggregated_parts=aggregated,
    )


def _find_excel_files(path: str, extract_dir: Optional[str] = None) -> List[str]:
    """Найти все .xlsx и .xls файлы по указанному пути (рекурсивно).

    Поддерживает как папки, так и ZIP-архивы.
    Для ZIP-архива файлы извлекаются в extract_dir (или persistent temp dir).
    """
    files: List[str] = []

    if os.path.isfile(path) and path.lower().endswith(".zip"):
        if extract_dir is None:
            extract_dir = tempfile.mkdtemp(prefix="burlak_cards_")
        else:
            os.makedirs(extract_dir, exist_ok=True)

        logger.info("Распаковка архива %s в %s...", path, extract_dir)
        with zipfile.ZipFile(path, "r") as z:
            z.extractall(extract_dir)

        for root, _, filenames in os.walk(extract_dir):
            for fn in filenames:
                if (fn.endswith(".xlsx") or fn.endswith(".xls")) and not fn.startswith("~$"):
                    files.append(os.path.join(root, fn))
    elif os.path.isdir(path):
        for root, _, filenames in os.walk(path):
            for fn in filenames:
                if (fn.endswith(".xlsx") or fn.endswith(".xls")) and not fn.startswith("~$"):
                    files.append(os.path.join(root, fn))
    elif os.path.isfile(path) and (path.endswith(".xlsx") or path.endswith(".xls")):
        files.append(path)

    return files


def parse_cards(input_path: str,
                extract_dir: Optional[str] = None,
                show_progress: bool = True) -> CardsData:
    """Разобрать все операционные карты из указанного источника.

    Args:
        input_path: Путь к папке с картами, ZIP-архиву или одному .xlsx/.xls файлу.
        extract_dir: Директория для извлечения ZIP (если None, создаётся временная).
        show_progress: Показывать прогресс-бар.

    Returns:
        CardsData с агрегированными данными всех карт.
    """
    files = _find_excel_files(input_path, extract_dir)
    logger.info("Найдено .xlsx/.xls файлов: %d", len(files))

    if not files:
        raise FileNotFoundError(f"Не найдено .xlsx/.xls файлов в '{input_path}'")

    card_results: List[CardParseResult] = []
    all_aggregated: Dict[str, float] = {}
    part_sources: Dict[str, List[Tuple[str, str, float]]] = {}
    total_sheets = 0
    total_skipped = 0

    iterator = tqdm(files, desc="Парсинг карт", unit="файл") if show_progress else files

    for file_path in iterator:
        try:
            result = parse_card_file(file_path)
            card_results.append(result)
            total_sheets += len(result.sheets)
            total_skipped += sum(1 for s in result.sheets if not s.is_valid)

            # Агрегируем по всем картам
            for part_no, qty in result.aggregated_parts.items():
                if part_no in all_aggregated:
                    all_aggregated[part_no] += qty
                else:
                    all_aggregated[part_no] = qty

                # Запоминаем источники
                if part_no not in part_sources:
                    part_sources[part_no] = []
                part_sources[part_no].append(
                    (result.card_number, result.file_path, qty)
                )

        except Exception as e:
            logger.warning("Ошибка при обработке %s: %s", file_path, e)
            if show_progress:
                tqdm.write(f"⚠️  Ошибка: {e}")

    logger.info("Обработано карт: %d", len(card_results))
    logger.info("Всего листов: %d, пропущено (пустых): %d", total_sheets, total_skipped)
    logger.info("Уникальных деталей найдено: %d", len(all_aggregated))

    return CardsData(
        all_parts=all_aggregated,
        part_sources=part_sources,
        card_results=card_results,
        total_cards_processed=len(card_results),
        total_sheets_processed=total_sheets - total_skipped,
        total_sheets_skipped=total_skipped,
    )


def split_cards_to_files(cards_data: CardsData, output_dir: str) -> List[str]:
    """Разделить многолистовые файлы на отдельные .xlsx файлы.

    Каждый лист операционной карты = отдельный .xlsx файл.
    Пропускаются пустые листы.
    .xls файлы не разделяются (только .xlsx поддерживает запись).

    Args:
        cards_data: Данные распарсенных карт.
        output_dir: Директория для сохранения разделённых файлов.

    Returns:
        Список путей к созданным файлам.
    """
    os.makedirs(output_dir, exist_ok=True)
    created_files: List[str] = []

    for result in cards_data.card_results:
        if not result.file_path.lower().endswith(".xlsx"):
            logger.debug("Пропускаем разделение .xls файла: %s", result.file_path)
            continue

        try:
            wb = openpyxl.load_workbook(result.file_path)
        except Exception as e:
            logger.warning("Не удалось открыть %s: %s", result.file_path, e)
            continue

        for sheet_info in result.sheets:
            if not sheet_info.is_valid:
                continue

            new_wb = openpyxl.Workbook()
            new_ws = new_wb.active

            src_ws = wb[sheet_info.sheet_name]
            for row in src_ws.iter_rows(min_row=1, max_row=src_ws.max_row or 0,
                                         max_col=src_ws.max_column or 0):
                for cell in row:
                    new_ws.cell(row=cell.row, column=cell.column, value=cell.value)

            safe_card = re.sub(r"[^\w\-]", "_", result.card_number)[:50]
            safe_sheet = re.sub(r"[^\w\-]", "_", sheet_info.sheet_name)[:50]
            output_filename = f"{safe_card}_{safe_sheet}.xlsx"
            output_path = os.path.join(output_dir, output_filename)

            counter = 1
            while os.path.exists(output_path):
                base, ext = os.path.splitext(output_filename)
                output_path = os.path.join(output_dir, f"{base}_{counter}{ext}")
                counter += 1

            new_wb.save(output_path)
            new_wb.close()
            created_files.append(output_path)

        wb.close()

    logger.info("Создано отдельных файлов: %d", len(created_files))
    return created_files
