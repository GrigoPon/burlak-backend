"""Модуль чтения операционных карт (ОК).

Формат: Множество файлов .xlsx/.xls (распределённых по папкам или архивом).

Алгоритм обработки:
  - Автоматическая фильтрация: операционные карты vs служебные файлы.
  - Каждый файл может содержать несколько листов. Один лист = одна операция.
  - Пустые листы или листы без номера карты — игнорируются.
  - Извлекаются: [Парт-номер запчасти] и [Необходимое количество].
  - Если парт-номер переносится на следующую строку (символ «-» на конце) —
    система склеивает строки.
  - Повторяющиеся детали в одной карте или в разных картах — суммируются.
  - Поддерживаются .xlsx (openpyxl) и .xls (xlrd).
  - Многопоточный парсинг (ProcessPoolExecutor) для больших объёмов (>1500 карт).

Использует эвристический анализатор (heuristic_analyzer.py) для универсального
поиска таблиц деталей и извлечения номеров карт без привязки к брендам.

Класс CardService — обёртка для использования в FastAPI/серверной архитектуре.
"""

from __future__ import annotations

import logging
import os
import re
import tempfile
import zipfile
from concurrent.futures import ProcessPoolExecutor, as_completed
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

from tqdm import tqdm

from burlak_parser.heuristic_analyzer import (
    HeuristicAnalyzer,
    extract_card_number,
)
from burlak_parser.normalizer import (
    normalize_quantity,
    clean_part_number,
    is_valid_part_number,
)

logger = logging.getLogger(__name__)


def _safe_extractall(zf: zipfile.ZipFile, extract_dir: str) -> None:
    """Безопасная распаковка ZIP с защитой от path traversal (zip slip).

    Проверяет, что все пути в архивах не выходят за пределы целевой директории.
    """
    for info in zf.infolist():
        target_path = os.path.normpath(os.path.join(extract_dir, info.filename))
        if not target_path.startswith(os.path.normpath(extract_dir) + os.sep) and target_path != os.path.normpath(extract_dir):
            raise ValueError(f"ZIP path traversal detected: {info.filename}")
    zf.extractall(extract_dir)


# ─── Универсальный загрузчик Excel (.xlsx + .xls) ────────────────────────────


class ExcelReader:
    """Универсальный читатель Excel-файлов.
    Поддерживает .xlsx (openpyxl) и .xls (xlrd).
    Использует openpyxl для .xlsx, xlrd для .xls.
    """

    def __init__(self, file_path: str):
        self.file_path = file_path
        self._wb: Any = None
        self._engine: str = ""
        self._sheet_names: List[str] = []
        self._sheets: Dict[str, Any] = {}
        self._load()

    def _load(self) -> None:
        ext = os.path.splitext(self.file_path)[1].lower()

        if ext == ".xls":
            self._load_via_xlrd()
        else:
            try:
                import openpyxl
                wb = openpyxl.load_workbook(
                    self.file_path, data_only=True,
                )
                self._engine = "openpyxl"
                self._wb = wb
                self._sheet_names = list(wb.sheetnames)
                for sn in self._sheet_names:
                    self._sheets[sn] = wb[sn]
                logger.debug(
                    "Файл %s загружен через openpyxl",
                    os.path.basename(self.file_path),
                )
                return
            except Exception:
                pass
            # Fallback: read_only mode (handles WPS/slightly corrupted files)
            try:
                import openpyxl
                wb = openpyxl.load_workbook(
                    self.file_path, data_only=True, read_only=True,
                )
                self._engine = "openpyxl"
                self._wb = wb
                self._sheet_names = list(wb.sheetnames)
                for sn in self._sheet_names:
                    self._sheets[sn] = wb[sn]
                logger.info(
                    "Файл %s загружен через openpyxl (read_only fallback)",
                    os.path.basename(self.file_path),
                )
                return
            except Exception:
                pass
            self._load_via_xlrd()

    def _load_via_xlrd(self) -> None:
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
            self._sheet_names = list(wb.sheet_names())
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
        else:
            return self._ws.nrows

    @property
    def max_column(self) -> int:
        if self._engine == "openpyxl":
            return self._ws.max_column or 0
        else:
            return self._ws.ncols

    def cell_value(self, row: int, column: int) -> Any:
        try:
            if self._engine == "openpyxl":
                return self._ws.cell(row=row, column=column).value
            else:
                val = self._ws.cell_value(row - 1, column - 1)
                if val == "" or val is None:
                    return None
                if isinstance(val, float) and val == int(val):
                    return int(val)
                return val
        except Exception:
            return None


# ─── Структуры данных ────────────────────────────────────────────────────────


@dataclass
class CardSheetInfo:
    """Информация об одном листе операционной карты."""
    card_number: str
    sheet_name: str
    operation_name: str = ""
    is_valid: bool = False
    has_data: bool = False


@dataclass
class CardPart:
    """Деталь, найденная в операционной карте."""
    part_number: str
    quantity: float
    source_card: str
    source_sheet: str


@dataclass
class CardParseResult:
    """Результат парсинга одной операционной карты."""
    card_number: str
    file_path: str
    sheets: List[CardSheetInfo]
    parts: List[CardPart]
    aggregated_parts: Dict[str, float]  # part_number -> total_qty
    is_service_file: bool = False  # True если файл был определён как служебный


@dataclass
class CardsData:
    """Результат парсинга всех операционных карт."""
    all_parts: Dict[str, float]  # part_number -> суммарное количество
    original_part_numbers: Dict[str, str] = field(default_factory=dict)  # cleaned_part_no -> оригинальный (с тире и т.д.)
    part_sources: Dict[str, List[Tuple[str, str, float]]] = field(default_factory=dict)
    card_results: List[CardParseResult] = field(default_factory=list)
    total_cards_processed: int = 0
    total_sheets_processed: int = 0
    total_sheets_skipped: int = 0
    service_files_skipped: int = 0
    corrupted_files: List[str] = field(default_factory=list)
    split_stats: Optional[SplitStatistics] = None


@dataclass
class FileSplitStats:
    """Статистика разделения для одного файла."""
    file_path: str
    file_name: str
    card_number: str
    is_xlsx: bool
    is_service_file: bool
    total_sheets: int
    sheets_split: int
    sheets_skipped: int
    split_reason: str = ""  # Почему файл был пропущен или что с ним произошло
    skip_reasons: Dict[str, List[str]] = field(default_factory=dict)  # причина -> [имена листов]
    created_files: int = 0
    has_error: bool = False
    error_message: str = ""


@dataclass
class SplitStatistics:
    """Агрегированная статистика разделения всех файлов."""
    file_stats: List[FileSplitStats] = field(default_factory=list)
    total_xlsx: int = 0
    total_xls: int = 0
    total_service_files: int = 0
    total_sheets_all: int = 0
    total_sheets_split: int = 0
    total_sheets_skipped: int = 0
    total_files_created: int = 0
    total_errors: int = 0
    openpyxl_fallback_count: int = 0
    openpyxl_fallback_files: List[str] = field(default_factory=list)

    def get_top_skip_reasons(self, n: int = 5) -> List[Tuple[str, int]]:
        """Топ-N причин пропуска листов по всем файлам."""
        from collections import Counter
        counter: Counter = Counter()
        for fs in self.file_stats:
            for reason, sheets in fs.skip_reasons.items():
                counter[reason] += len(sheets)
        return counter.most_common(n)

    def get_files_with_most_skips(self, n: int = 5) -> List[Tuple[str, int, int]]:
        """Топ-N файлов по количеству пропущенных листов.
        Returns: [(file_name, total_sheets, sheets_skipped)]
        """
        sorted_stats = sorted(
            [fs for fs in self.file_stats if fs.sheets_skipped > 0],
            key=lambda x: x.sheets_skipped,
            reverse=True,
        )
        return [(s.file_name, s.total_sheets, s.sheets_skipped) for s in sorted_stats[:n]]


# ─── Вспомогательные функции ─────────────────────────────────────────────────


def _extract_card_number(file_path: str, ws: "ExcelSheet") -> str:
    """Извлечь номер карты: сначала из содержимого листа, затем из имени файла."""
    # Используем эвристический анализатор
    card_no = extract_card_number(file_path, ws)
    if card_no:
        return card_no

    # Абсолютный fallback: базовое имя файла
    basename = os.path.basename(file_path)
    return os.path.splitext(basename)[0]


def _merge_multiline_part_numbers(
    rows: List[Tuple[int, str, float, str, int]],
) -> List[Tuple[str, float, str, int]]:
    """Склеить парт-номера, перенесённые на следующую строку."""
    merged: List[Tuple[str, float, str, int]] = []
    buffer = ""
    buffer_qty: Optional[float] = None
    buffer_name = ""
    buffer_row = 0
    last_was_continued = False

    for row_idx, raw_part_no, qty, name, _ in rows:
        if last_was_continued:
            buffer += clean_part_number(raw_part_no)
            last_was_continued = False
        elif raw_part_no.rstrip().endswith("-") or raw_part_no.rstrip().endswith("—") or raw_part_no.rstrip().endswith("–"):
            buffer = clean_part_number(raw_part_no.rstrip("-—–"))
            buffer_qty = qty
            buffer_name = name
            buffer_row = row_idx
            last_was_continued = True
            continue
        else:
            buffer = clean_part_number(raw_part_no)
            buffer_qty = qty
            buffer_name = name
            buffer_row = row_idx

        if buffer and buffer_qty is not None:
            if is_valid_part_number(buffer):
                merged.append((buffer, buffer_qty, buffer_name, buffer_row))
            buffer = ""
            buffer_qty = None
            buffer_name = ""

    if buffer and buffer_qty is not None:
        if is_valid_part_number(buffer):
            merged.append((buffer, buffer_qty, buffer_name, buffer_row))

    return merged


# ─── Парсинг одного файла ────────────────────────────────────────────────────


def parse_card_file(
    file_path: str,
    is_service_file: bool = False,
) -> CardParseResult:
    """Разобрать один файл операционной карты.

    Использует эвристический анализатор для поиска таблицы деталей
    и извлечения номеров карт без привязки к конкретным брендам.

    Args:
        file_path: Путь к .xlsx или .xls файлу.
        is_service_file: True если файл классифицирован как служебный.

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

    # Если файл служебный — только собираем информацию о листах, без парсинга деталей
    if is_service_file:
        for sheet_name in reader.sheet_names:
            ws = reader.get_sheet(sheet_name)
            sheet_has_data = _check_sheet_has_data(ws)
            sheets_info.append(CardSheetInfo(
                card_number=card_number or basename,
                sheet_name=sheet_name,
                is_valid=False,
                has_data=sheet_has_data,
            ))
        reader.close()
        return CardParseResult(
            card_number=basename,
            file_path=file_path,
            sheets=sheets_info,
            parts=card_parts,
            aggregated_parts=aggregated,
            is_service_file=True,
        )

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
                has_data=False,
            ))
            continue

        sheet_has_data = _check_sheet_has_data(ws)

        if not sheet_has_data:
            sheets_info.append(CardSheetInfo(
                card_number=card_number or basename,
                sheet_name=sheet_name,
                is_valid=False,
                has_data=False,
            ))
            continue

        # Извлекаем номер карты из первого непустого листа (эвристически)
        if not card_number:
            card_number = _extract_card_number(file_path, ws)

        # Ищем таблицы с деталями через эвристический анализатор
        # Поддерживает многооперационные листы (SWM карты)
        first_table_info = HeuristicAnalyzer.find_part_table(ws)
        if first_table_info is None:
            # Fallback: Changan-формат с 图示编号 (Graphic Number)
            header_rows = HeuristicAnalyzer.find_header_rows(ws)
            col_types = HeuristicAnalyzer.detect_column_types(ws, header_rows) if header_rows else {}
            graphic_parts = _collect_parts_with_graphic_number(
                ws, max_row, max_col, header_rows, col_types, basename,
            )
            if graphic_parts:
                for part_no, qty, name, graphic_number in graphic_parts:
                    card_parts.append(CardPart(
                        part_number=part_no,
                        quantity=qty,
                        source_card=graphic_number or card_number or basename,
                        source_sheet=sheet_name,
                    ))
                    aggregated[part_no] = aggregated.get(part_no, 0.0) + qty

                sheets_info.append(CardSheetInfo(
                    card_number=card_number or basename,
                    sheet_name=sheet_name,
                    operation_name=f"Graphic number linked ({len(graphic_parts)} parts)",
                    is_valid=True,
                    has_data=True,
                ))
            else:
                sheets_info.append(CardSheetInfo(
                    card_number=card_number or basename,
                    sheet_name=sheet_name,
                    operation_name="Лист без таблицы деталей",
                    is_valid=False,
                    has_data=True,
                ))
            continue

        header_row, part_no_col, qty_col, name_col = first_table_info

        # Извлекаем название операции
        operation_name = HeuristicAnalyzer.extract_operation_name(ws, header_row)

        # Собираем детали из ВСЕХ таблиц на листе (многооперационные карты)
        merged_parts = _collect_all_tables(
            ws, max_row, max_col, basename,
        )

        # Добавляем в результаты
        for part_no, qty, name, _ in merged_parts:
            card_parts.append(CardPart(
                part_number=part_no,
                quantity=qty,
                source_card=card_number or basename,
                source_sheet=sheet_name,
            ))
            aggregated[part_no] = aggregated.get(part_no, 0.0) + qty

        sheets_info.append(CardSheetInfo(
            card_number=card_number or basename,
            sheet_name=sheet_name,
            operation_name=operation_name,
            is_valid=len(merged_parts) > 0,
            has_data=True,
        ))

    reader.close()

    return CardParseResult(
        card_number=card_number or basename,
        file_path=file_path,
        sheets=sheets_info,
        parts=card_parts,
        aggregated_parts=aggregated,

    )


def _check_sheet_has_data(ws: ExcelSheet) -> bool:
    """Проверить, есть ли данные на листе (проверка начала и сэмплирование)."""
    max_row = ws.max_row
    max_col = ws.max_column

    for r in range(1, min(max_row, 10) + 1):
        for c in range(1, min(max_col, 10) + 1):
            if ws.cell_value(r, c) is not None:
                return True

    if max_row > 10:
        step = max(1, (max_row - 15) // 10)
        for r in range(15, max_row + 1, step):
            for c in range(1, min(max_col, 10) + 1):
                if ws.cell_value(r, c) is not None:
                    return True

    return False


def _collect_raw_rows(
    ws: ExcelSheet,
    header_row: int,
    max_row: int,
    max_col: int,
    part_no_col: int,
    qty_col: int,
    name_col: int,
    basename: str,
) -> List[Tuple[int, str, float, str, int]]:
    """Собрать сырые строки таблицы деталей.

    Останавливается при обнаружении границы секции:
      - Строка с >= 2 непустыми ячейками, содержащая PART_NO_KEYWORD (новый заголовок)
      - 3+ последовательных пустых строк в колонке part_no

    Returns:
        Список кортежей (row_idx, raw_part_no, qty, name, part_no_col).
    """
    raw_rows: List[Tuple[int, str, float, str, int]] = []
    max_data_row = max_row
    consecutive_empty_pn = 0

    for row_idx in range(header_row + 1, max_data_row + 1):
        try:
            raw_part_no = ws.cell_value(row_idx, part_no_col)

            # ── Проверка на границу секции: новый заголовок ──
            # Строка с >= 2 непустыми ячейками, содержащая PART_NO_KEYWORD
            non_empty = 0
            row_values_check: List[str] = []
            for c in range(1, min(max_col + 1, 25)):
                v = ws.cell_value(row_idx, c)
                if v is not None:
                    non_empty += 1
                    row_values_check.append(str(v).strip().lower())

            if non_empty >= 2:
                # Проверяем, есть ли ячейка с PART_NO_KEYWORD И длина < 50 символов
                # (чтобы не спутать с длинными описаниями, содержащими "деталь")
                has_part_no_keyword_short = any(
                    len(rv) < 50 and any(kw in rv for kw in HeuristicAnalyzer._get_part_no_keywords())
                    for rv in row_values_check
                )
                if has_part_no_keyword_short:
                    # Это новый заголовок таблицы — останавливаем сбор
                    logger.debug(
                        "Граница секции на строке %d (новый заголовок с PART_NO_KEYWORD)",
                        row_idx,
                    )
                    break

            if raw_part_no is None:
                # Проверка на пустую строку
                all_empty = True
                for c in range(1, min(max_col + 1, 20)):
                    if ws.cell_value(row_idx, c) is not None:
                        all_empty = False
                        break
                if all_empty:
                    consecutive_empty_pn += 1
                    if consecutive_empty_pn >= 3:
                        logger.debug(
                            "Граница секции на строке %d (3+ пустых строк)", row_idx,
                        )
                        break
                    continue
                continue

            # Сброс счётчика пустых строк
            consecutive_empty_pn = 0

            raw_part_no_str = str(raw_part_no).strip()
            if not raw_part_no_str:
                continue

            skip_keywords = [
                "物料清单", "变更记录", "编制", "校对", "审核", "批准",
                "说明性符号", "工具", "夹具", "文件编号", "文件版次", "无",
            ]
            if any(kw in raw_part_no_str.lower() for kw in skip_keywords):
                continue

            if qty_col > 0:
                raw_qty = ws.cell_value(row_idx, qty_col)
                try:
                    qty = float(raw_qty) if raw_qty is not None else 1.0
                except (ValueError, TypeError):
                    qty = 1.0
            else:
                qty = 1.0

            name = ""
            if name_col > 0:
                name_val = ws.cell_value(row_idx, name_col)
                name = str(name_val).strip() if name_val is not None else ""

            raw_rows.append((row_idx, raw_part_no_str, qty, name, part_no_col))

        except Exception:
            logger.debug("Ошибка при обработке строки %d в %s, пропускаем", row_idx, basename)
            continue

    return raw_rows


def _collect_all_tables(
    ws: ExcelSheet,
    max_row: int,
    max_col: int,
    basename: str,
) -> List[Tuple[str, float, str, int]]:
    """Собрать детали из ВСЕХ таблиц на листе (многооперационные карты).

    Последовательно находит таблицы деталей через find_part_table(),
    собирает строки из каждой, и агрегирует результаты.
    Пропускает найденные таблицы, если в них нет валидных part-number.

    Returns:
        Список кортежей (part_no, qty, name, source_row) для всех найденных деталей.
    """
    all_parts: List[Tuple[str, float, str, int]] = []
    total_part_nos_collected = 0
    start_search = 1
    max_tables = 200  # поддержка больших файлов (G01: 90+ операционных карт)

    for table_idx in range(max_tables):
        table_info = HeuristicAnalyzer.find_part_table(ws, start_row=start_search)
        if table_info is None:
            break

        header_row, part_no_col, qty_col, name_col = table_info

        # Если заголовок уже обработан — выходим
        if header_row < start_search:
            break

        raw_rows = _collect_raw_rows(
            ws, header_row, max_row, max_col,
            part_no_col, qty_col, name_col, basename,
        )

        merged_parts = _merge_multiline_part_numbers(raw_rows)

        if merged_parts:
            total_part_nos_collected += len(merged_parts)
            all_parts.extend(merged_parts)
            logger.debug(
                "Таблица #%d (R%d): %d деталей",
                table_idx + 1, header_row, len(merged_parts),
            )

        # Продолжаем поиск со следующей строки после последней собранной
        # (или после заголовка, если данных нет)
        last_data_row = header_row
        if raw_rows:
            last_data_row = max(r[0] for r in raw_rows)
        start_search = last_data_row + 1

        # Если таблица оказалась пустой — выходим (защита от цикла)
        if start_search >= max_row:
            break

    if total_part_nos_collected == 0:
        logger.debug("Не найдено таблиц с деталями")

    return all_parts


def _collect_parts_with_graphic_number(
    ws: ExcelSheet,
    max_row: int,
    max_col: int,
    header_rows: List[int],
    col_types: Dict[str, int],
    basename: str,
) -> List[Tuple[str, float, str, str]]:
    """Собрать детали из таблицы с колонкой 图示编号 (Graphic Number).

    Используется для Changan-формата, где детали привязаны к операционным картам
    через колонку с номером схемы/операции (например, DP-CH-A01).

    Args:
        ws: Лист Excel.
        max_row: Максимальная строка.
        max_col: Максимальная колонка.
        header_rows: Найденные строки заголовков.
        col_types: Определённые типы колонок.
        basename: Имя файла.

    Returns:
        Список кортежей (part_no, qty, name, graphic_number).
    """
    graphic_col = HeuristicAnalyzer.find_graphic_number_column(ws, header_rows, col_types)
    if graphic_col == 0:
        return []

    part_no_col = col_types.get("part_no", 0)
    name_col = col_types.get("name_cn", 0) or col_types.get("name_en", 0)

    if part_no_col == 0:
        return []

    # Находим количественную колонку (qty или config columns)
    qty_col = col_types.get("qty", 0)

    data_start = header_rows[-1] + 1 if header_rows else 2
    parts: List[Tuple[str, float, str, str]] = []

    for r in range(data_start, max_row + 1):
        pn = ws.cell_value(r, part_no_col)
        if pn is None:
            continue
        pn_str = str(pn).strip()
        if not pn_str or not is_valid_part_number(pn_str):
            continue

        pn_clean = clean_part_number(pn_str)

        # Количество
        qty = 1.0
        if qty_col > 0:
            qty_val = ws.cell_value(r, qty_col)
            qty = normalize_quantity(qty_val, default=1.0)
            if qty <= 0:
                qty = 1.0

        # Название
        name = ""
        if name_col > 0:
            name_val = ws.cell_value(r, name_col)
            name = str(name_val).strip() if name_val is not None else ""

        # Номер схемы/операции
        graphic_val = ws.cell_value(r, graphic_col)
        graphic_number = str(graphic_val).strip() if graphic_val is not None else ""

        if graphic_number:
            parts.append((pn_clean, qty, name, graphic_number))

    logger.debug(
        "Graphic number extraction: %d parts with graphic_col=%d",
        len(parts), graphic_col,
    )
    return parts


# ─── Поиск файлов ────────────────────────────────────────────────────────────


def _is_os_temp_file(filename: str) -> bool:
    """Проверить, является ли файл служебным файлом ОС.

    Фильтрует:
      - ~$filename.xlsx — файлы блокировки Excel
      - ._filename.xlsx — метаданные macOS (AppleDouble)
      - ~filename.xlsx — временные файлы
    """
    basename = os.path.basename(filename)
    if basename.startswith("~$"):
        return True
    if basename.startswith("._"):
        return True
    if basename.startswith("~") and basename.endswith((".xlsx", ".xls")):
        return True
    return False


def _find_excel_files(path: str, extract_dir: Optional[str] = None, _seen_names: Optional[set] = None) -> List[str]:
    """Найти все .xlsx и .xls файлы рекурсивно (папка или ZIP).

    Поддерживает вложенные ZIP-архивы (рекурсивно) с извлечением
    в отдельные поддиректории.
    Проверяет дубликаты по ИМЕНИ файла (базовое имя без пути).
    Фильтрует временные файлы (~$, ._*) и не-Excel форматы.
    Удаляет мусор только из временных директорий извлечения.
    """
    files: List[str] = []
    if _seen_names is None:
        _seen_names = set()

    if os.path.isfile(path) and path.lower().endswith(".zip"):
        if extract_dir is None:
            extract_dir = tempfile.mkdtemp(prefix="burlak_cards_")
        else:
            os.makedirs(extract_dir, exist_ok=True)

        logger.info("Распаковка архива %s в %s...", path, extract_dir)
        with zipfile.ZipFile(path, "r", metadata_encoding="gbk") as z:
            _safe_extractall(z, extract_dir)

        _walk_extracted_dir(extract_dir, extract_dir, files, _seen_names, is_temp=True)

    elif os.path.isdir(path):
        _walk_extracted_dir(path, extract_dir or path, files, _seen_names, is_temp=False)

    elif os.path.isfile(path) and path.lower().endswith((".xlsx", ".xls")):
        if not _is_os_temp_file(path):
            files.append(path)

    return files


def _walk_extracted_dir(walk_root: str, extract_base: str, files: List[str], _seen_names: set, is_temp: bool, is_nested: bool = False) -> None:
    """Обойти директорию, фильтруя только .xlsx/.xls/.zip.

    Основные файлы (is_nested=False) собираются ВСЕ без дедупликации.
    Вложенные файлы (is_nested=True) проверяются на дубликат по имени
    относительно уже собранных основных.
    """
    nested_zips: List[str] = []

    # Детерминированный обход: сортируем корни и имена файлов
    for root, _, filenames in sorted(os.walk(walk_root), key=lambda x: x[0]):
        for fn in sorted(filenames):
            # Фильтруем файлы блокировки Excel (~$), метаданные macOS (._*), и мусор
            if fn.startswith("~$") or fn.startswith("._"):
                if is_temp:
                    _safe_remove(os.path.join(root, fn))
                continue
            full_path = os.path.join(root, fn)
            ext = os.path.splitext(fn)[1].lower()

            if ext in (".xlsx", ".xls"):
                if is_nested and fn in _seen_names:
                    logger.debug("Пропуск дубликата из вложенного архива: %s", fn)
                    if is_temp:
                        _safe_remove(full_path)
                else:
                    _seen_names.add(fn)
                    files.append(full_path)
            elif ext == ".zip":
                nested_zips.append(full_path)
            else:
                if is_temp:
                    _safe_remove(full_path)

    # Обрабатываем вложенные ZIP ПОСЛЕ основных файлов (отсортировано)
    for full_path in sorted(nested_zips):
        fn = os.path.basename(full_path)
        nested_dir = os.path.join(extract_base, f"_nested_{_safe_name(fn)}")
        os.makedirs(nested_dir, exist_ok=True)
        try:
            with zipfile.ZipFile(full_path, "r", metadata_encoding="gbk") as z:
                _safe_extractall(z, nested_dir)
            _walk_extracted_dir(nested_dir, extract_base, files, _seen_names, is_temp=True, is_nested=True)
        except Exception as e:
            logger.warning("Не удалось распаковать вложенный архив %s: %s", fn, e)
        if is_temp:
            _safe_remove(full_path)


def _safe_remove(file_path: str) -> None:
    """Безопасно удалить файл."""
    try:
        os.remove(file_path)
    except Exception:
        pass


def _safe_name(filename: str) -> str:
    """Безопасное имя для поддиректории."""
    return re.sub(r"[^\w\-]", "_", os.path.splitext(filename)[0], flags=re.ASCII)[:50]


# ─── Парсинг всех карт ───────────────────────────────────────────────────────


def parse_cards(
    input_path: str,
    extract_dir: Optional[str] = None,
    show_progress: bool = True,
    max_workers: Optional[int] = None,
) -> CardsData:
    """Разобрать все операционные карты из указанного источника.

    Выполняет классификацию файлов (операционные vs служебные) и
    параллельный парсинг с использованием ProcessPoolExecutor.

    Args:
        input_path: Путь к папке, ZIP-архиву или одному .xlsx/.xls файлу.
        extract_dir: Директория для извлечения ZIP.
        show_progress: Показывать прогресс-бар.
        max_workers: Количество процессов для параллельного парсинга.

    Returns:
        CardsData с агрегированными данными всех карт.
    """
    from burlak_parser.file_classifier import filter_operational_cards

    all_files = _find_excel_files(input_path, extract_dir)
    all_files.sort()  # Детерминированный порядок
    logger.info("Найдено .xlsx/.xls файлов: %d", len(all_files))

    if not all_files:
        raise FileNotFoundError(f"Не найдено .xlsx/.xls файлов в '{input_path}'")

    # Классифицируем все файлы
    classifications = filter_operational_cards(all_files)

    operational_files = sorted(
        [c for c in classifications if c.should_parse_parts],
        key=lambda c: c.file_path,
    )
    service_files = sorted(
        [c for c in classifications if c.is_service_file],
        key=lambda c: c.file_path,
    )
    skipped_files = sorted(
        [c for c in classifications if not c.should_parse_parts and not c.is_service_file],
        key=lambda c: c.file_path,
    )
    total_classified = len(operational_files) + len(service_files) + len(skipped_files)
    logger.info(
        "Найдено файлов: %d → операционных: %d, служебных: %d, пропущено: %d",
        total_classified, len(operational_files), len(service_files), len(skipped_files),
    )
    if total_classified != len(all_files):
        logger.warning(
            "Несовпадение подсчёта: найдено %d, классифицировано %d (возможны дубликаты)",
            len(all_files), total_classified,
        )

    # Парсим операционные карты
    card_results: List[CardParseResult] = []
    all_aggregated: Dict[str, float] = {}
    part_sources: Dict[str, List[Tuple[str, str, float]]] = {}
    total_sheets = 0
    total_skipped = 0
    corrupted: List[str] = []

    parse_files = [(c.file_path, False) for c in operational_files]

    workers = max_workers or min(os.cpu_count() or 4, max(1, len(parse_files)))

    if workers > 1 and len(parse_files) > 1:
        # Параллельный парсинг
        with ProcessPoolExecutor(max_workers=workers) as executor:
            futures = {
                executor.submit(
                    parse_card_file, fp, is_svc,
                ): fp for fp, is_svc in parse_files
            }
            iterator = tqdm(
                as_completed(futures),
                total=len(futures),
                desc="Парсинг карт",
                unit="файл",
            ) if show_progress else as_completed(futures)

            for future in iterator:
                file_path = futures[future]
                try:
                    result = future.result()
                    card_results.append(result)
                    total_sheets += len(result.sheets)
                    total_skipped += sum(1 for s in result.sheets if not s.is_valid)

                    # Агрегируем
                    for part_no, qty in result.aggregated_parts.items():
                        all_aggregated[part_no] = all_aggregated.get(part_no, 0.0) + qty
                        if part_no not in part_sources:
                            part_sources[part_no] = []
                        part_sources[part_no].append(
                            (result.card_number, result.file_path, qty),
                        )
                        # Сохраняем оригинальный формат номера (с тире и т.д.)
                        # Используем первый попавшийся оригинальный номер из деталей карты

                except Exception as e:
                    logger.warning("Ошибка при обработке %s: %s", file_path, e)
                    corrupted.append(file_path)
                    if show_progress:
                        tqdm.write(f"⚠️  Ошибка: {e}")
    else:
        # Последовательный парсинг
        iterator = tqdm(parse_files, desc="Парсинг карт", unit="файл") if show_progress else parse_files
        for file_path, is_svc in iterator:
            try:
                result = parse_card_file(file_path, is_service_file=is_svc)
                card_results.append(result)
                total_sheets += len(result.sheets)
                total_skipped += sum(1 for s in result.sheets if not s.is_valid)

                for part_no, qty in result.aggregated_parts.items():
                    all_aggregated[part_no] = all_aggregated.get(part_no, 0.0) + qty
                    if part_no not in part_sources:
                        part_sources[part_no] = []
                    part_sources[part_no].append(
                        (result.card_number, result.file_path, qty),
                    )
            except Exception as e:
                logger.warning("Ошибка при обработке %s: %s", file_path, e)
                corrupted.append(file_path)
                if show_progress:
                    tqdm.write(f"⚠️  Ошибка: {e}")

    # Добавляем служебные файлы в card_results (для split_cards)
    for svc in service_files:
        try:
            result = parse_card_file(
                svc.file_path,
                is_service_file=True,
            )
            card_results.append(result)
            total_sheets += len(result.sheets)
            total_skipped += sum(1 for s in result.sheets if not s.is_valid)
        except Exception as e:
            logger.warning("Ошибка при обработке служебного файла %s: %s", svc.file_path, e)

    processed = len(card_results)
    logger.info("Обработано карт: %d", processed)
    logger.info("Всего листов: %d, пропущено (пустых): %d", total_sheets, total_skipped)
    logger.info("Уникальных деталей найдено: %d", len(all_aggregated))

    # Сортируем card_results для детерминированного порядка
    card_results.sort(key=lambda r: r.file_path)

    # Строим словарь оригинальных номеров деталей из карт
    original_part_numbers: Dict[str, str] = {}
    for result in card_results:
        for cp in result.parts:
            clean_pn = clean_part_number(cp.part_number)
            if clean_pn not in original_part_numbers:
                original_part_numbers[clean_pn] = cp.part_number

    return CardsData(
        all_parts=all_aggregated,
        original_part_numbers=original_part_numbers,
        part_sources=part_sources,
        card_results=card_results,
        total_cards_processed=processed,
        total_sheets_processed=total_sheets - total_skipped,
        total_sheets_skipped=total_skipped,
        service_files_skipped=len(service_files),
        corrupted_files=corrupted,
    )


# ─── Сервис ──────────────────────────────────────────────────────────────────


class CardService:
    """Сервис парсинга операционных карт.

    Готов к использованию в серверной архитектуре (FastAPI).
    Поддерживает:
      - Загрузку из файла/папки/ZIP (load)
      - Загрузку из памяти (load_from_bytes) — для HTTP upload
      - Асинхронную загрузку (load_async) — не блокирует event loop
      - Автоочистку временных файлов (cleanup / context manager)
    """

    def __init__(self, max_workers: Optional[int] = None):
        self._cards: Optional[CardsData] = None
        self._temp_paths: List[str] = []
        self._temp_dirs: List[str] = []
        self.max_workers = max_workers or os.cpu_count() or 4

    @property
    def cards(self) -> Optional[CardsData]:
        return self._cards

    @property
    def is_loaded(self) -> bool:
        return self._cards is not None

    def load(self, input_path: str, extract_dir: Optional[str] = None) -> CardsData:
        """Загрузить и распарсить операционные карты.

        Args:
            input_path: Путь к папке/ZIP-архиву с картами.
            extract_dir: Директория для извлечения ZIP.

        Returns:
            Данные CardsData.
        """
        self._cards = parse_cards(
            input_path,
            extract_dir=extract_dir,
            max_workers=self.max_workers,
            show_progress=False,
        )
        return self._cards

    def load_from_bytes(self, data: bytes, filename: str) -> CardsData:
        """Загрузить карты из байтового содержимого (in-memory upload).

        Принимает ZIP-архив, .xlsx или .xls файл как байты.
        Сохраняет во временный файл, извлекает/парсит, возвращает результат.
        Временные файлы будут удалены при вызове cleanup() или выходе из
        контекстного менеджера.

        Args:
            data: Байтовое содержимое (ZIP, .xlsx или .xls).
            filename: Имя файла для определения расширения и формата.

        Returns:
            Данные CardsData.
        """
        suffix = os.path.splitext(filename)[1] or ".zip"
        fd, path = tempfile.mkstemp(suffix=suffix, prefix="cards_upload_")
        os.close(fd)
        with open(path, "wb") as f:
            f.write(data)
        self._temp_paths.append(path)

        # Для ZIP — создаём отдельную директорию извлечения
        extract_dir: Optional[str] = None
        if suffix.lower() == ".zip":
            extract_dir = tempfile.mkdtemp(prefix="cards_extract_")
            self._temp_dirs.append(extract_dir)

        return self.load(path, extract_dir=extract_dir)

    async def load_async(self, data: bytes, filename: str) -> CardsData:
        """Асинхронная загрузка карт из байтов.

        Парсинг CPU-bound — выполняется в отдельном потоке,
        не блокируя event loop.

        Args:
            data: Байтовое содержимое (ZIP, .xlsx или .xls).
            filename: Имя файла для определения расширения.

        Returns:
            Данные CardsData.
        """
        import asyncio
        return await asyncio.to_thread(self.load_from_bytes, data, filename)

    def cleanup(self) -> None:
        """Удалить все временные файлы и директории."""
        for path in self._temp_paths:
            try:
                if os.path.isfile(path):
                    os.remove(path)
            except Exception:
                pass
        for d in self._temp_dirs:
            try:
                if os.path.isdir(d):
                    import shutil
                    shutil.rmtree(d, ignore_errors=True)
            except Exception:
                pass
        self._temp_paths.clear()
        self._temp_dirs.clear()

    def __enter__(self) -> CardService:
        return self

    def __exit__(self, *args: object) -> None:
        self.cleanup()

    def get_all_parts(self) -> Dict[str, float]:
        """Получить все агрегированные детали из карт."""
        if not self._cards:
            raise RuntimeError("Карты не загружены. Вызовите load() сначала.")
        return dict(self._cards.all_parts)

    def get_part_sources(self) -> Dict[str, List[Tuple[str, str, float]]]:
        """Получить источники для каждой детали."""
        if not self._cards:
            raise RuntimeError("Карты не загружены.")
        return dict(self._cards.part_sources)

    def get_card_results(self) -> List[CardParseResult]:
        """Получить результаты парсинга каждой карты."""
        if not self._cards:
            raise RuntimeError("Карты не загружены.")
        return list(self._cards.card_results)


# ─── Разделение листов (делегировано в splitter.py) ─────────────────────────


TEMPLATE_SHEET_KEYWORDS = ["空表", "填写范本", "范本"]


# Причины пропуска листов
SKIP_REASON_TEMPLATE = "Имя листа содержит ключевое слово шаблона"
SKIP_REASON_NO_DATA = "Пустой лист (нет данных)"
SKIP_REASON_NOT_VALID = "Нет валидных деталей"

# Причины пропуска файлов
FILE_SKIP_NOT_XLSX = "Не .xlsx формат (пропущен)"
FILE_SKIP_SERVICE = "Служебный файл (пропущен)"

# Ключевые слова служебных листов — пропускаем при разделении
# (空表 обрабатывается отдельно через TEMPLATE_SHEET_KEYWORDS с проверкой has_data)
_SPLITTER_SERVICE_SHEET_KEYWORDS = ["封面", "目录"]


def split_cards_to_files(
    cards_data: CardsData,
    output_dir: str,
    split_all_non_empty: bool = True,
    max_workers: Optional[int] = None,
) -> List[str]:
    """Разделить многолистовые файлы на отдельные .xlsx файлы.

    Использует CardSplitter из burlak_parser.splitter для ZIP-разделения
    с очисткой named ranges.

    Args:
        cards_data: Данные распарсенных карт.
        output_dir: Директория для сохранения разделённых файлов.
        split_all_non_empty: Если True, разделяет все непустые листы;
                             если False — только листы с таблицей деталей.
        max_workers: Количество процессов для параллельного разделения.

    Returns:
        Список путей к созданным файлам.
    """
    from burlak_parser.splitter import CardSplitter, _safe_filename

    os.makedirs(output_dir, exist_ok=True)
    splitter = CardSplitter(max_workers=max_workers)

    # ── Собираем per-file статистику ──
    tasks: List[Tuple[str, str, List[str], str]] = []
    file_stats: List[FileSplitStats] = []
    total_xlsx = 0
    total_xls = 0
    total_service = 0
    total_sheets_all = 0
    total_sheets_split = 0
    total_sheets_skipped = 0

    for result in cards_data.card_results:
        file_name = os.path.basename(result.file_path)
        ext = os.path.splitext(file_name)[1].lower()
        is_xlsx = ext == ".xlsx"
        is_xls = ext == ".xls"

        total_sheets_all += len(result.sheets)

        # ── Защита: пропуск служебных файлов ДО любых операций ──
        if result.is_service_file:
            total_service += 1
            if is_xlsx:
                total_xlsx += 1
            else:
                total_xls += 1
            file_stats.append(FileSplitStats(
                file_path=result.file_path,
                file_name=file_name,
                card_number=result.card_number,
                is_xlsx=is_xlsx,
                is_service_file=True,
                total_sheets=len(result.sheets),
                sheets_split=0,
                sheets_skipped=len(result.sheets),
                split_reason=FILE_SKIP_SERVICE,
            ))
            continue

        # ── Пропуск не-.xlsx файлов ──
        if not is_xlsx:
            total_xls += 1
            file_stats.append(FileSplitStats(
                file_path=result.file_path,
                file_name=file_name,
                card_number=result.card_number,
                is_xlsx=False,
                is_service_file=False,
                total_sheets=len(result.sheets),
                sheets_split=0,
                sheets_skipped=len(result.sheets),
                split_reason=FILE_SKIP_NOT_XLSX,
            ))
            continue

        total_xlsx += 1

        # Анализируем листы — какие будут разделены, какие пропущены
        sheets_to_split = []
        skip_reasons: Dict[str, List[str]] = {}
        sheets_skipped = 0

        for sheet_info in result.sheets:
            # ── Пропуск служебных листов (封面, 目录, 空表) ──
            is_service_sheet = any(
                kw in sheet_info.sheet_name for kw in _SPLITTER_SERVICE_SHEET_KEYWORDS
            )
            if is_service_sheet:
                skip_reasons.setdefault("Служебный лист (пропущен)", []).append(
                    sheet_info.sheet_name,
                )
                sheets_skipped += 1
                continue

            is_template_name = any(kw in sheet_info.sheet_name for kw in TEMPLATE_SHEET_KEYWORDS)
            if is_template_name and not sheet_info.has_data:
                skip_reasons.setdefault(SKIP_REASON_TEMPLATE, []).append(sheet_info.sheet_name)
                sheets_skipped += 1
                continue
            if split_all_non_empty:
                if sheet_info.has_data:
                    sheets_to_split.append(sheet_info.sheet_name)
                else:
                    skip_reasons.setdefault(SKIP_REASON_NO_DATA, []).append(sheet_info.sheet_name)
                    sheets_skipped += 1
            else:
                if sheet_info.is_valid:
                    sheets_to_split.append(sheet_info.sheet_name)
                else:
                    skip_reasons.setdefault(SKIP_REASON_NOT_VALID, []).append(sheet_info.sheet_name)
                    sheets_skipped += 1

        total_sheets_skipped += sheets_skipped

        if sheets_to_split:
            tasks.append((
                result.file_path,
                output_dir,
                sheets_to_split,
                result.card_number,
            ))
            total_sheets_split += len(sheets_to_split)

        file_stats.append(FileSplitStats(
            file_path=result.file_path,
            file_name=file_name,
            card_number=result.card_number,
            is_xlsx=True,
            is_service_file=result.is_service_file,
            total_sheets=len(result.sheets),
            sheets_split=len(sheets_to_split),
            sheets_skipped=sheets_skipped,
            skip_reasons=skip_reasons,
            split_reason="",
        ))

    if not tasks:
        cards_data.split_stats = SplitStatistics(
            file_stats=file_stats,
            total_xlsx=total_xlsx,
            total_xls=total_xls,
            total_service_files=total_service,
            total_sheets_all=total_sheets_all,
            total_sheets_split=total_sheets_split,
            total_sheets_skipped=total_sheets_skipped,
            total_files_created=0,
        )
        return []

    # Сортируем задачи по имени исходного файла для детерминированного порядка
    tasks.sort(key=lambda t: t[0])

    workers = max_workers or os.cpu_count() or 4
    all_created: List[str] = []
    corrupted: List[str] = []
    openpyxl_count = 0
    openpyxl_files: List[str] = []
    manifest: Dict[str, List[str]] = {}

    if workers > 1 and len(tasks) > 1:
        result_files, split_errors, oxl_count, oxl_files, worker_manifest = (
            splitter.split_many_parallel(tasks)
        )
        all_created.extend(result_files)
        openpyxl_count = oxl_count
        openpyxl_files = oxl_files
        manifest = worker_manifest
        for source_path, err_msg in split_errors:
            corrupted.append(source_path)
            for fs in file_stats:
                if fs.file_path == source_path:
                    fs.has_error = True
                    fs.error_message = err_msg
                    break
    else:
        for source_path, out_dir, sheet_names, file_label in tasks:
            try:
                created = splitter.split_file(source_path, out_dir, sheet_names, file_label)
                all_created.extend(created)
                openpyxl_count += splitter.openpyxl_fallback_count
                openpyxl_files.extend(sorted(splitter.openpyxl_fallback_files))
                for orig, generated in splitter.manifest.items():
                    manifest.setdefault(orig, []).extend(generated)
                # Reset per-file counters for next iteration
                splitter.openpyxl_fallback_count = 0
                splitter.openpyxl_fallback_files.clear()
                splitter.manifest.clear()
            except Exception as e:
                err_msg = str(e)
                logger.warning("Повреждённый файл при разделении %s: %s", os.path.basename(source_path), err_msg)
                corrupted.append(source_path)
                # Обновить статистику для этого файла
                for fs in file_stats:
                    if fs.file_path == source_path:
                        fs.has_error = True
                        fs.error_message = err_msg
                        break

    # Сортируем результаты для детерминированного порядка
    all_created.sort()
    corrupted.sort()

    if corrupted:
        logger.warning("Повреждённых файлов при разделении: %d", len(corrupted))
        # Выводим имена повреждённых файлов в лог
        for cf in corrupted:
            logger.warning("  ⚠️  %s", os.path.basename(cf))
    if getattr(cards_data, 'corrupted_files', None) is not None:
        cards_data.corrupted_files.extend(corrupted)
    else:
        cards_data.corrupted_files = list(corrupted)

    # Подсчёт created_files на задачу: для каждого task сопоставляем
    # созданные файлы по префиксу из safe_label (совпадает с именованием splitter)
    all_created_basenames: List[str] = [os.path.basename(f) for f in all_created]
    task_file_to_stats: Dict[str, FileSplitStats] = {}
    for fs in file_stats:
        if not fs.split_reason:
            task_file_to_stats[fs.file_path] = fs
    for source_path, _out_dir, _sheets, file_label in tasks:
        if source_path in task_file_to_stats:
            safe_label = _safe_filename(file_label)[:50]
            created_count = sum(1 for bn in all_created_basenames if bn.startswith(safe_label))
            task_file_to_stats[source_path].created_files = created_count

    total_errors = len(corrupted)

    cards_data.split_stats = SplitStatistics(
        file_stats=sorted(file_stats, key=lambda fs: fs.file_path),
        total_xlsx=total_xlsx,
        total_xls=total_xls,
        total_service_files=total_service,
        total_sheets_all=total_sheets_all,
        total_sheets_split=total_sheets_split,
        total_sheets_skipped=total_sheets_skipped,
        total_files_created=len(all_created),
        total_errors=total_errors,
        openpyxl_fallback_count=openpyxl_count,
        openpyxl_fallback_files=sorted(set(openpyxl_files)),
    )

    logger.info("Создано отдельных файлов: %d", len(all_created))
    if openpyxl_count > 0:
        logger.info(
            "Успешно спасены через openpyxl (метод fallback): %d файлов: %s",
            openpyxl_count, sorted(set(openpyxl_files)),
        )

    # Сохраняем манифест генерации файлов
    if manifest:
        import json
        manifest_path = os.path.join(output_dir, "split_manifest.json")
        # Сортируем для детерминированного вывода
        sorted_manifest = {k: sorted(v) for k, v in sorted(manifest.items())}
        try:
            with open(manifest_path, "w", encoding="utf-8") as f:
                json.dump(sorted_manifest, f, ensure_ascii=False, indent=2)
            logger.info("Журнал генерации файлов сохранен: %s", manifest_path)
        except Exception as e:
            logger.warning("Не удалось сохранить манифест: %s", e)

    return all_created
