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
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Set, Tuple

from tqdm import tqdm

from burlak_parser.heuristic_analyzer import (
    HeuristicAnalyzer,
    extract_card_number,
    clean_part_number,
    is_valid_part_number,
)

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
    is_final_check: bool = False   # True если из папки CP7/CP8


@dataclass
class CardsData:
    """Результат парсинга всех операционных карт."""
    all_parts: Dict[str, float]  # part_number -> суммарное количество
    part_sources: Dict[str, List[Tuple[str, str, float]]]  # part_number -> [(card, file, qty)]
    card_results: List[CardParseResult]
    total_cards_processed: int = 0
    total_sheets_processed: int = 0
    total_sheets_skipped: int = 0
    service_files_skipped: int = 0
    corrupted_files: List[str] = None


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
    is_final_check: bool = False,
) -> CardParseResult:
    """Разобрать один файл операционной карты.

    Использует эвристический анализатор для поиска таблицы деталей
    и извлечения номеров карт без привязки к конкретным брендам.

    Args:
        file_path: Путь к .xlsx или .xls файлу.
        is_service_file: True если файл классифицирован как служебный.
        is_final_check: True если файл из папки CP7/CP8.

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
            is_final_check=is_final_check,
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
        is_final_check=is_final_check,
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
    max_data_row = min(max_row, header_row + 500)
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
    max_tables = 10  # защита от бесконечного цикла

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


# ─── Поиск файлов ────────────────────────────────────────────────────────────


def _find_excel_files(path: str, extract_dir: Optional[str] = None, _seen_names: Optional[set] = None) -> List[str]:
    """Найти все .xlsx и .xls файлы рекурсивно (папка или ZIP).

    Поддерживает вложенные ZIP-архивы (рекурсивно) с извлечением
    в отдельные поддиректории.
    Проверяет дубликаты по ИМЕНИ файла (базовое имя без пути).
    Фильтрует временные файлы (~$) и не-Excel форматы.
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
            z.extractall(extract_dir)

        _walk_extracted_dir(extract_dir, extract_dir, files, _seen_names, is_temp=True)

    elif os.path.isdir(path):
        _walk_extracted_dir(path, extract_dir or path, files, _seen_names, is_temp=False)

    elif os.path.isfile(path) and path.lower().endswith((".xlsx", ".xls")):
        files.append(path)

    return files


def _walk_extracted_dir(walk_root: str, extract_base: str, files: List[str], _seen_names: set, is_temp: bool, is_nested: bool = False) -> None:
    """Обойти директорию, фильтруя только .xlsx/.xls/.zip.

    Основные файлы (is_nested=False) собираются ВСЕ без дедупликации.
    Вложенные файлы (is_nested=True) проверяются на дубликат по имени
    относительно уже собранных основных.
    """
    nested_zips: List[str] = []

    for root, _, filenames in os.walk(walk_root):
        for fn in filenames:
            if fn.startswith("~$"):
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

    # Обрабатываем вложенные ZIP ПОСЛЕ основных файлов
    for full_path in nested_zips:
        fn = os.path.basename(full_path)
        nested_dir = os.path.join(extract_base, f"_nested_{_safe_name(fn)}")
        os.makedirs(nested_dir, exist_ok=True)
        try:
            with zipfile.ZipFile(full_path, "r", metadata_encoding="gbk") as z:
                z.extractall(nested_dir)
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
    logger.info("Найдено .xlsx/.xls файлов: %d", len(all_files))

    if not all_files:
        raise FileNotFoundError(f"Не найдено .xlsx/.xls файлов в '{input_path}'")

    # Классифицируем все файлы
    classifications = filter_operational_cards(all_files)

    operational_files = [
        c for c in classifications if c.should_parse_parts
    ]
    service_files = [
        c for c in classifications if not c.should_parse_parts
    ]
    logger.info(
        "Операционных карт: %d, служебных файлов: %d",
        len(operational_files), len(service_files),
    )

    # Парсим операционные карты
    card_results: List[CardParseResult] = []
    all_aggregated: Dict[str, float] = {}
    part_sources: Dict[str, List[Tuple[str, str, float]]] = {}
    total_sheets = 0
    total_skipped = 0
    corrupted: List[str] = []

    parse_files = [(c.file_path, False, c.is_final_check) for c in operational_files]

    workers = max_workers or min(os.cpu_count() or 4, max(1, len(parse_files)))

    if workers > 1 and len(parse_files) > 1:
        # Параллельный парсинг
        with ProcessPoolExecutor(max_workers=workers) as executor:
            futures = {
                executor.submit(
                    parse_card_file, fp, is_svc, is_fc,
                ): fp for fp, is_svc, is_fc in parse_files
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
                except Exception as e:
                    logger.warning("Ошибка при обработке %s: %s", file_path, e)
                    corrupted.append(file_path)
                    if show_progress:
                        tqdm.write(f"⚠️  Ошибка: {e}")
    else:
        # Последовательный парсинг
        iterator = tqdm(parse_files, desc="Парсинг карт", unit="файл") if show_progress else parse_files
        for file_path, is_svc, is_fc in iterator:
            try:
                result = parse_card_file(file_path, is_service_file=is_svc, is_final_check=is_fc)
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
                is_final_check=svc.is_final_check,
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

    return CardsData(
        all_parts=all_aggregated,
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

    Подготовлен для миграции на серверную архитектуру (FastAPI + SQLite + Redis).
    Инкапсулирует всю логику парсинга и разделения карт в одном классе.
    """

    def __init__(self, max_workers: Optional[int] = None):
        self._cards: Optional[CardsData] = None
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
        )
        return self._cards

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
    from burlak_parser.splitter import CardSplitter

    os.makedirs(output_dir, exist_ok=True)
    splitter = CardSplitter(max_workers=max_workers)

    tasks: List[Tuple[str, str, List[str], str]] = []

    for result in cards_data.card_results:
        if not result.file_path.lower().endswith(".xlsx"):
            continue

        if result.is_final_check:
            continue

        if result.is_service_file:
            continue

        sheets_to_split = []
        for sheet_info in result.sheets:
            if any(kw in sheet_info.sheet_name for kw in TEMPLATE_SHEET_KEYWORDS):
                continue
            if split_all_non_empty:
                if sheet_info.has_data:
                    sheets_to_split.append(sheet_info.sheet_name)
            else:
                if sheet_info.is_valid:
                    sheets_to_split.append(sheet_info.sheet_name)

        if sheets_to_split:
            tasks.append((
                result.file_path,
                output_dir,
                sheets_to_split,
                result.card_number,
            ))

    if not tasks:
        return []

    workers = max_workers or os.cpu_count() or 4
    all_created: List[str] = []
    corrupted: List[str] = []

    if workers > 1 and len(tasks) > 1:
        result_files = splitter.split_many_parallel(tasks)
        all_created.extend(result_files)
    else:
        for source_path, out_dir, sheet_names, file_label in tasks:
            try:
                created = splitter.split_file(source_path, out_dir, sheet_names, file_label)
                all_created.extend(created)
            except Exception as e:
                logger.warning("Повреждённый файл при разделении %s: %s", os.path.basename(source_path), e)
                corrupted.append(source_path)

    if corrupted:
        logger.warning("Повреждённых файлов при разделении: %d", len(corrupted))

    if hasattr(cards_data, 'corrupted_files') and cards_data.corrupted_files is not None:
        cards_data.corrupted_files.extend(corrupted)
    else:
        cards_data.corrupted_files = list(corrupted)

    logger.info("Создано отдельных файлов: %d", len(all_created))
    return all_created
