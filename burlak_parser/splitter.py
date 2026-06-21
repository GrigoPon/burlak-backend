"""Модуль разделения многолистовых Excel-файлов на отдельные одностраничные файлы.

Использует метод «удаления лишнего» (а не «копирования нужного»):
  1. Загружает исходный .xlsx как ZIP-архив XML.
  2. Для каждого листа создаёт копию всего Workbook.
  3. Удаляет из копии все листы, кроме целевого.
  4. Очищает глобальные именованные диапазоны (defined names / named ranges),
     ссылающиеся на удалённые листы — это устраняет ошибку Excel
     "Removed Feature: Named range from /xl/workbook.xml part (Workbook)".

Преимущества метода:
  - 100% сохранение форматирования, стилей, картинок, шрифтов.
  - Сохраняется ширина колонок, высота строк, объединённые ячейки.
  - Сохраняются изображения, диаграммы, заморозка панелей.
  - Нет ошибки "Named range" при открытии.

Поддерживает параллелизацию через ProcessPoolExecutor.
"""

from __future__ import annotations

import io
import logging
import os
import re
import shutil
import warnings
import xml.etree.ElementTree as ET
import zipfile
from concurrent.futures import ProcessPoolExecutor, as_completed
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Set, Tuple, Union

# Подавляем предупреждения openpyxl о DrawingML (неполная поддержка)
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

logger = logging.getLogger(__name__)

# Символы, запрещённые в именах файлов Windows/Linux
_ILLEGAL_FS_CHARS_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

# Декоративные Unicode-символы, которые нужно удалять из имён файлов
# (звёздочки, ромбы, кружки, стрелки и т.д.)
_DECORATIVE_CHARS_RE = re.compile(r'[☆★●○◆◇■□▲△▼▽♠♣♥♦↗→←↑↓«»""''„]')

# Множественные подчёркивания/точки/пробелы → одинарные
_MULTI_SEP_RE = re.compile(r'[_ .]{2,}')

# Регулярка для cell reference: "A3390" → groups ("A", "3390")
_CELL_REF_RE = re.compile(r'^([A-Z]+)(\d+)$')

# Пространства имён Excel OOXML
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_PKG_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_DRAWINGML = "http://schemas.openxmlformats.org/drawingml/2006/main"
VML_NS = "urn:schemas-microsoft-com:vml"
OFFICE_NS = "urn:schemas-microsoft-com:office:office"

# Регистрируем пространства имён глобально
# Нужно зарегистрировать ВСЕ namespace-ы, которые могут встречаться
# в OOXML-файлах, чтобы избежать появления ns0:/ns1: префиксов.
# _serialize_xml() динамически переключает default namespace при каждом вызове.
ET.register_namespace('', NS_MAIN)
ET.register_namespace('r', NS_R)
ET.register_namespace('xdr', NS_DRAWING)
ET.register_namespace('a', NS_DRAWINGML)
ET.register_namespace('ct', NS_CT)


def _serialize_xml(root: ET.Element, default_ns_uri: str,
                   extra_ns: Optional[Dict[str, str]] = None) -> bytes:
    """Serialize an ET.Element with a specific default namespace.

    Saves and restores the global ET._namespace_map to avoid corruption.
    Used because different OOXML files require different default namespaces:
      - sheet XML:  NS_MAIN as default
      - Content_Types:  NS_CT as default
      - rels files:  NS_PKG_RELS as default

    CRITICAL: Uses custom XML declaration with standalone="yes", double quotes,
    and Windows-style \r\n line endings. MS Excel requires these for compatibility.
    Python 3.14: ET.tostring(standalone=True) raises TypeError, so we work around it.
    """
    old_default_uri = None
    for uri_key, prefix_val in ET._namespace_map.items():
        if prefix_val == '':
            old_default_uri = uri_key
            break
    old_extras: Dict[str, Optional[str]] = {}
    try:
        ET.register_namespace('', default_ns_uri)
        if extra_ns:
            for p, uri in extra_ns.items():
                old_extras[p] = ET._namespace_map.get(uri)
                ET.register_namespace(p, uri)
        # Serialize without declaration, then prepend MS Excel-compatible declaration
        body = ET.tostring(root, xml_declaration=False, encoding='UTF-8')
        declaration = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        return declaration + body
    finally:
        try:
            ET.register_namespace('', old_default_uri)
        except TypeError:
            pass
        if extra_ns:
            for p, uri in extra_ns.items():
                old_prefix = old_extras.get(p)
                if old_prefix is not None:
                    ET._namespace_map[uri] = old_prefix
                else:
                    ET._namespace_map.pop(uri, None)


# ─── Типы для вертикального split ───────────────────────────────────


@dataclass
class TableBoundary:
    """Границы одной таблицы (операции) внутри листа."""
    header_row: int         # Строка заголовка таблицы
    data_start: int         # Первая строка данных (header_row + 1)
    data_end: int           # Последняя строка данных
    operation_name: str = ""  # Название операции
    source_path: str = ""
    sheet_name: str = ""
    card_label: str = ""


class CardSplitter:
    """Сервис разделения многолистовых операционных карт на отдельные файлы."""

    def __init__(self, max_workers: Optional[int] = None):
        """Инициализировать сплиттер.

        Args:
            max_workers: Максимальное количество процессов для параллельного разделения.
                         По умолчанию: количество CPU.
        """
        self.max_workers = max_workers or os.cpu_count() or 4
        self.openpyxl_fallback_count = 0
        self.openpyxl_fallback_files: Set[str] = set()
        self.manifest: Dict[str, List[str]] = {}

    def split_file(
        self,
        source_path: str,
        output_dir: str,
        sheet_names: List[str],
        file_label: str = "",
    ) -> List[str]:
        """Разделить один .xlsx файл на несколько однолистовых файлов.

        Для .xls файлов (legacy) — копирует как есть без разделения.

        Args:
            source_path: Путь к исходному .xlsx/.xls файлу.
            output_dir: Директория для сохранения результатов.
            sheet_names: Имена листов, которые нужно выделить.
            file_label: Метка файла для именования выходных файлов.

        Returns:
            Список путей к созданным файлам.
        """
        os.makedirs(output_dir, exist_ok=True)
        created: List[str] = []
        original_name = os.path.basename(source_path)

        ext_lower = os.path.splitext(source_path)[1].lower()

        if ext_lower == ".xls":
            safe_label = _safe_filename(file_label)[:50] if file_label else ""
            basename = os.path.splitext(os.path.basename(source_path))[0]
            out_name = f"{safe_label}_{basename}.xls" if safe_label else f"{basename}.xls"
            output_path = os.path.join(output_dir, out_name)
            counter = 1
            while True:
                try:
                    fd = os.open(output_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                    os.close(fd)
                    break
                except FileExistsError:
                    base, ext = os.path.splitext(out_name)
                    output_path = os.path.join(output_dir, f"{base}_{counter}{ext}")
                    counter += 1
            shutil.copy2(source_path, output_path)
            created.append(output_path)
            self.manifest.setdefault(original_name, []).append(os.path.basename(output_path))
            logger.info("Скопирован .xls файл (без разделения): %s", os.path.basename(source_path))
            return created

        if ext_lower != ".xlsx":
            logger.debug("Пропуск не-.xlsx/.xls файла: %s", source_path)
            return created

        for sheet_name in sheet_names:
            safe_label = _safe_filename(file_label)[:50] if file_label else ""
            safe_sheet = _safe_filename(sheet_name)[:50]
            if safe_label:
                output_filename = f"{safe_label}_{safe_sheet}.xlsx"
            else:
                output_filename = f"{safe_sheet}.xlsx"

            output_path = os.path.join(output_dir, output_filename)

            # Skip if path already exists (pre-allocated by main thread)
            if os.path.exists(output_path):
                continue

            try:
                self._extract_sheet(source_path, output_path, sheet_name)
                created.append(output_path)
                self.manifest.setdefault(original_name, []).append(os.path.basename(output_path))
                logger.debug("Создан: %s", os.path.basename(output_path))
            except Exception as e:
                logger.warning(
                    "Ошибка разделения листа '%s' из %s: %s",
                    sheet_name, os.path.basename(source_path), e,
                )

        return created

    def split_many_parallel(
        self,
        tasks: List[Tuple[str, str, List[str], str]],
    ) -> Tuple[List[str], List[Tuple[str, str]], int, List[str], Dict[str, List[str]]]:
        """Разделить множество файлов параллельно.

        Гарантирует детерминированный порядок: результаты сортируются
        по полному пути для воспроизводимости.

        Args:
            tasks: Список кортежей (source_path, output_dir, sheet_names, file_label).

        Returns:
            Кортеж (all_created_files, errors, openpyxl_fallback_count,
                    openpyxl_fallback_files, manifest).
        """
        all_created: List[str] = []
        errors: List[Tuple[str, str]] = []
        all_openpyxl_count = 0
        all_openpyxl_files: List[str] = []
        merged_manifest: Dict[str, List[str]] = {}

        # Предвычисляем все пути детерминированно (синхронно, главный поток)
        path_map = preallocate_split_paths(tasks, tasks[0][1] if tasks else "")

        # Собираем плоские задачи (source_path, output_path, sheet_name)
        sheet_tasks: List[Tuple[str, str, str]] = []
        for source_path, out_dir, sheet_names, file_label in tasks:
            for sheet_name in sheet_names:
                output_path = path_map.get((source_path, sheet_name))
                if output_path:
                    sheet_tasks.append((source_path, output_path, sheet_name))

        with ProcessPoolExecutor(max_workers=self.max_workers) as executor:
            futures = {}
            for src, out, sheet in sheet_tasks:
                future = executor.submit(
                    _extract_to_path_worker,
                    src, out, sheet,
                )
                futures[future] = (src, out, sheet)

            for future in as_completed(futures):
                src, out, sheet = futures[future]
                try:
                    worker_result = future.result()
                    result_path = worker_result.get("path")
                    err_msg = worker_result.get("error")
                    source_basename = worker_result.get("source_basename", "")

                    if result_path:
                        all_created.append(result_path)
                        if worker_result.get("used_fallback"):
                            all_openpyxl_count += 1
                            if source_basename:
                                all_openpyxl_files.append(source_basename)
                        if source_basename:
                            merged_manifest.setdefault(source_basename, []).append(
                                os.path.basename(result_path),
                            )
                    else:
                        logger.error(
                            "Ошибка разделения %s: %s",
                            os.path.basename(src), err_msg,
                        )
                        errors.append((src, err_msg or "Unknown error"))
                except Exception as e:
                    err_msg = str(e)
                    logger.error(
                        "Критическая ошибка параллельного разделения %s: %s",
                        os.path.basename(src), err_msg,
                    )
                    errors.append((src, err_msg))

        # Сортируем для детерминированного порядка
        all_created.sort()
        errors.sort(key=lambda x: x[0])
        all_openpyxl_files.sort()

        # Сортируем значения в манифесте
        for orig in merged_manifest:
            merged_manifest[orig].sort()

        return all_created, errors, all_openpyxl_count, all_openpyxl_files, merged_manifest

    def _extract_sheet(
        self, source_path: str, output_path: str, keep_sheet_name: str,
    ) -> None:
        """Выделить один лист из .xlsx файла.

        Стратегия ЛИНЕЙНАЯ (без рекурсии), приоритет производительности:
          1. ZIP-метод: быстрый (миллисекунды), обрабатывает 90%+ файлов.
          2. При ошибке ZIP → ОДНА попытка openpyxl (медленный, для WPS/битых).
          3. При ошибке openpyxl → исключение.

        Args:
            source_path: Путь к исходному .xlsx файлу.
            output_path: Путь для сохранения нового .xlsx файла.
            keep_sheet_name: Имя листа, который нужно оставить.

        Raises:
            ValueError: Если целевой лист не найден в файле.
            Exception: Если оба метода завершились ошибкой.
        """
        # Попытка 1: ZIP (быстро — миллисекунды на файл)
        try:
            self._extract_sheet_via_zip(source_path, output_path, keep_sheet_name)
            if not _validate_split_file(output_path):
                try:
                    os.remove(output_path)
                except OSError:
                    pass
                raise ValueError(f"Invalid split output: {output_path}")
            return
        except Exception as e:
            logger.warning(
                "ZIP-метод не смог разделить %s: %s. Пробуем openpyxl...",
                os.path.basename(source_path), e,
            )

        # Попытка 2: openpyxl (медленно — для WPS/битых файлов, одна попытка)
        try:
            self._extract_sheet_via_openpyxl(source_path, output_path, keep_sheet_name)
            self.openpyxl_fallback_count += 1
            self.openpyxl_fallback_files.add(os.path.basename(source_path))
            logger.info(
                "openpyxl успешно разделил лист: %s в файле %s",
                keep_sheet_name, os.path.basename(source_path),
            )
        except Exception as openpyxl_e:
            logger.error(
                "Оба метода разделения листа '%s' из %s завершились ошибкой. "
                "ZIP: см. выше. openpyxl: %s",
                keep_sheet_name, os.path.basename(source_path), openpyxl_e,
            )
            raise

    def _extract_sheet_via_openpyxl(
        self, source_path: str, output_path: str, keep_sheet_name: str,
    ) -> None:
        """Выделить один лист через openpyxl (load → remove sheets → save).

        Этот метод корректно обрабатывает файлы, созданные WPS Office
        и другими генераторами OOXML, которые могут содержать
        нестандартные CRC-суммы или повреждённые записи ZIP.

        Включает обход бага WPS: DefinedNameDict без атрибута definedName.

        Args:
            source_path: Путь к исходному .xlsx файлу.
            output_path: Путь для сохранения нового .xlsx файла.
            keep_sheet_name: Имя листа, который нужно оставить.

        Raises:
            ValueError: Если целевой лист не найден.
            Exception: При ошибке загрузки/сохранения openpyxl.
        """
        import openpyxl

        with warnings.catch_warnings():
            warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
            wb = openpyxl.load_workbook(source_path)
        sheet_names = wb.sheetnames

        if keep_sheet_name not in sheet_names:
            wb.close()
            raise ValueError(f"Лист '{keep_sheet_name}' не найден в файле")

        if len(sheet_names) <= 1:
            wb.save(output_path)
            wb.close()
            return

        sheets_to_remove = [n for n in sheet_names if n != keep_sheet_name]

        # ── WPS BUG FIX: Патчим DefinedNameDict перед удалением листов ──
        # WPS Office создаёт повреждённые OOXML, где wb.defined_names
        # не имеет атрибута definedName. openpyxl падает при del wb[sheet]
        # с AttributeError: 'DefinedNameDict' object has no attribute 'definedName'.
        # Решение: принудительно создаём пустые атрибуты.
        dn = getattr(wb, 'defined_names', None)
        if dn is not None:
            if not hasattr(dn, 'definedName'):
                dn.definedName = []
            if not hasattr(dn, 'elements'):
                dn.elements = []

        # Удаляем named ranges, ссылающиеся на удаляемые листы
        try:
            if dn is not None and dn.definedName:
                to_delete = []
                for defined_name in dn.definedName:
                    attr_text = getattr(defined_name, 'attr_text', None) or str(defined_name)
                    for deleted in sheets_to_remove:
                        if deleted in attr_text or f"'{deleted}'" in attr_text:
                            to_delete.append(defined_name)
                            break
                for defined_name in to_delete:
                    try:
                        dn.definedName.remove(defined_name)
                    except Exception as e:
                        logger.debug("Failed to remove defined name: %s", e)
        except Exception as e:
            logger.debug("Defined name cleanup failed (non-critical): %s", e)

        for name in sheets_to_remove:
            try:
                del wb[name]
            except Exception as e:
                logger.debug("Failed to remove sheet %s: %s", name, e)

        with warnings.catch_warnings():
            warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
            wb.save(output_path)
        wb.close()

    def _extract_sheet_via_zip(
        self, source_path: str, output_path: str, keep_sheet_name: str,
    ) -> None:
        """Выделить один лист через чистую ZIP-манипуляцию.

        АЛГОРИТМ (КЛЮЧЕВОЙ):
        Стратегия «сохраняем только нужное»:
          1. Читаем оригинальный ZIP в память.
          2. Находим rId и путь к сохранённому листу.
          3. Рекурсивно трассируем все .rels от листа → находим все нужные файлы
             (drawing XML, VML, OLE, изображения, printerSettings, их .rels).
          4. Добавляем обязательные: workbook, styles, theme, sharedStrings, docProps.
          5. Создаём НОВЫЙ workbook.xml: только 1 лист + очищенные definedNames.
             ВАЖНО: все остальные элементы (fileVersion, workbookPr, bookViews,
             calcPr, AlternateContent) копируются из оригинала AS-IS.
          6. Создаём НОВЫЙ workbook.xml.rels: только лист + shared items.
          7. Фильтруем Content_Types.xml: только Override для существующих файлов.
          8. Записываем новый ZIP.

        ВАЖНО: НИКАКОГО openpyxl, НИКАКОГО переименования файлов!
        Все оригинальные XML-файлы и бинарные данные копируются AS-IS
        с их оригинальными именами (sheet8.xml остаётся sheet8.xml).
        Это гарантирует 100% сохранение всех ссылок внутри drawing,
        VML, OLE и других файлов.

        Args:
            source_path: Путь к исходному .xlsx файлу.
            output_path: Путь для сохранения нового .xlsx файла.
            keep_sheet_name: Имя листа, который нужно оставить.

        Raises:
            ValueError: Если целевой лист не найден в файле.
        """
        # ── ФАЗА 1: Прочитать оригинальный ZIP ──
        with open(source_path, 'rb') as f:
            zip_data = f.read()

        orig_entries: Dict[str, bytes] = {}
        try:
            with zipfile.ZipFile(io.BytesIO(zip_data), 'r') as zf:
                for name in zf.namelist():
                    try:
                        orig_entries[name] = zf.read(name)
                    except (zipfile.BadZipFile, Exception):
                        pass
        except zipfile.BadZipFile as e:
            raise ValueError(f"Cannot read source ZIP: {e}")

        # ── ФАЗА 2: Найти лист в workbook.xml ──
        wb_xml = orig_entries.get('xl/workbook.xml')
        if wb_xml is None:
            raise ValueError("xl/workbook.xml not found")

        wb_root = ET.fromstring(wb_xml)
        sheets_elem = wb_root.find(f'{{{NS_MAIN}}}sheets')
        if sheets_elem is None:
            raise ValueError("<sheets> not found in original workbook.xml")

        # Находим rId сохранённого листа
        target_r_id: Optional[str] = None
        for sheet_el in sheets_elem.findall(f'{{{NS_MAIN}}}sheet'):
            if sheet_el.get('name') == keep_sheet_name:
                target_r_id = sheet_el.get(f'{{{NS_R}}}id') or sheet_el.get('r:id')
                break

        if target_r_id is None:
            raise ValueError(f"Sheet '{keep_sheet_name}' not found")

        # ── ФАЗА 3a: Найти путь к листу из workbook.xml.rels ──
        rels_xml = orig_entries.get('xl/_rels/workbook.xml.rels')
        if rels_xml is None:
            raise ValueError("xl/_rels/workbook.xml.rels not found")

        rels_root = ET.fromstring(rels_xml)
        orig_sheet_path = ''
        for rel_el in rels_root:
            if rel_el.get('Id') == target_r_id:
                orig_sheet_path = rel_el.get('Target', '')
                break

        if not orig_sheet_path:
            raise ValueError(f"No target for rId {target_r_id}")

        # Нормализуем путь
        orig_sheet_path = orig_sheet_path.lstrip('/')
        if not orig_sheet_path.startswith('xl/'):
            orig_sheet_path = 'xl/' + orig_sheet_path

        # ── ФАЗА 3b: Рекурсивно трассировать все .rels ──
        needed: Set[str] = set()

        def _trace_rels(rels_path: str, base_dir: str) -> None:
            """Рекурсивно трассировать .rels, добавляя все найденные файлы."""
            if rels_path not in orig_entries:
                return
            try:
                tr_root = ET.fromstring(orig_entries[rels_path])
                for tr_el in tr_root:
                    target = tr_el.get('Target', '')
                    if not target:
                        continue
                    # Ресолвим относительный путь от base_dir
                    resolved = os.path.normpath(
                        os.path.join(base_dir, target)
                    ).replace(os.sep, '/')
                    if resolved in orig_entries and resolved not in needed:
                        needed.add(resolved)
                        # Ищем под-rels (drawing.rels, vml.rels)
                        res_dir = os.path.dirname(resolved)
                        res_base = os.path.basename(resolved)
                        sub_rels = f"{res_dir}/_rels/{res_base}.rels"
                        if sub_rels in orig_entries:
                            needed.add(sub_rels)
                            _trace_rels(sub_rels, res_dir)
            except Exception as e:
                logger.debug("Trace rels failed for %s: %s", rels_path, e)

        # Всегда нужны базовые файлы
        needed.add('[Content_Types].xml')
        needed.add('_rels/.rels')
        needed.add('xl/workbook.xml')
        needed.add('xl/_rels/workbook.xml.rels')

        # Сам лист
        needed.add(orig_sheet_path)

        # .rels файл листа и его рекурсивные зависимости
        sheet_dir = os.path.dirname(orig_sheet_path)
        sheet_base = os.path.basename(orig_sheet_path)
        sheet_rels_path = f"{sheet_dir}/_rels/{sheet_base}.rels"
        if sheet_rels_path in orig_entries:
            needed.add(sheet_rels_path)
            _trace_rels(sheet_rels_path, sheet_dir)

        # Добавляем shared items (styles, theme, sharedStrings) из workbook.xml.rels
        for rel_el in rels_root:
            rel_id = rel_el.get('Id', '')
            rel_type = rel_el.get('Type', '')
            rel_target = rel_el.get('Target', '')
            if rel_id == target_r_id:
                continue  # Пропускаем сам лист (уже добавлен)
            # Добавляем styles, theme, sharedStrings
            if ('styles' in rel_type.lower()
                    or 'theme' in rel_type.lower()
                    or 'sharedstrings' in rel_type.lower()):
                resolved = os.path.normpath(
                    os.path.join('xl', rel_target)
                ).replace(os.sep, '/')
                if resolved in orig_entries:
                    needed.add(resolved)

        # Добавляем docProps (core, app, custom) — не влияют на загрузку листа
        doc_props = [n for n in orig_entries if n.startswith('docProps/')]
        needed.update(doc_props)

        # Добавляем customXml (если есть)
        custom_xml = [n for n in orig_entries if n.startswith('customXml/')]
        needed.update(custom_xml)

        # ── ФАЗА 4: Собрать имена удалённых листов ──
        other_sheet_names: Set[str] = set()
        for sheet_el in sheets_elem.findall(f'{{{NS_MAIN}}}sheet'):
            sn = sheet_el.get('name', '')
            if sn != keep_sheet_name:
                other_sheet_names.add(sn)

        # ── ФАЗА 5: Модифицировать workbook.xml через строковые операции ──
        # ВАЖНО: используем строковые операции, а НЕ XML парсинг,
        # чтобы сохранить оригинальные namespace declarations, XML declaration,
        # line endings и все остальные детали исходного файла AS-IS.
        new_wb_text = _modify_workbook_xml_text(
            orig_entries['xl/workbook.xml'].decode('utf-8'),
            keep_sheet_name,
            target_r_id,
            other_sheet_names,
        )

        # ── ФАЗА 6: Модифицировать workbook.xml.rels — удалить лишние Relationship ──
        new_rels_text = _modify_workbook_rels_text(
            orig_entries['xl/_rels/workbook.xml.rels'].decode('utf-8'),
            target_r_id,
        )

        # ── ФАЗА 7: Собрать выходной словарь ──
        output_entries: Dict[str, bytes] = {}

        for name in needed:
            if name == 'xl/workbook.xml':
                output_entries[name] = new_wb_text.encode('utf-8')
            elif name == 'xl/_rels/workbook.xml.rels':
                output_entries[name] = new_rels_text.encode('utf-8')
            else:
                output_entries[name] = orig_entries[name]

        # ── ФАЗА 8: Фильтровать Content_Types.xml — удалить Override для отсутствующих файлов ──
        if '[Content_Types].xml' in needed:
            ct_text = orig_entries['[Content_Types].xml'].decode('utf-8')
            new_ct_text = _filter_content_types_text(ct_text, set(output_entries.keys()))
            output_entries['[Content_Types].xml'] = new_ct_text.encode('utf-8')

        # ── ФАЗА 8: Записать новый ZIP с сохранением оригинального сжатия ──
        # ВАЖНО: MS Excel требует, чтобы изображения (PNG, EMF, JPEG) были
        # STORED (без сжатия), а XML/DATA файлы — DEFLATED.
        # Используем оригинальный compression_type если известен.
        if os.path.exists(output_path):
            os.remove(output_path)

        def _get_compress_type(name: str) -> int:
            """Определить метод сжатия: STORED для изображений, DEFLATED для всего остального.

            MS Office хранит изображения в исходном виде (STORED), так как они
            уже сжаты. XML и другие текстовые данные — DEFLATED.
            """
            name_lower = name.lower()
            # Изображения — без сжатия (уже сжаты, DEFLATE не помогает)
            if any(name_lower.endswith(ext) for ext in ['.png', '.emf', '.wmf', '.jpeg', '.jpg',
                                                         '.gif', '.tiff', '.tif', '.bmp', '.svg']):
                return zipfile.ZIP_STORED
            return zipfile.ZIP_DEFLATED

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for name in sorted(output_entries.keys()):
                compress_type = _get_compress_type(name)
                zout.writestr(name, output_entries[name], compress_type=compress_type)


_CT_CACHE: Dict[str, Optional[str]] = {}

def _modify_workbook_xml_text(
    xml_text: str,
    keep_sheet_name: str,
    target_r_id: str,
    other_sheet_names: Set[str],
) -> str:
    """Модифицировать workbook.xml строковыми операциями.

    1. Удалить лишние <sheet> из <sheets>.
    2. Удалить <definedName>, ссылающиеся на удалённые листы.

    ВСЁ остальное сохраняется AS-IS (XML declaration, namespace, line endings).
    """
    # ── 1. Замена <sheets> — оставляем только 1 лист ──
    def _replace_sheets(m: re.Match) -> str:
        """Callback для замены содержимого <sheets>."""
        open_tag = m.group(1)
        close_tag = m.group(3)
        content = m.group(2)
        # Ищем сохранённый лист по name или r:id
        kept = None
        for sh in re.finditer(r'<sheet[^>]*/>', content):
            sh_tag = sh.group(0)
            name_m = re.search(r'name="([^"]+)"', sh_tag)
            if name_m and name_m.group(1) == keep_sheet_name:
                kept = sh_tag
                break
        # Fallback: по r:id
        if kept is None:
            for sh in re.finditer(r'<sheet[^>]*/>', content):
                sh_tag = sh.group(0)
                rid_m = re.search(r'r:id="([^"]+)"', sh_tag)
                if rid_m and rid_m.group(1) == target_r_id:
                    kept = sh_tag
                    break
        if kept:
            return f'{open_tag}\n{kept}\n{close_tag}'
        return m.group(0)  # fallback: без изменений

    xml_text = re.sub(r'(<sheets[^>]*>)(.*?)(</sheets>)', _replace_sheets, xml_text, count=1, flags=re.DOTALL)

    # ── 2. Очистка definedNames ──
    def _filter_defined_names(m: re.Match) -> str:
        """Callback: удалить definedName, ссылающиеся на other_sheet_names."""
        dn_block = m.group(0)
        # Находим границы тега
        dn_open_m = re.match(r'(<definedNames[^>]*>)', dn_block)
        if not dn_open_m:
            return dn_block
        dn_open = dn_open_m.group(1)
        # Находим закрывающий тег
        close_idx = dn_block.rfind('</definedNames>')
        if close_idx == -1:
            return dn_block
        content = dn_block[len(dn_open):close_idx]

        kept_lines = []
        for dn_match_inner in re.finditer(r'<definedName[^>]*>.*?</definedName>', content, re.DOTALL):
            dn_xml = dn_match_inner.group(0)
            dn_text = re.sub(r'<[^>]+>', '', dn_xml).strip()  # extract text content
            formula = dn_text
            should_remove = False
            for deleted_name in other_sheet_names:
                if f"'{deleted_name}'!" in formula or formula.startswith(f"{deleted_name}!"):
                    should_remove = True
                    break
            if not should_remove:
                # Обновляем localSheetId на 0 (сохранённый лист теперь единственный)
                dn_xml = re.sub(r'localSheetId="[^"]+"', 'localSheetId="0"', dn_xml)
                kept_lines.append(dn_xml)

        if not kept_lines:
            # definedNames пуст — удаляем весь блок
            return ''
        return dn_open + ''.join(kept_lines) + '</definedNames>'

    xml_text = re.sub(r'<definedNames[^>]*>.*?</definedNames>', _filter_defined_names, xml_text, count=1, flags=re.DOTALL)

    return xml_text


def _modify_workbook_rels_text(
    rels_text: str,
    target_r_id: str,
) -> str:
    """Модифицировать workbook.xml.rels — удалить Relationship для других листов.

    Оставляет ТОЛЬКО:
      - worksheet (сохранённый лист)
      - styles
      - theme
      - sharedStrings

    ВСЁ остальное (XML declaration, форматирование) сохраняется AS-IS.
    Используется re.sub для удаления отдельных <Relationship .../> строк.
    """
    def _keep_relevant_rels(m: re.Match) -> str:
        """Callback: вернуть Relationship строку только если она нужна."""
        rel = m.group(0)
        rid_m = re.search(r'Id="([^"]+)"', rel)
        rtype_m = re.search(r'Type="([^"]+)"', rel)
        rid = rid_m.group(1) if rid_m else ''
        rtype = rtype_m.group(1).lower() if rtype_m else ''

        # Всегда оставляем сохранённый лист
        if rid == target_r_id:
            return rel
        # Оставляем shared items
        if any(st in rtype for st in ['styles', 'theme', 'sharedstrings']):
            return rel
        # Удаляем всё остальное
        return ''

    return re.sub(r'<Relationship[^>]*/>', _keep_relevant_rels, rels_text)


def _filter_content_types_text(
    ct_text: str,
    existing_files: Set[str],
) -> str:
    """Фильтровать [Content_Types].xml — удалить Override для несуществующих файлов.

    Args:
        ct_text: Оригинальный текст [Content_Types].xml.
        existing_files: Множество путей файлов в выходном ZIP.

    Returns:
        Отфильтрованный XML текст (ВСЁ остальное AS-IS).
    """
    def _filter_override(m: re.Match) -> str:
        """Callback: вернуть Override только если файл существует."""
        override_line = m.group(0)
        pn_m = re.search(r'PartName="([^"]+)"', override_line)
        if pn_m:
            part_name = pn_m.group(1)
            if part_name.startswith('/'):
                clean_name = part_name[1:]
            else:
                clean_name = part_name
            if clean_name not in existing_files:
                return ''  # Удаляем
        return override_line

    return re.sub(r'<Override[^>]*/>', _filter_override, ct_text)


def _infer_content_type(path: str) -> Optional[str]:
    """Определить OOXML ContentType по пути файла."""
    if path in _CT_CACHE:
        return _CT_CACHE[path]

    result: Optional[str] = None
    path_lower = path.lower()

    if path_lower.endswith('.xml'):
        if 'drawing' in path_lower and 'rels' not in path_lower:
            result = 'application/vnd.openxmlformats-officedocument.drawing+xml'
        elif 'vml' in path_lower:
            result = 'application/vnd.openxmlformats-officedocument.vmlDrawing'
    elif path_lower.endswith('.bin'):
        result = 'application/vnd.openxmlformats-officedocument.oleObject'
    elif path_lower.endswith('.rels'):
        result = 'application/vnd.openxmlformats-package.relationships+xml'
    elif path_lower.endswith('.png'):
        result = 'image/png'
    elif path_lower.endswith('.jpeg') or path_lower.endswith('.jpg'):
        result = 'image/jpeg'
    elif path_lower.endswith('.emf'):
        result = 'image/x-emf'
    elif path_lower.endswith('.wmf'):
        result = 'image/x-wmf'
    elif path_lower.endswith('.gif'):
        result = 'image/gif'
    elif path_lower.endswith('.tiff') or path_lower.endswith('.tif'):
        result = 'image/tiff'
    elif path_lower.endswith('.bmp'):
        result = 'image/bmp'
    elif path_lower.endswith('.svg'):
        result = 'image/svg+xml'

    _CT_CACHE[path] = result
    return result


def _validate_split_file(path: str) -> bool:
    """Verify a split .xlsx file has valid sheet XML and can be opened.

    Returns True if the file is valid, False if it should be deleted.
    """
    try:
        with zipfile.ZipFile(path, 'r') as zf:
            has_sheet = False
            for name in zf.namelist():
                if (name.endswith('.xml')
                        and 'sheet' in name.lower()
                        and '_rels' not in name):
                    data = zf.read(name)
                    root = ET.fromstring(data)
                    ns = f'{{{NS_MAIN}}}sheetData'
                    if root.find(ns) is None:
                        logger.warning(
                            "Invalid split file %s: missing sheetData in %s",
                            os.path.basename(path), name)
                        return False
                    has_sheet = True
                    break
            return has_sheet
    except (zipfile.BadZipFile, ET.ParseError, OSError) as e:
        logger.warning("Invalid split file %s: %s", os.path.basename(path), e)
        return False


def _safe_filename(name: str) -> str:
    """Очистить имя файла, сохранив Unicode-символы.

    Удаляет:
      - Символы, запрещённые в именах файлов ОС: < > : " / \\ | ? *
      - Управляющие символы (0x00-0x1f)
      - Декоративные Unicode: ☆ ★ ● ○ ◆ ◇ ■ □ и т.д.
      - Множественные подчёркивания/точки/пробелы → одинарные

    Args:
        name: Исходное имя файла.

    Returns:
        Безопасное имя файла с сохранёнными кириллицей/иероглифами.
    """
    result = _ILLEGAL_FS_CHARS_RE.sub("_", name)
    result = _DECORATIVE_CHARS_RE.sub("", result)
    result = _MULTI_SEP_RE.sub("_", result)
    return result.strip("_ .")


def _collect_related_files(
    zip_entries: Dict[str, bytes],
    removed_sheet: str,
    files_to_remove: Set[str],
) -> None:
    """Собрать все файлы, связанные с удаляемым листом (.rels, drawings, VML, charts).

    Args:
        zip_entries: Словарь {имя_в_zip: содержимое}.
        removed_sheet: Путь к удаляемому листу (напр. 'xl/worksheets/sheet2.xml').
        files_to_remove: Множество для добавления найденных файлов.
    """
    # .rels файл для листа
    base = os.path.basename(removed_sheet)
    removed_rels = f"xl/worksheets/_rels/{base}.rels"
    if removed_rels in zip_entries:
        files_to_remove.add(removed_rels)

        # Находим связанные drawings, VML, charts
        try:
            sr_root = ET.fromstring(zip_entries[removed_rels])
            # OOXML relationship targets resolve relative to the package part
            # (e.g. xl/worksheets/), NOT relative to the .rels directory
            # (e.g. xl/worksheets/_rels/). Go up one extra level.
            sr_dir = os.path.dirname(os.path.dirname(removed_rels))
            for sr_el in sr_root:
                sr_target = sr_el.get('Target', '')
                # Резолвим относительный путь (../drawings/drawing1.xml)
                resolved = os.path.normpath(os.path.join(sr_dir, sr_target))
                resolved = resolved.replace(os.sep, '/')
                files_to_remove.add(resolved)

                # Рекурсивно: .rels для drawing, VML
                resolved_base = os.path.basename(resolved)
                resolved_dir = os.path.dirname(resolved)
                resolved_rels = f"{resolved_dir}/_rels/{resolved_base}.rels"
                if resolved_rels in zip_entries:
                    files_to_remove.add(resolved_rels)
        except Exception as e:
            logger.debug("Relationship resolution failed: %s", e)


def _clean_named_ranges(
    wb_root: ET.Element,
    deleted_sheet_names: Set[str],
    keep_sheet_name: str,
) -> None:
    """Очистить definedNames (named ranges), ссылающиеся на удалённые листы.

    Это КЛЮЧЕВОЙ шаг для устранения ошибки Excel:
    "Removed Feature: Named range from /xl/workbook.xml part (Workbook)"

    Алгоритм:
      1. Найти элемент <definedNames> в workbook.xml.
      2. Для каждого <definedName> проверить, ссылается ли он на удалённый лист.
      3. Ссылка на лист в definedName обычно выглядит как: SheetName!$A$1
         или заключена в одиночные кавычки если имя с пробелами: 'Sheet Name'!$A$1.
      4. Удалить все definedNames, ссылающиеся на удалённые листы.

    Args:
        wb_root: Корневой элемент workbook.xml.
        deleted_sheet_names: Множество имён удалённых листов.
        keep_sheet_name: Имя оставленного листа.
    """
    defined_names_elem = wb_root.find(f'{{{NS_MAIN}}}definedNames')
    if defined_names_elem is None:
        return  # Нет именованных диапазонов — нечего чистить

    names_to_remove: List[ET.Element] = []

    for dn in defined_names_elem.findall(f'{{{NS_MAIN}}}definedName'):
        # Текст definedName — это формула со ссылкой на лист
        formula = (dn.text or '').strip()
        name_attr = dn.get('name', '')

        # Проверяем, ссылается ли definedName на удалённый лист
        # Шаблоны ссылок:
        #   'Sheet Name'!$A$1:$B$2
        #   SheetName!$A$1
        #   SheetName!$A$1:$B$2
        should_remove = False

        for deleted_name in deleted_sheet_names:
            # Проверка с кавычками (для имён с пробелами/спецсимволами)
            if f"'{deleted_name}'!" in formula:
                should_remove = True
                break
            # Проверка без кавычек
            if formula.startswith(f"{deleted_name}!"):
                should_remove = True
                break
            # Проверка на вхождение (менее точная, но покрывает edge cases)
            # Ищем паттерн: граница слова + имя листа + !
            if re.search(rf"\b{re.escape(deleted_name)}!", formula):
                should_remove = True
                break

        # Также проверяем локальные имена (localSheetId атрибут)
        if not should_remove:
            local_sheet_id = dn.get('localSheetId')
            if local_sheet_id is not None and local_sheet_id != '0':
                # После удаления всех остальных листов, оставшийся лист
                # становится единственным с индексом 0.
                # Обновляем localSheetId на 0.
                dn.set('localSheetId', '0')

        if should_remove:
            names_to_remove.append(dn)
            logger.debug(
                "Удалён definedName '%s' (ссылка на удалённый лист)",
                name_attr,
            )

    for dn in names_to_remove:
        defined_names_elem.remove(dn)

    # Если после очистки definedNames пуст — удаляем элемент целиком
    if len(defined_names_elem) == 0:
        wb_root.remove(defined_names_elem)


def preallocate_split_paths(
    tasks: List[Tuple[str, str, List[str], str]],
    output_dir: str,
) -> Dict[Tuple[str, str], str]:
    """Детерминированная предварительная разметка путей для всех листов.

    ВЫПОЛНЯЕТСЯ В ГЛАВНОМ ПОТОКЕ (один поток, детерминированно).

    Алгоритм:
      1. Собрать все задачи (источник + лист + метка) в плоский список.
      2. Отсортировать по (source_path, sheet_name) — детерминированный порядок.
      3. Для каждой задачи вычислить целевой путь.
      4. При коллизии имён — разрешить последовательно (_1, _2, ...).
         Так как список отсортирован, разрешение 100% детерминированно.
      5. Вернуть словарь {(source_path, sheet_name) -> abs_output_path}.

    Args:
        tasks: Список кортежей (source_path, output_dir, sheet_names, file_label)
               — такой же формат, как в split_many_parallel.
        output_dir: Директория для сохранения результатов.

    Returns:
        Словарь, отображающий (source_path, sheet_name) в уникальный
        абсолютный путь выходного файла.
    """
    path_registry: Set[str] = set()
    path_map: Dict[Tuple[str, str], str] = {}

    # 1. Собираем плоский список (source_path, sheet_name, file_label)
    sheet_tasks: List[Tuple[str, str, str]] = []
    for source_path, _out_dir, sheet_names, file_label in tasks:
        for sheet_name in sheet_names:
            sheet_tasks.append((source_path, sheet_name, file_label or ""))

    # 2. Детерминированная сортировка
    sheet_tasks.sort(key=lambda t: (t[0], t[1], t[2]))

    # 3-4. Предвычисляем пути с детерминированным разрешением коллизий
    for source_path, sheet_name, file_label in sheet_tasks:
        safe_label = _safe_filename(file_label)[:50] if file_label else ""
        safe_sheet = _safe_filename(sheet_name)[:50]
        if safe_label:
            output_filename = f"{safe_label}_{safe_sheet}.xlsx"
        else:
            output_filename = f"{safe_sheet}.xlsx"

        output_path = os.path.join(output_dir, output_filename)
        base_no_ext = os.path.splitext(output_filename)[0]
        ext = ".xlsx"

        # Детерминированное разрешение коллизий (проверка по set, не по файловой системе)
        counter = 1
        while output_path in path_registry:
            output_path = os.path.join(output_dir, f"{base_no_ext}_{counter}{ext}")
            counter += 1

        path_registry.add(output_path)
        path_map[(source_path, sheet_name)] = output_path

    return path_map


def _verify_xlsx_integrity(file_path: str) -> Tuple[bool, str]:
    """Проверить целостность .xlsx файла (ЛЕНЬЯНАЯ проверка, как Microsoft Excel).

    Файл считается ПОВРЕЖДЁННЫМ (вернёт False) только если openpyxl не может
    загрузить его даже в read_only режиме — т.е. при фатальных исключениях:
      - zipfile.BadZipFile (архив не является ZIP)
      - InvalidFileException (битая OOXML-структура)

    Args:
        file_path: Путь к .xlsx файлу.

    Returns:
        Кортеж (is_valid: bool, error_message: str).
        error_message пуст если файл корректен.
    """
    import warnings

    try:
        with warnings.catch_warnings():
            warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
            warnings.filterwarnings('ignore', category=DeprecationWarning)
            wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            _ = wb.sheetnames
            wb.close()
    except zipfile.BadZipFile as e:
        return False, f"BadZipFile: {e}"
    except Exception as e:
        exc_name = type(e).__name__
        if exc_name in ("InvalidFileException", "InvalidFormatException", "LoadWorkbookException"):
            return False, f"{exc_name}: {e}"
        # All other exceptions are non-fatal (warnings, DrawingML, etc.)
        return True, ""

    return True, ""





def _extract_to_path_worker(
    source_path: str,
    output_path: str,
    sheet_name: str,
) -> Dict[str, Any]:
    """Рабочая функция: выделить один лист в предварительно размеченный путь.

    Выполняется в отдельном процессе. НЕ проверяет существование файла —
    уникальность пути гарантирована главным потоком через preallocate_split_paths.

    Вызывает _extract_sheet(), который пробует ZIP-метод, затем openpyxl fallback.
    Если оба метода завершаются ошибкой — файл считается повреждённым.

    Args:
        source_path: Путь к исходному .xlsx файлу.
        output_path: Абсолютный путь для сохранения (уже гарантированно уникальный).
        sheet_name: Имя листа для выделения.

    Returns:
        Словарь с результатами:
          - "path": output_path при успехе, None при ошибке
          - "error": сообщение об ошибке или None
          - "used_fallback": True если использован openpyxl fallback
          - "source_basename": os.path.basename(source_path)
    """
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    splitter = CardSplitter(max_workers=1)
    result: Dict[str, Any] = {
        "path": None,
        "error": None,
        "used_fallback": False,
        "source_basename": os.path.basename(source_path),
        "source_path": source_path,
        "sheet_name": sheet_name,
    }
    try:
        splitter._extract_sheet(source_path, output_path, sheet_name)
        if splitter.openpyxl_fallback_count > 0:
            result["used_fallback"] = True
        result["path"] = output_path
    except Exception as e:
        err_msg = f"Лист '{sheet_name}' из {os.path.basename(source_path)}: {e}"
        logger.error("Ошибка разделения: %s", err_msg)
        result["error"] = err_msg
    return result




def find_table_boundaries(
    source_path: str,
    sheet_name: str,
) -> List[TableBoundary]:
    """Обнаружить границы таблиц (операций) внутри одного листа.

    Для SWM-карт, где один лист содержит несколько операций,
    каждая со своим заголовком и данными.

    Args:
        source_path: Путь к .xlsx файлу.
        sheet_name: Имя листа для анализа.

    Returns:
        Список TableBoundary с границами каждой таблицы.
    """
    from burlak_parser.heuristic_analyzer import HeuristicAnalyzer
    from burlak_parser.card_parser import ExcelReader

    boundaries: List[TableBoundary] = []

    reader = ExcelReader(source_path)
    try:
        if sheet_name not in reader.sheet_names:
            return boundaries

        ws = reader.get_sheet(sheet_name)
        start_search = 1
        max_row = ws.max_row or 0

        # ── GUARD: NEVER trigger vertical split on sheets with <= 500 rows ──
        # Standard Jetour/Changan cards are small (20-200 rows), and vertical
        # splitting them deletes the original single-sheet file, losing data.
        # Only SWM megasheets (1000-8000 rows) should be vertically split.
        if max_row <= 500:
            return boundaries

        max_tables = 500

        for table_idx in range(max_tables):
            if start_search >= max_row:
                break

            table_info = HeuristicAnalyzer.find_part_table(ws, start_row=start_search)
            if table_info is None:
                break

            header_row, part_no_col, qty_col, name_col = table_info

            if header_row < start_search:
                break

            # False-positive check: reject if qty_col not found (real tables need qty)
            if qty_col is None or qty_col <= 0:
                start_search = header_row + 1
                continue

            # False-positive check: skip if header row has too few non-empty cells
            # (change-record rows like "标记 | 处数 | 更改文件号" have few meaningful cells)
            header_non_empty = 0
            max_check_col = min((ws.max_column or 10) + 1, 50)
            for hc in range(1, max_check_col):
                hv = ws.cell_value(header_row, hc)
                if hv is not None and str(hv).strip():
                    header_non_empty += 1
            if header_non_empty < 3:
                start_search = header_row + 1
                continue

            # Определяем operation_name
            operation_name = HeuristicAnalyzer.extract_operation_name(ws, header_row)

            # Определяем последнюю строку данных (data_end)
            data_end = _find_table_data_end(ws, header_row, max_row, part_no_col)

            boundaries.append(TableBoundary(
                header_row=header_row,
                data_start=header_row + 1,
                data_end=data_end,
                operation_name=operation_name,
                source_path=source_path,
                sheet_name=sheet_name,
                card_label=f"{table_idx + 1:03d}_{_safe_filename(operation_name)[:30]}" if operation_name else f"Op{table_idx + 1:03d}",
            ))

            start_search = data_end + 1

    finally:
        reader.close()

    # Fallback: обнаружение таблиц проверки качества (检验项目 pattern)
    # Если стандартные таблицы деталей не найдены, ищем повторяющиеся
    # блоки с заголовком "检验项目" в колонке B каждые ~20 строк.
    if not boundaries:
        boundaries = _detect_inspection_boundaries(source_path, sheet_name)

    return boundaries


# Ключевые слова для обнаружения таблиц проверки качества
_INSPECTION_HEADER_KW = '检验项目'
_INSPECTION_SUBHEADER_KW = '作业内容图示'


def _detect_inspection_boundaries(
    source_path: str,
    sheet_name: str,
) -> List[TableBoundary]:
    """Обнаружить границы таблиц проверки качества (检验作业指导书).

    Ищет повторяющиеся блоки с заголовком "检验项目" в колонке B.
    Каждый блок содержит операцию проверки качества.

    Args:
        source_path: Путь к .xlsx файлу.
        sheet_name: Имя листа.

    Returns:
        Список TableBoundary для каждой операции проверки.
    """
    from burlak_parser.card_parser import ExcelReader

    boundaries: List[TableBoundary] = []
    reader = ExcelReader(source_path)
    try:
        if sheet_name not in reader.sheet_names:
            return boundaries

        ws = reader.get_sheet(sheet_name)
        max_row = ws.max_row or 0
        if max_row < 3:
            return boundaries

        # Находим все строки с "检验项目" в колонке B (col 2)
        header_rows: List[int] = []
        for r in range(1, max_row + 1):
            val = ws.cell_value(r, 2)
            if val is not None and _INSPECTION_HEADER_KW in str(val):
                header_rows.append(r)

        if len(header_rows) < 2:
            return boundaries

        # Определяем шаг между заголовками (медиана интервалов)
        spacings = [header_rows[i + 1] - header_rows[i]
                    for i in range(len(header_rows) - 1)]
        if not spacings:
            return boundaries
        step = sorted(spacings)[len(spacings) // 2]  # медиана

        # Проверяем что шаг стабилен (>50% интервалов в пределах ±3 от медианы)
        consistent = sum(1 for s in spacings if abs(s - step) <= 3)
        if consistent < len(spacings) * 0.5:
            return boundaries

        # Группируем заголовки: каждый заголовок — отдельная операция,
        # данные идут до следующего заголовка
        for group_idx, header_row in enumerate(header_rows):
            # Определяем границы: от текущего заголовка до следующего
            if group_idx + 1 < len(header_rows):
                data_end = header_rows[group_idx + 1] - 1
            else:
                data_end = max_row

            # Извлекаем имя операции из колонки D той же строки
            op_name = ""
            op_val = ws.cell_value(header_row, 4)
            if op_val is not None and str(op_val).strip():
                op_name = str(op_val).strip()

            boundaries.append(TableBoundary(
                header_row=header_row,
                data_start=header_row + 1,
                data_end=data_end,
                operation_name=op_name,
                source_path=source_path,
                sheet_name=sheet_name,
                card_label=(
                    f"{group_idx + 1:03d}_{_safe_filename(op_name)[:30]}"
                    if op_name else f"Op{group_idx + 1:03d}"
                ),
            ))

    finally:
        reader.close()

    return boundaries


def _find_table_data_end(
    ws: Any,
    header_row: int,
    max_row: int,
    part_no_col: int,
) -> int:
    """Найти последнюю строку данных таблицы.

    Определяет границу между текущей таблицей и следующей операцией,
    используя полный список PART_NO_KEYWORDS и сканирование 25 колонок.
    """
    from burlak_parser.heuristic_analyzer import HeuristicAnalyzer, PART_NO_KEYWORDS

    empty_run = 0
    for r in range(header_row + 1, min(max_row + 1, header_row + 500)):
        # ── New-header detection: check ALL rows, not just empty part_no_col ──
        # A new operation header has 2+ non-empty cells and a PART_NO_KEYWORD
        non_empty = 0
        row_values_check: List[str] = []
        max_check_col = min((ws.max_column or 10) + 1, 25)
        for c in range(1, max_check_col):
            v = ws.cell_value(r, c)
            if v is not None:
                non_empty += 1
                rv = str(v).strip().lower()
                if len(rv) < 50:
                    row_values_check.append(rv)

        if non_empty >= 2:
            has_part_no_keyword = any(
                any(kw in rv for kw in PART_NO_KEYWORDS)
                for rv in row_values_check
            )
            if has_part_no_keyword:
                # This is a new table header — end current table at previous row
                return r - 1

        # ── Empty-run detection (fallback for tables with blank separator rows) ──
        val = ws.cell_value(r, part_no_col)
        if val is None or (isinstance(val, str) and not val.strip()):
            empty_run += 1
            if empty_run >= 3:
                return r - 3
        else:
            empty_run = 0

    return min(max_row, header_row + 499)


def _vertical_split_worker(
    source_path: str,
    output_dir: str,
    sheet_name: str,
    boundaries: List[TableBoundary],
    card_label: str,
    preloaded_zip: Optional[bytes] = None,
) -> List[str]:
    """Разделить один лист по вертикальным границам через ZIP-манипуляцию.

    Сохраняет изображения, форматирование и стили, работая напрямую
    с ZIP-структурой .xlsx файла (а не через openpyxl Workbook).

    Алгоритм для каждой операции:
      1. Скопировать исходный .xlsx (один лист, все изображения).
      2. Отфильтровать sheet XML: оставить только <row> нужного диапазона.
      3. Отфильтровать drawing XML: оставить только anchors нужного диапазона.
      4. Скорректировать row-позиции в anchors.
      5. Записать изменённый ZIP.

    Args:
        source_path: Путь к исходному .xlsx файлу (уже один лист).
        output_dir: Директория для сохранения результатов.
        sheet_name: Имя листа.
        boundaries: Список границ таблиц (операций).
        card_label: Метка для именования файлов.

    Returns:
        Список путей к созданным файлам.
    """
    os.makedirs(output_path_dir := output_dir, exist_ok=True)
    created: List[str] = []
    safe_label = _safe_filename(card_label)[:50] if card_label else ""

    # Читаем ZIP в память
    if preloaded_zip is not None:
        zip_data = preloaded_zip
    else:
        try:
            with open(source_path, 'rb') as f:
                zip_data = f.read()
        except OSError as e:
            logger.error("Cannot read source for vertical split: %s", e)
            return created

    # Читаем все ZIP-entries один раз
    try:
        with zipfile.ZipFile(io.BytesIO(zip_data), 'r') as zf:
            all_entries: Dict[str, bytes] = {}
            for name in zf.namelist():
                try:
                    all_entries[name] = zf.read(name)
                except zipfile.BadZipFile as e:
                    logger.warning("Skipping corrupt entry %s: %s", name, e)
    except zipfile.BadZipFile as e:
        logger.error("Cannot open source ZIP for vertical split: %s", e)
        return created

    all_names = set(all_entries.keys())

    # Находим имя листа в workbook.xml для определения rId
    wb_xml_bytes = all_entries.get('xl/workbook.xml')
    if wb_xml_bytes is None:
        return created

    wb_root = ET.fromstring(wb_xml_bytes)
    sheets_elem = wb_root.find(f'{{{NS_MAIN}}}sheets')
    if sheets_elem is None:
        return created

    # rId → sheet name mapping
    target_r_id: Optional[str] = None
    for sheet_el in sheets_elem.findall(f'{{{NS_MAIN}}}sheet'):
        if sheet_el.get('name') == sheet_name:
            target_r_id = (sheet_el.get(f'{{{NS_R}}}id')
                           or sheet_el.get('r:id'))
            break

    if target_r_id is None:
        return created

    # rId → target path из workbook.xml.rels
    sheet_target: Optional[str] = None
    rels_bytes = all_entries.get('xl/_rels/workbook.xml.rels')
    if rels_bytes is not None:
        rels_root = ET.fromstring(rels_bytes)
        for rel_el in rels_root:
            if (rel_el.get('Id') == target_r_id
                    and 'worksheet' in rel_el.get('Type', '')):
                sheet_target = rel_el.get('Target', '').lstrip('/')
                if not sheet_target.startswith('xl/'):
                    sheet_target = 'xl/' + sheet_target
                break

    if sheet_target is None:
        return created

    # Определяем drawing XML для этого листа
    sheet_dir = os.path.dirname(sheet_target)
    sheet_base = os.path.basename(sheet_target)
    sheet_rels_path = f"{sheet_dir}/_rels/{sheet_base}.rels"
    drawing_path: Optional[str] = None
    vml_path: Optional[str] = None

    sr_bytes = all_entries.get(sheet_rels_path)
    if sr_bytes is not None:
        sr_root = ET.fromstring(sr_bytes)
        sr_base_dir = os.path.dirname(sheet_target)
        for sr_el in sr_root:
            target = sr_el.get('Target', '')
            rtype = sr_el.get('Type', '')
            resolved = os.path.normpath(
                os.path.join(sr_base_dir, target)).replace(os.sep, '/')
            if 'drawing' in rtype.lower() and 'vml' not in rtype.lower():
                drawing_path = resolved
            elif 'vml' in rtype.lower():
                vml_path = resolved

    # Читаем drawing XML (если есть)
    drawing_xml: Optional[ET.Element] = None
    if drawing_path and drawing_path in all_entries:
        drawing_xml = ET.fromstring(all_entries[drawing_path])

    safe_label_prefix = _safe_filename(card_label)[:50] if card_label else ""

    for i, boundary in enumerate(boundaries):
        op_label = boundary.card_label or f"Op{i + 1:03d}"
        if safe_label_prefix:
            output_filename = (
                f"{safe_label_prefix}_{_safe_filename(op_label)[:40]}.xlsx")
        else:
            output_filename = f"{_safe_filename(op_label)[:50]}.xlsx"

        output_path = os.path.join(output_dir, output_filename)

        # Post-split validation: skip empty operations
        if boundary.data_end <= boundary.header_row:
            logger.warning(
                "Skipping empty operation %d in %s",
                i + 1, os.path.basename(source_path))
            continue

        try:
            with zipfile.ZipFile(output_path, 'w',
                                 zipfile.ZIP_DEFLATED) as zf_write:
                for name, data in all_entries.items():
                    if name == sheet_target:
                        data = _filter_sheet_xml(
                            data, boundary.header_row, boundary.data_end)
                    elif name == drawing_path and drawing_xml is not None:
                        data = _filter_drawing_xml(
                            drawing_xml, boundary.header_row,
                            boundary.data_end)
                    elif name == vml_path and vml_path is not None:
                        data = _filter_vml_xml(
                            data, boundary.header_row,
                            boundary.data_end)

                    zf_write.writestr(name, data)

        except zipfile.BadZipFile as e:
            logger.error("Failed to create vertical split %s: %s",
                         output_filename, e)
            continue

        # Validate the output file
        if not _validate_split_file(output_path):
            try:
                os.remove(output_path)
            except OSError:
                pass
            logger.warning("Deleted invalid split file: %s", output_filename)
            continue

        created.append(output_path)
        logger.info(
            "Вертикальный split (ZIP): операция %d '%s' [%d-%d] → %s",
            i + 1, boundary.operation_name[:30] or "",
            boundary.header_row, boundary.data_end,
            os.path.basename(output_path))

    return created


def _filter_sheet_xml(
    sheet_data: bytes,
    keep_from_row: int,
    keep_to_row: int,
) -> bytes:
    """Отфильтровать sheet XML, оставляя только строки в диапазоне.

    Корректирует позиции строк, merged cells, auto filter,
    conditional formatting и data validation.
    """
    root = ET.fromstring(sheet_data)
    ns = NS_MAIN

    # 1. Фильтруем <row> элементы
    sheet_data_elem = root.find(f'{{{ns}}}sheetData')
    if sheet_data_elem is not None:
        rows_to_remove = []
        for row_el in sheet_data_elem.findall(f'{{{ns}}}row'):
            r = int(row_el.get('r', '0'))
            if r < keep_from_row or r > keep_to_row:
                rows_to_remove.append(row_el)
            else:
                # Корректируем номер строки
                new_r = r - keep_from_row + 1
                row_el.set('r', str(new_r))
                # Корректируем cell references (r="A3390" → r="A1")
                for c_el in row_el.findall(f'{{{ns}}}c'):
                    ref = c_el.get('r', '')
                    m = _CELL_REF_RE.match(ref)
                    if m:
                        c_el.set('r', f'{m.group(1)}{new_r}')
                # Корректируем row spans
                spans = row_el.get('spans')
                if spans:
                    row_el.set('spans', spans)  # spans — col range, не меняем
        for row_el in rows_to_remove:
            sheet_data_elem.remove(row_el)

    # 1b. Обновляем <dimension> чтобы отражал реальное количество строк
    dim_elem = root.find(f'{{{ns}}}dimension')
    if dim_elem is not None:
        # Parse existing dimension to get max column
        old_ref = dim_elem.get('ref', 'A1')
        from openpyxl.utils import get_column_letter, range_boundaries
        try:
            _, _, max_col, _ = range_boundaries(old_ref)
        except (ValueError, IndexError):
            max_col = 10
        new_count = keep_to_row - keep_from_row + 1
        dim_elem.set('ref', f'A1:{get_column_letter(max_col)}{new_count}')

    # 2. Фильтруем mergeCells
    merge_cells = root.find(f'{{{ns}}}mergeCells')
    if merge_cells is not None:
        to_remove = []
        for mc in merge_cells.findall(f'{{{ns}}}mergeCell'):
            ref = mc.get('ref', '')
            if not ref:
                continue
            parts = ref.split(':')
            if len(parts) != 2:
                continue
            try:
                from openpyxl.utils import range_boundaries
                min_col, min_r, max_col, max_r = range_boundaries(ref)
            except (ValueError, IndexError):
                continue
            if max_r < keep_from_row or min_r > keep_to_row:
                to_remove.append(mc)
            else:
                new_min_r = max(min_r, keep_from_row) - keep_from_row + 1
                new_max_r = min(max_r, keep_to_row) - keep_from_row + 1
                from openpyxl.utils import get_column_letter
                new_ref = (
                    f"{get_column_letter(min_col)}{new_min_r}:"
                    f"{get_column_letter(max_col)}{new_max_r}")
                mc.set('ref', new_ref)
        for mc in to_remove:
            merge_cells.remove(mc)

    # 3. Очищаем autoFilter (может ссылаться на удалённые строки)
    auto_filter = root.find(f'{{{ns}}}autoFilter')
    if auto_filter is not None:
        root.remove(auto_filter)

    # 4. Очищаем dataValidations
    data_validations = root.find(f'{{{ns}}}dataValidations')
    if data_validations is not None:
        root.remove(data_validations)

    # 5. Фильтруем pageBreaks (rowBreaks)
    for tag in (f'{{{ns}}}rowBreaks', f'{{{ns}}}colBreaks'):
        breaks = root.find(tag)
        if breaks is None:
            continue
        if tag.endswith('}rowBreaks'):
            to_remove = []
            for br in breaks.findall(f'{{{ns}}}brk'):
                r = int(br.get('id', '0'))
                if r < keep_from_row or r > keep_to_row:
                    to_remove.append(br)
                else:
                    br.set('id', str(r - keep_from_row + 1))
            for br in to_remove:
                breaks.remove(br)

    return _serialize_xml(root, NS_MAIN)


def _filter_drawing_xml(
    drawing_root: ET.Element,
    keep_from_row: int,
    keep_to_row: int,
) -> bytes:
    """Отфильтровать drawing XML, оставляя только anchors нужного диапазона.

    Поддерживает twoCellAnchor и oneCellAnchor.
    Корректирует row-позиции anchors.
    """
    ns = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'

    for tag in (f'{{{ns}}}twoCellAnchor', f'{{{ns}}}oneCellAnchor',
                f'{{{ns}}}absoluteAnchor'):
        to_remove = []
        for anchor in drawing_root.findall(f'.//{tag}'):
            from_elem = anchor.find(f'{{{ns}}}from')
            to_elem = anchor.find(f'{{{ns}}}to')

            if from_elem is not None:
                row_elem = from_elem.find(f'{{{ns}}}row')
                if row_elem is not None and row_elem.text is not None:
                    from_row = int(row_elem.text)
                else:
                    from_row = 0
            else:
                from_row = 0

            if to_elem is not None:
                row_elem = to_elem.find(f'{{{ns}}}row')
                if row_elem is not None and row_elem.text is not None:
                    to_row = int(row_elem.text)
                else:
                    to_row = from_row
            else:
                to_row = from_row

            # Проверяем overlap с диапазоном строк
            if to_row < keep_from_row or from_row > keep_to_row:
                to_remove.append(anchor)
                continue

            # Корректируем row-позиции
            if from_elem is not None:
                row_elem = from_elem.find(f'{{{ns}}}row')
                if row_elem is not None:
                    row_elem.text = str(from_row - keep_from_row + 1)
            if to_elem is not None:
                row_elem = to_elem.find(f'{{{ns}}}row')
                if row_elem is not None:
                    row_elem.text = str(to_row - keep_from_row + 1)

        for anchor in to_remove:
            for parent in drawing_root.iter():
                if anchor in list(parent):
                    parent.remove(anchor)
                    break

    return _serialize_xml(
        drawing_root, NS_MAIN,
        extra_ns={'xdr': NS_DRAWING, 'a': NS_DRAWINGML},
    )


def _filter_vml_xml(
    vml_data: bytes,
    keep_from_row: int,
    keep_to_row: int,
) -> bytes:
    """Отфильтровать VML XML, удаляя элементы за пределами диапазона.

    VML использует style="position:absolute;left:...;top:..."
    для позиционирования. Упрощённая фильтрация: удаляем элементы
    с row-атрибутами вне диапазона.
    """
    try:
        root = ET.fromstring(vml_data)
    except ET.ParseError:
        return vml_data

    # VML namespace
    vml_ns = 'urn:schemas-microsoft-com:vml'
    office_ns = 'urn:schemas-microsoft-com:office:office'

    to_remove = []
    for elem in root.iter(f'{{{vml_ns}}}shape'):
        row_attr = elem.get('row') or elem.get(f'{{{office_ns}}}row')
        if row_attr is not None:
            try:
                row_num = int(row_attr.split()[0])
                if row_num < keep_from_row or row_num > keep_to_row:
                    to_remove.append(elem)
            except (ValueError, IndexError):
                pass

    for elem in to_remove:
        for parent in root.iter():
            if elem in list(parent):
                parent.remove(elem)
                break

    return _serialize_xml(root, NS_MAIN, extra_ns={'v': VML_NS, 'o': OFFICE_NS})
