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

# Пространства имён Excel OOXML
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_PKG_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_DRAWINGML = "http://schemas.openxmlformats.org/drawingml/2006/main"
VML_NS = "urn:schemas-microsoft-com:vml"
OFFICE_NS = "urn:schemas-microsoft-com:office:office"

# Регистрируем пространства имён глобально (только для sheet XML)
ET.register_namespace('', NS_MAIN)
ET.register_namespace('r', NS_R)
ET.register_namespace('xdr', NS_DRAWING)
ET.register_namespace('a', NS_DRAWINGML)


def _serialize_xml(root: ET.Element, default_ns_uri: str,
                   extra_ns: Optional[Dict[str, str]] = None) -> bytes:
    """Serialize an ET.Element with a specific default namespace.

    Saves and restores the global ET._namespace_map to avoid corruption.
    Used because different OOXML files require different default namespaces:
      - sheet XML:  NS_MAIN as default
      - Content_Types:  NS_CT as default
      - rels files:  NS_PKG_RELS as default
    """
    old_default = ET._namespace_map.get('')
    old_extras: Dict[str, Optional[str]] = {}
    try:
        ET.register_namespace('', default_ns_uri)
        if extra_ns:
            for p, uri in extra_ns.items():
                old_extras[p] = ET._namespace_map.get(p)
                ET.register_namespace(p, uri)
        return ET.tostring(root, xml_declaration=True, encoding='UTF-8')
    finally:
        ET.register_namespace('', old_default)
        if extra_ns:
            for p in old_extras:
                v = old_extras[p]
                if v is not None:
                    ET.register_namespace(p, v)
                else:
                    ET._namespace_map.pop(p, None)


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
        """Выделить один лист через прямую манипуляцию ZIP (без openpyxl save/load).

        Алгоритм «удаления лишнего»:
          1. Копировать исходный .xlsx как бинарный файл (shutil.copy2).
          2. Прочитать как ZIP-архив → получить все записи.
          3. Найти в xl/workbook.xml целевой лист.
          4. Удалить из workbook.xml все остальные <sheet>.
          5. Удалить файлы ненужных листов (.xml, .rels, drawings, vml).
          6. Очистить definedNames (named ranges), ссылающиеся на удалённые листы.
          7. Обновить xl/_rels/workbook.xml.rels.
          8. Обновить [Content_Types].xml.
          9. Записать изменённый ZIP.

        Args:
            source_path: Путь к исходному .xlsx файлу.
            output_path: Путь для сохранения нового .xlsx файла.
            keep_sheet_name: Имя листа, который нужно оставить.

        Raises:
            ValueError: Если целевой лист не найден в файле.
        """
        # Копируем исходный файл побайтово (shutil.copy2 сохраняет метаданные)
        shutil.copy2(source_path, output_path)

        # File size guard — warn for large files
        file_size_mb = os.path.getsize(source_path) / (1024 * 1024)
        if file_size_mb > 100:
            logger.warning(
                "Large file (%.1f MB): %s — ZIP extraction may use significant memory",
                file_size_mb, os.path.basename(source_path),
            )

        # Читаем ZIP-архив в память
        with open(output_path, 'rb') as f:
            zip_data = f.read()

        with zipfile.ZipFile(io.BytesIO(zip_data), 'r') as zf:
            zip_entries: Dict[str, bytes] = {}
            for name in zf.namelist():
                try:
                    zip_entries[name] = zf.read(name)
                except zipfile.BadZipFile as e:
                    logger.warning("Skipping corrupt ZIP entry %s: %s", name, e)

        # ── 1. Найти целевой лист ──
        wb_xml = zip_entries.get('xl/workbook.xml')
        if wb_xml is None:
            raise ValueError("Не найден xl/workbook.xml в архиве")

        wb_root = ET.fromstring(wb_xml)
        sheets_elem = wb_root.find(f'{{{NS_MAIN}}}sheets')
        if sheets_elem is None:
            raise ValueError("Не найдена секция <sheets> в workbook.xml")

        target_r_id: Optional[str] = None
        sheets_to_delete: List[Tuple[str, str, ET.Element]] = []

        for sheet_el in sheets_elem.findall(f'{{{NS_MAIN}}}sheet'):
            name = sheet_el.get('name', '')
            r_id = sheet_el.get(f'{{{NS_R}}}id') or sheet_el.get('r:id')
            if name == keep_sheet_name:
                target_r_id = r_id
            else:
                sheets_to_delete.append((name, r_id, sheet_el))

        if target_r_id is None:
            raise ValueError(f"Лист '{keep_sheet_name}' не найден в файле")

        # Если лист единственный — файл уже готов
        if not sheets_to_delete:
            return

        # ── 2. Получить rId→Target маппинг из .rels ──
        rels_xml = zip_entries.get('xl/_rels/workbook.xml.rels')
        if rels_xml is None:
            raise ValueError("Не найден xl/_rels/workbook.xml.rels")

        rels_root = ET.fromstring(rels_xml)
        WORKSHEET_TYPE = f"{NS_R}/worksheet"

        r_id_to_target: Dict[str, str] = {}
        for rel_el in rels_root:
            rid = rel_el.get('Id', '')
            target = rel_el.get('Target', '')
            r_id_to_target[rid] = target

        # Собираем имена удаляемых листов для фильтрации definedNames
        deleted_sheet_names: Set[str] = {name for name, _, _ in sheets_to_delete}

        # ── 3. Удалить ненужные <sheet> из workbook.xml ──
        for _, _, sheet_el in sheets_to_delete:
            sheets_elem.remove(sheet_el)

        # ── 4. Удалить relationship'ы для ненужных листов ──
        keep_r_ids: Set[str] = {target_r_id}
        for rel_el in list(rels_root):
            rid = rel_el.get('Id', '')
            r_type = rel_el.get('Type', '')
            if r_type == WORKSHEET_TYPE and rid not in keep_r_ids:
                rels_root.remove(rel_el)

        # ── 5. Собрать список файлов для удаления ──
        files_to_remove: Set[str] = set()

        for name, r_id, _ in sheets_to_delete:
            if r_id and r_id in r_id_to_target:
                target = r_id_to_target[r_id]
                # Нормализуем путь: убираем ведущий / и добавляем xl/ при необходимости
                # openpyxl генерирует абсолютные пути (/xl/worksheets/sheet2.xml),
                # другие генераторы — относительные (worksheets/sheet2.xml)
                removed_sheet = target.lstrip('/')
                if not removed_sheet.startswith('xl/'):
                    removed_sheet = 'xl/' + removed_sheet
                files_to_remove.add(removed_sheet)
                _collect_related_files(zip_entries, removed_sheet, files_to_remove)

        # Удаляем calcChain.xml (Excel перегенерирует)
        zip_entries.pop('xl/calcChain.xml', None)
        files_to_remove.add('xl/calcChain.xml')

        # ── 6. Очистить definedNames (named ranges) ──
        _clean_named_ranges(wb_root, deleted_sheet_names, keep_sheet_name)

        # ── 6.5 Очистить View-элементы (устраняет ошибку "Removed Records: View") ──
        book_views = wb_root.find(f'{{{NS_MAIN}}}bookViews')
        if book_views is not None:
            for wv in book_views.findall(f'{{{NS_MAIN}}}workbookView'):
                wv.attrib.pop('activeTab', None)
                wv.attrib.pop('firstSheet', None)

        custom_views = wb_root.find(f'{{{NS_MAIN}}}customWorkbookViews')
        if custom_views is not None:
            wb_root.remove(custom_views)

        # ── 7. Обновить [Content_Types].xml ──
        ct_xml = zip_entries.get('[Content_Types].xml')
        if ct_xml is not None:
            ct_root = ET.fromstring(ct_xml)
            for override_el in list(ct_root.findall(f'{{{NS_CT}}}Override')):
                part_name = override_el.get('PartName', '')
                if part_name.startswith('/'):
                    part_name = part_name[1:]
                if part_name in files_to_remove:
                    ct_root.remove(override_el)
            zip_entries['[Content_Types].xml'] = _serialize_xml(
                ct_root, NS_CT,
            )

        # ── 8. Удалить файлы из архива ──
        for fname in list(files_to_remove):
            zip_entries.pop(fname, None)

        # ── 9. Записать обновлённые XML ──
        zip_entries['xl/workbook.xml'] = _serialize_xml(
            wb_root, NS_MAIN,
        )
        zip_entries['xl/_rels/workbook.xml.rels'] = _serialize_xml(
            rels_root, NS_PKG_RELS,
        )

        # ── 10. Записать новый ZIP ──
        os.remove(output_path)
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for name, data in zip_entries.items():
                zout.writestr(name, data)


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

    return ET.tostring(root, xml_declaration=True, encoding='UTF-8')


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
