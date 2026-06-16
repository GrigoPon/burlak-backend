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
from typing import Dict, List, Optional, Set, Tuple

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

# Регистрируем пространства имён глобально
ET.register_namespace('', NS_MAIN)
ET.register_namespace('r', NS_R)
ET.register_namespace('ct', NS_CT)


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
            while os.path.exists(output_path):
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

            # Разрешение коллизий имён (детерминированное)
            counter = 1
            while os.path.exists(output_path):
                base, ext = os.path.splitext(output_filename)
                output_path = os.path.join(output_dir, f"{base}_{counter}{ext}")
                counter += 1

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

        with ProcessPoolExecutor(max_workers=self.max_workers) as executor:
            futures = {}
            for source_path, output_dir, sheet_names, file_label in tasks:
                future = executor.submit(
                    _split_file_worker,
                    source_path, output_dir, sheet_names, file_label,
                )
                futures[future] = source_path

            for future in as_completed(futures):
                source_path = futures[future]
                try:
                    result = future.result()
                    created, oxl_count, oxl_files, worker_manifest = result
                    all_created.extend(created)
                    all_openpyxl_count += oxl_count
                    all_openpyxl_files.extend(oxl_files)
                    for orig, generated in worker_manifest.items():
                        merged_manifest.setdefault(orig, []).extend(generated)
                except Exception as e:
                    err_msg = str(e)
                    logger.error(
                        "Ошибка параллельного разделения %s: %s",
                        os.path.basename(source_path), err_msg,
                    )
                    errors.append((source_path, err_msg))

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
                    except Exception:
                        pass
        except Exception:
            # Если что-то пошло не чисто — не критично, openpyxl сам обработает
            pass

        for name in sheets_to_remove:
            try:
                del wb[name]
            except Exception:
                pass

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

        # Читаем ZIP-архив в память
        with open(output_path, 'rb') as f:
            zip_data = f.read()

        with zipfile.ZipFile(io.BytesIO(zip_data), 'r') as zf:
            zip_entries = {name: zf.read(name) for name in zf.namelist()}

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
            zip_entries['[Content_Types].xml'] = ET.tostring(
                ct_root, encoding='UTF-8', xml_declaration=True,
            )

        # ── 8. Удалить файлы из архива ──
        for fname in list(files_to_remove):
            zip_entries.pop(fname, None)

        # ── 9. Записать обновлённые XML ──
        zip_entries['xl/workbook.xml'] = ET.tostring(
            wb_root, encoding='UTF-8', xml_declaration=True,
        )
        zip_entries['xl/_rels/workbook.xml.rels'] = ET.tostring(
            rels_root, encoding='UTF-8', xml_declaration=True,
        )

        # ── 10. Записать новый ZIP ──
        os.remove(output_path)
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for name, data in zip_entries.items():
                zout.writestr(name, data)


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
            sr_dir = os.path.dirname(removed_rels)
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
        except Exception:
            pass


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


def _split_file_worker(
    source_path: str,
    output_dir: str,
    sheet_names: List[str],
    file_label: str,
) -> Tuple[List[str], int, List[str], Dict[str, List[str]]]:
    """Рабочая функция для параллельного разделения (выполняется в отдельном процессе).

    Args:
        source_path: Путь к исходному файлу.
        output_dir: Директория для результатов.
        sheet_names: Список имён листов для выделения.
        file_label: Метка файла для именования.

    Returns:
        Кортеж (created_files, openpyxl_fallback_count, openpyxl_fallback_files, manifest).
    """
    splitter = CardSplitter(max_workers=1)  # Внутри процесса — один поток
    created = splitter.split_file(source_path, output_dir, sheet_names, file_label)
    return (
        created,
        splitter.openpyxl_fallback_count,
        sorted(splitter.openpyxl_fallback_files),
        dict(splitter.manifest),
    )
