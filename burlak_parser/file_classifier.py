"""Модуль классификации файлов операционных карт.

Определяет, является ли Excel-файл операционной картой (подлежит парсингу деталей)
или служебным документом (игнорируется при парсинге, но может быть разделён на листы).

УНИВЕРСАЛЬНАЯ классификация:
  - Не зависит от конкретных префиксов (SQRT, SQR, G01, T1L, SWM, и т.д.)
  - Использует эвристические паттерны для определения операционных карт
  - Поддерживает любые буквенно-цифровые комбинации в именах файлов

Правила идентификации операционных карт:
  - Имя файла содержит номер операции (цифры, буквы+цифры) в начале
  - Или соответствует паттерну "Префикс-A-AS-Номер"
  - Или содержит известный идентификатор процесса/операции

Правила идентификации служебных файлов:
  - Имя файла содержит ключевые слова: 封面, 目录, 记录表, 空表
  - Файл не соответствует ни одному из паттернов операционной карты

Исключение: папки CP7/CP8 — файлы проверок/прошивок/тестирования.
Не содержат каталожных номеров, но ВСЕ РАВНО разделяются на листы.
"""

from __future__ import annotations

import logging
import os
import re
from dataclasses import dataclass
from typing import List, Optional

from burlak_parser.heuristic_analyzer import (
    extract_card_number_from_filepath,
)

logger = logging.getLogger(__name__)

# Ключевые слова служебных файлов (китайский / английский / русский)
SERVICE_FILE_KEYWORDS = [
    "封面",        # обложка / титульный лист
    "目录",        # каталог / оглавление
    "记录表",      # таблица учёта / реестр выдачи документов
    "空表",        # пустая форма / шаблон
    "填写范本",    # образец заполнения / fill template
    "填写说明",    # инструкция по заполнению / fill instructions
    "工艺现场工时汇总清单",  # сводка трудозатрат / work hours summary
    "工时汇总",    # сводка трудозатрат (краткая форма)
    "обложка",     # обложка
    "содержание",  # содержание
    "cover",       # cover page
    "toc",         # table of contents
    "template",    # template
]

# Папки финальных операций — файлы из них всегда разделяются на листы
FINAL_CHECK_FOLDERS = {"CP7", "CP8", "cp7", "cp8"}

# Универсальное регулярное выражение для номера операции:
#   - 2+ цифры в начале имени (возможно с буквенным префиксом)
#   - Например: 038, A001, 1234, TP005
OPERATION_NUMBER_RE = re.compile(
    r"^(?:[A-Za-z]{1,3})?(\d{2,})",
)

# Паттерн для префикса модели + "-A-AS-" + номер (универсальный, не только SQRT)
# Например: SQRT1L-A-AS-04001, G01-A-AS-05001, SWM-A-AS-001
PREFIX_AS_RE = re.compile(
    r"^[A-Za-z0-9]+-[A-Za-z0-9]*-AS-\d+",
    re.IGNORECASE,
)

# Паттерн для "буквы + цифры" в начале имени (номер карты)
CARD_START_RE = re.compile(
    r"^([A-Za-z]{1,4}\d{2,})",  # LG01, T1L, A001, TP01
)

# Паттерн для "цифры в начале" (номер операции)
DIGIT_START_RE = re.compile(
    r"^(\d{2,})",
)


@dataclass
class FileClassification:
    """Результат классификации одного файла."""

    file_path: str
    file_name: str  # basename без расширения
    parent_folder: str  # имя родительской папки
    is_operational_card: bool  # операционная карта (содержит таблицу деталей)
    is_service_file: bool  # служебный файл (без таблицы деталей)
    is_final_check: bool  # файл из папки CP7/CP8
    should_split: bool  # нужно ли разделять на листы
    should_parse_parts: bool  # нужно ли парсить детали
    operation_number: str = ""  # номер операции (если определён)


def classify_file(file_path: str) -> FileClassification:
    """Классифицировать файл Excel.

    Args:
        file_path: Полный путь к .xlsx/.xls файлу.

    Returns:
        FileClassification с результатом классификации.
    """
    basename = os.path.basename(file_path)
    file_name = os.path.splitext(basename)[0]
    parent_dir = os.path.basename(os.path.dirname(file_path))

    # Определяем, из какой папки файл (CP7/CP8 или вложенная)
    path_parts = os.path.normpath(file_path).split(os.sep)
    is_final_check = any(p in FINAL_CHECK_FOLDERS for p in path_parts)

    # Проверяем служебные ключевые слова
    is_service_file = _contains_service_keywords(file_name)

    # Пытаемся извлечь номер операции из имени файла
    operation_number = _extract_operation_number(file_name)

    # Определяем, является ли файл операционной картой (проверяем по номеру операции)
    has_card_pattern = bool(operation_number)

    # Определяем тип файла
    if has_card_pattern or operation_number:
        # Файл с номером операции или паттерном карты — всегда операционная карта
        is_operational = True
        should_parse = True
        should_split = True
        is_service_file = False  # номер операции перекрывает ключевые слова
    elif is_service_file:
        # Служебный файл без номера операции — не парсим детали
        is_operational = False
        should_parse = False
        is_template = any(kw in file_name for kw in ["空表", "范本", "说明"])
        should_split = is_final_check and not is_template
    elif is_final_check:
        # Файл из CP7/CP8 без номера операции — не парсим, но разделяем
        is_operational = False
        should_parse = False
        should_split = True
    else:
        # Неизвестный формат — пробуем извлечь номер карты эвристически
        card_no = extract_card_number_from_filepath(file_path)
        if card_no and card_no != file_name:
            # Имя содержит номер карты — считаем операционной
            logger.debug("Файл определён как операционная карта (эвристика): %s", basename)
            is_operational = True
            should_parse = True
            should_split = True
            operation_number = card_no
        else:
            # Неизвестный формат — пропускаем
            logger.warning("Неизвестный формат файла, пропускается: %s", basename)
            is_operational = False
            should_parse = False
            should_split = False

    classification = FileClassification(
        file_path=file_path,
        file_name=file_name,
        parent_folder=parent_dir,
        is_operational_card=is_operational,
        is_service_file=is_service_file,
        is_final_check=is_final_check,
        should_split=should_split,
        should_parse_parts=should_parse,
        operation_number=operation_number,
    )

    return classification


def filter_operational_cards(file_paths: List[str]) -> List[FileClassification]:
    """Отфильтровать список файлов, классифицируя каждый.

    Args:
        file_paths: Список путей к Excel-файлам.

    Returns:
        Список FileClassification для всех файлов.
    """
    return [classify_file(fp) for fp in file_paths]


def get_parseable_files(classifications: List[FileClassification]) -> List[str]:
    """Получить список файлов, из которых нужно парсить детали."""
    return [c.file_path for c in classifications if c.should_parse_parts]


def get_splittable_files(classifications: List[FileClassification]) -> List[FileClassification]:
    """Получить список классификаций файлов, которые нужно разделять на листы."""
    return [c for c in classifications if c.should_split]


def _contains_service_keywords(file_name: str) -> bool:
    """Проверить, содержит ли имя файла служебные ключевые слова."""
    name_lower = file_name.lower()
    for kw in SERVICE_FILE_KEYWORDS:
        if kw in name_lower:
            return True
        # Английские ключевые слова
        if kw in name_lower.replace("_", " ").replace("-", " "):
            return True
    return False


def _extract_operation_number(file_name: str) -> str:
    """Извлечь номер операции из имени файла.

    Универсальный алгоритм:
      1. Проверяет префикс "Модель-A-AS-Номер"
      2. Проверяет букву + 2+ цифры в начале
      3. Проверяет 2+ цифры в начале

    Args:
        file_name: Имя файла без расширения.

    Returns:
        Номер операции или пустую строку.
    """
    # Префикс-AS-паттерн (универсальный)
    match = PREFIX_AS_RE.match(file_name)
    if match:
        return match.group(0)

    # Буква + 2+ цифры в начале (A001, T1L, и т.д.)
    match = CARD_START_RE.match(file_name)
    if match:
        return match.group(1)

    # 2+ цифры в начале
    match = DIGIT_START_RE.match(file_name)
    if match:
        return match.group(1)

    return ""



