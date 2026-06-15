"""Модуль классификации файлов операционных карт.

Определяет, является ли Excel-файл операционной картой (подлежит парсингу деталей)
или служебным документом (игнорируется при парсинге, но может быть разделён на листы).

Правила идентификации операционных карт:
  - Тип 1: "<PREFIX>-A-AS-<Number>-<Process>.xlsx"
    где <PREFIX> — любой буквенный код модели (напр. SQRT1L, SQRX90).
  - Тип 2: "<Number>-<Process>.xlsx"
    где <Number> — 3 цифры (001-999) или 1 заглавная буква + 3 цифры (A001-Z999).

Правила идентификации служебных файлов (игнорируются при парсинге деталей):
  - Имя файла не начинается с номера операции (цифры или буквы+цифры).
  - Имя файла содержит ключевые слова: 封面, 目录, 记录表, 空表.

Исключение: папки CP7/CP8 — файлы проверок/прошивок/тестирования.
Не содержат каталожных номеров, но ВСЕ РАВНО разделяются на листы.
"""

from __future__ import annotations

import logging
import os
import re
from dataclasses import dataclass
from typing import List, Optional

logger = logging.getLogger(__name__)

# Ключевые слова служебных файлов (китайский)
SERVICE_FILE_KEYWORDS = [
    "封面",        # обложка / титульный лист
    "目录",        # каталог / оглавление
    "记录表",      # таблица учёта / реестр выдачи документов
    "空表",        # пустая форма / шаблон
    "填写范本",    # образец заполнения / fill template
    "填写说明",    # инструкция по заполнению / fill instructions
    "工艺现场工时汇总清单",  # сводка трудозатрат / work hours summary
    "工时汇总",    # сводка трудозатрат (краткая форма)
]

# Папки финальных операций — файлы из них всегда разделяются на листы
FINAL_CHECK_FOLDERS = {"CP7", "CP8", "cp7", "cp8"}

# Регулярное выражение для номера операции:
#   - 3 цифры (001-999 или просто 1-999 без ведущих нулей)
#   - или 1 заглавная латинская буква + 3 цифры (A001-Z999)
OPERATION_NUMBER_RE = re.compile(
    r"^([A-Z]\d{3}|\d{1,3})"
)

# Регулярное выражение для паттерна "PREFIX-A-AS-NUMBER-PROCESS"
# Префикс: SQRT1L, SQRX90 и т.д.
PREFIX_AS_RE = re.compile(
    r"^[A-Z0-9]+-A-AS-\d+",
    re.IGNORECASE,
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
    # Родительская папка — непосредственный родитель файла
    parent_dir = os.path.basename(os.path.dirname(file_path))

    # Определяем, из какой папки файл (CP7/CP8 или вложенная)
    # Нужно проверить всю цепочку папок
    path_parts = os.path.normpath(file_path).split(os.sep)
    is_final_check = any(
        p in FINAL_CHECK_FOLDERS for p in path_parts
    )

    # Проверяем служебные ключевые слова
    is_service_file = _contains_service_keywords(file_name)
    if is_service_file:
        logger.debug("Служебный файл (ключевое слово): %s", basename)

    # Пытаемся извлечь номер операции из имени файла
    operation_number = _extract_operation_number(file_name)

    # Определяем тип файла — номер операции / префикс важнее ключевых слов
    if operation_number or PREFIX_AS_RE.match(file_name):
        # Файл с номером операции — всегда операционная карта
        is_operational = True
        should_parse = True
        should_split = True
        is_service_file = False  # номер операции перекрывает ключевые слова
    elif is_service_file:
        # Служебный файл без номера операции — не парсим детали
        is_operational = False
        should_parse = False
        # CP7/CP8 разделяем, НО НЕ шаблоны/инструкции (空表, 范本, 说明)
        is_template = any(kw in file_name for kw in ["空表", "范本", "说明"])
        should_split = is_final_check and not is_template
    elif is_final_check:
        # Файл из CP7/CP8 без номера операции — не парсим, но разделяем
        is_operational = False
        should_parse = False
        should_split = True
    else:
        # Неизвестный формат — пропускаем (не парсим, не разделяем)
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
    """Получить список файлов, из которых нужно парсить детали.

    Args:
        classifications: Результаты классификации.

    Returns:
        Список путей к файлам для парсинга деталей.
    """
    return [c.file_path for c in classifications if c.should_parse_parts]


def get_splittable_files(classifications: List[FileClassification]) -> List[FileClassification]:
    """Получить список классификаций файлов, которые нужно разделять на листы.

    Args:
        classifications: Результаты классификации.

    Returns:
        Список FileClassification для разделения.
    """
    return [c for c in classifications if c.should_split]


def _contains_service_keywords(file_name: str) -> bool:
    """Проверить, содержит ли имя файла служебные ключевые слова."""
    for kw in SERVICE_FILE_KEYWORDS:
        if kw in file_name:
            return True
    return False


def _extract_operation_number(file_name: str) -> str:
    """Извлечь номер операции из имени файла.

    Ищет в начале имени файла паттерн:
      - 3 цифры (001-999)
      - 1 заглавная буква + 3 цифры (A001-Z999)

    Args:
        file_name: Имя файла без расширения.

    Returns:
        Номер операции или пустую строку.
    """
    # Убираем возможный префикс SQRT...-A-AS- если есть
    # Пример: "SQRT1L-A-AS-038-Установка" → "038-Установка"
    # Но сначала пробуем извлечь номер из полного имени

    # Пробуем найти паттерн в начале имени
    match = OPERATION_NUMBER_RE.match(file_name)
    if match:
        return match.group(1)

    # Пробуем найти номер после "-A-AS-" префикса
    prefix_match = PREFIX_AS_RE.match(file_name)
    if prefix_match:
        remainder = file_name[prefix_match.end():]
        # Пропускаем ведущий дефис если есть
        if remainder.startswith("-"):
            remainder = remainder[1:]
        op_match = OPERATION_NUMBER_RE.match(remainder)
        if op_match:
            return op_match.group(1)

    return ""
