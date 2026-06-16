"""Модуль классификации файлов операционных карт.

Определяет, является ли Excel-файл операционной картой (подлежит парсингу деталей)
или служебным документом (игнорируется при парсинге и разделении).

УНИВЕРСАЛЬНАЯ классификация:
  - Не зависит от конкретных префиксов (SQRT, SQR, G01, T1L, SWM, и т.д.)
  - Не зависит от структуры папок (нет хардкода CP7/CP8)
  - Использует эвристические паттерны для определения операционных карт
  - Поддерживает любые буквенно-цифровые комбинации в именах файлов

Правила идентификации операционных карт:
  - Имя файла содержит номер операции (цифры, буквы+цифры) в начале
  - Или соответствует паттерну "Префикс-A-AS-Номер"
  - Или содержит известный идентификатор процесса/операции

Правила идентификации служебных файлов:
  - Имя файла содержит ключевые слова: 封面, 目录, 记录表, 空表
  - Файл не соответствует ни одному из паттернов операционной карты

Правила идентификации остальных файлов:
  - Если файл не является операционной картой и не служебный — пропускается
"""

from __future__ import annotations

import logging
import os
import re
from dataclasses import dataclass
from typing import List

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
    "对比",        # сравнение / comparison (служебный файл сравнения)
    "обложка",     # обложка
    "содержание",  # содержание
    "cover",       # cover page
    "toc",         # table of contents
    "template",    # template
]

# Ключевые слова операционных карт (файлы с этими словами — всегда ОК)
OPERATIONAL_CARD_KEYWORDS = [
    "作业指导书",   # BAIC: рабочая инструкция / work instruction
    "作业要领书",   # аналогичное / similar
    "操作指导",     # инструкция по операции
    "工艺卡",       # технологическая карта
    "工序卡",       # карта工序
]



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

# Паттерн для поиска номера карты в любой части имени (не только в начале)
# Например: "5. G01Pш╜жщЧич║┐х╖ешЙ║хНб" -> "G01P"
CARD_NUMBER_ANYWHERE_RE = re.compile(
    r"[A-Za-z]{1,4}\d{2,}[A-Za-z0-9]*",
)


@dataclass
class FileClassification:
    """Результат классификации одного файла."""

    file_path: str
    file_name: str  # basename без расширения
    parent_folder: str  # имя родительской папки
    is_operational_card: bool  # операционная карта (содержит таблицу деталей)
    is_service_file: bool  # служебный файл (без таблицы деталей)
    should_split: bool  # нужно ли разделять на листы
    should_parse_parts: bool  # нужно ли парсить детали
    operation_number: str = ""  # номер операции (если определён)


def classify_file(file_path: str) -> FileClassification:
    """Классифицировать файл Excel.

    Порядок приоритетов (КЛЮЧЕВОЕ: служебные ключевые слова ПЕРВЫЕ):
      1. Служебные ключевые слова (封面, 目录, 空表 и т.д.) → служебный
         Даже если имя содержит "作业指导书" или номер операции.
         Пример: "CP7作业指导书封面及目录.xlsx" → служебный (содержит "封面"+"目录")
      2. Ключевые слова операционных карт (作业指导书 и т.д.) → карта
      3. Номер операции / паттерн карты → карта
      4. Эвристика (номер карты в имени) → карта
      5. Неизвестный формат → пропуск

    Args:
        file_path: Полный путь к .xlsx/.xls файлу.

    Returns:
        FileClassification с результатом классификации.
    """
    basename = os.path.basename(file_path)
    file_name = os.path.splitext(basename)[0]
    parent_dir = os.path.basename(os.path.dirname(file_path))

    # Проверяем служебные ключевые слова (ВЫСШИЙ ПРИОРИТЕТ)
    is_service_file = _contains_service_keywords(file_name)

    # Проверяем ключевые слова операционных карт (作业指导书 и т.д.)
    is_op_card_keyword = _contains_operational_card_keyword(file_name)

    # Пытаемся извлечь номер операции из имени файла
    operation_number = _extract_operation_number(file_name)
    has_card_pattern = bool(operation_number)

    # Определяем тип файла
    if is_service_file:
        # Служебный файл — ВСЕГДА служебный, даже если содержит
        # "作业指导书" или номер операции в имени.
        # Пример: "CP7作业指导书封面及目录.xlsx" → служебный
        is_operational = False
        should_parse = False
        should_split = False
    elif is_op_card_keyword:
        # Файл с ключевым словом операционной карты — операционная карта
        is_operational = True
        should_parse = True
        should_split = True
    elif operation_number:
        # Файл с номером операции — операционная карта
        is_operational = True
        should_parse = True
        should_split = True
    else:
        # Дополнительная эвристика: ищем номер карты эвристически
        card_no = extract_card_number_from_filepath(file_path)
        if card_no and card_no != file_name:
            logger.debug("Файл определён как операционная карта (эвристика): %s", basename)
            is_operational = True
            should_parse = True
            should_split = True
            operation_number = card_no
        else:
            # Альтернативная эвристика: паттерн букв+цифр в любой части имени
            alt_card_no = _find_card_number_in_name(file_name)
            if alt_card_no:
                logger.debug(
                    "Файл определён как операционная карта (альт.эвристика): %s",
                    basename,
                )
                is_operational = True
                should_parse = True
                should_split = True
                operation_number = alt_card_no
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


def _contains_operational_card_keyword(file_name: str) -> bool:
    """Проверить, содержит ли имя файла ключевое слово операционной карты."""
    name_lower = file_name.lower()
    for kw in OPERATIONAL_CARD_KEYWORDS:
        if kw in name_lower:
            return True
    return False


def _find_card_number_in_name(file_name: str) -> str:
    """Найти номер карты в любой части имени файла.

    Для файлов с изменённой кодировкой, где стандартные паттерны
    не срабатывают из-за небуквенных/нецифровых символов в начале.

    Примеры:
      "5. G01Pш╜жщЧич║┐х╖ешЙ║хНб" -> "G01P"
      "5. T1L总装卡" -> "T1L"

    Args:
        file_name: Имя файла без расширения.

    Returns:
        Найденный номер карты или пустую строку.
    """
    # Ищем паттерн буква+цифры где угодно в имени
    match = CARD_NUMBER_ANYWHERE_RE.search(file_name)
    if match:
        return match.group(0)
    return ""


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



