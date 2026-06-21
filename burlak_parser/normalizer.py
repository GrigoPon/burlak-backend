"""Модуль нормализации данных BOM и операционных карт.

Универсальная очистка и преобразование значений ячеек:
  - Количества: "–" → 0, "– –" → 0, пустые → 0, числа как есть
  - Парт-номера: удаление спецсимволов, приведение к верхнему регистру
  - Названия: удаление лишних пробелов, переносов строк

Маркеры S/- обрабатываются на уровне BOM-парсера:
  - BOM: "S" означает "деталь присутствует, количество из.qty-колонки"
  - BOM: "–"/пусто означает "детали нет, количество = 0"
  - Карты: маркеры S/- не встречаются

Используется как центральный модуль для нормализации данных,
чтобы не дублировать логику в BOM и Card парсерах.
"""

from __future__ import annotations

import logging
import re
from typing import Any, Optional

logger = logging.getLogger(__name__)


# Паттерн для определения числа (целое или дробное, с запятой или точкой)
NUMERIC_RE = re.compile(
    r"^\s*[-+]?\d+(?:[\.,]\d+)?\s*$"
)

# Паттерн для "S"-подобных маркеров (BAIC: "S" = есть деталь)
S_MARKER_RE = re.compile(r"^\s*[sS]\s*$")

# Паттерн для тире/дефисов (BAIC: "–" = нет детали)
DASH_MARKER_RE = re.compile(r"^\s*[\-\–\—\‒\―]{1,3}\s*$")

# Паттерн для "– –" (двойное тире)
DOUBLE_DASH_RE = re.compile(r"^\s*[\-\–\—\‒\―]\s*[\-\–\—\‒\―]\s*$")

# Символы для очистки парт-номеров (включая тире разных типов)
CLEAN_PN_CHARS = re.compile(r"[\s\-–—‒\―\.\/\_\,\;\:\'\"\(\)\[\]\{\}\|\\]+")

# Китайские иероглифы
_CJK_RE = re.compile(r"[\u4e00-\u9fff\u3400-\u4dbf]")

# Кириллица
_CYRILLIC_RE = re.compile(r"[\u0400-\u04ff]")

# Паттерн для lenient проверки (fuzzy matcher): буквы+цифры >=3, или цифры >=3
PART_NUMBER_LENIENT_RE = re.compile(
    r"^(?=.*[A-Za-z])[A-Za-z0-9\-\.\/\_\s]{3,}$"
    r"|^\d{3,}$"
)

# Артефакты кодировки Excel XML
XML_ARTIFACTS_RE = re.compile(r"(_x[0-9a-fA-F]{4}_|\r\n|[\r\n])", re.IGNORECASE)


class QuantityNormalizer:
    """Нормализатор значений количества.

    Поддерживает:
      - Числа (int, float, str)
      - Тире/дефисы: "–" → 0 (нет детали)
      - Пустые/None значения → 0
      - Строки-числа: "2.5" → 2.5, "10" → 10.0

    Маркеры S НЕ обрабатываются здесь — их обработка на уровне BOM-парсера
    (S означает "деталь присутствует, количество из.qty-колонки").
    """

    @staticmethod
    def normalize(value: Any, default: float = 0.0) -> float:
        """Нормализовать значение количества в float.

        Args:
            value: Исходное значение ячейки.
            default: Значение по умолчанию при невозможности преобразования.

        Returns:
            Числовое значение количества.

        Examples:
            >>> QuantityNormalizer.normalize(42)
            42.0
            >>> QuantityNormalizer.normalize("S")
            0.0
            >>> QuantityNormalizer.normalize("–")
            0.0
            >>> QuantityNormalizer.normalize(None)
            0.0
            >>> QuantityNormalizer.normalize("2.5")
            2.5
        """
        if value is None:
            return default

        # Reject non-numeric types that could be misinterpreted
        # bool is subclass of int, so check before int
        if isinstance(value, bool):
            logger.debug("bool value detected: %s → 0.0", value)
            return default
        # datetime/timedelta from Excel date columns
        from datetime import datetime, date, timedelta
        if isinstance(value, (datetime, date, timedelta)):
            logger.debug("datetime value detected: %s → 0.0", value)
            return default

        # Числа (int/float) — возвращаем как есть
        if isinstance(value, (int, float)):
            return float(value)

        # Строки
        s = str(value).strip()
        # Удаляем артефакты кодировки Excel XML
        match = XML_ARTIFACTS_RE.search(s)
        if match:
            s = s[:match.start()]
        s = s.strip()
        if not s:
            return default

        # "–" / "—" / "– –" = нет детали → 0
        if DASH_MARKER_RE.match(s) or DOUBLE_DASH_RE.match(s):
            logger.debug("dash-marker detected: '%s' → 0.0", value)
            return 0.0

        # Строки-числа: "2.5", "10", "-3"
        if NUMERIC_RE.match(s):
            try:
                # Заменяем запятую на точку (европейский формат)
                normalized = s.replace(",", ".")
                return float(normalized)
            except (ValueError, TypeError):
                return default

        # Нераспознанный формат (включая "S") — логируем и возвращаем default
        logger.debug(
            "Unrecognized quantity format: '%s' (type=%s) → default=%.1f",
            value, type(value).__name__, default,
        )
        return default


class PartNumberNormalizer:
    """Нормализатор парт-номеров.

    Удаляет спецсимволы, приводит к верхнему регистру.
    Гарантирует консистентность парт-номеров между BOM и картами.
    """

    @staticmethod
    def normalize(part_no: str) -> str:
        """Нормализовать парт-номер для сравнения.

        Args:
            part_no: Исходный парт-номер.

        Returns:
            Нормализованный парт-номер (только буквы и цифры, upper case).

        Examples:
            >>> PartNumberNormalizer.normalize("5306200-ED001-AC00000")
            '5306200ED001AC00000'
            >>> PartNumberNormalizer.normalize("ab-123-cd")
            'AB123CD'
            >>> PartNumberNormalizer.normalize(" A.1/B_2 ")
            'A1B2'
        """
        if not part_no or not isinstance(part_no, str):
            return ""
        s = part_no.strip()
        # Удаляем артефакты кодировки Excel XML (carriage return / newline)
        # Берём ТОЛЬКО ПЕРВУЮ часть до артефакта (двойные значения: китайский + английский)
        match = XML_ARTIFACTS_RE.search(s)
        if match:
            s = s[:match.start()]
        cleaned = CLEAN_PN_CHARS.sub("", s)
        return cleaned.upper()

    @staticmethod
    def is_valid(part_no: str, strict: bool = True) -> bool:
        """Проверить, похожа ли строка на парт-номер.

        Args:
            part_no: Строка для проверки.
            strict: True = строгая проверка (letters+digits, или 4+ digits).
                    False = мягкая проверка (fuzzy matcher: letters+digits >=3, или digits >=3).

        Returns:
            True если строка похожа на парт-номер.
        """
        if not part_no or not isinstance(part_no, str):
            return False
        cleaned = part_no.strip()
        if len(cleaned) < 3:
            return False
        # Мусор
        garbage = {"n/a", "na", "none", "无", "null", "-", "--", "---", "/"}
        if cleaned.lower() in garbage:
            return False
        # Строки с китайскими иероглифами или кириллицей — не парт-номера
        if _CJK_RE.search(cleaned) or _CYRILLIC_RE.search(cleaned):
            return False
        if strict:
            # Строгая проверка: буквы+цифры, или минимум 4 цифры
            has_alpha = bool(re.search(r"[A-Za-z]", cleaned))
            has_digit = bool(re.search(r"\d", cleaned))
            if has_alpha and has_digit:
                return True
            if has_digit and len(cleaned) >= 4:
                return True
            return False
        else:
            # Мягкая проверка (fuzzy matcher): буквы+цифры >=3, или цифры >=3
            return bool(PART_NUMBER_LENIENT_RE.match(cleaned))


class NameNormalizer:
    """Нормализатор названий деталей.

    Удаляет лишние пробелы, переносы строк, приводит к единообразию.
    """

    @staticmethod
    def normalize(name: str) -> str:
        """Нормализовать название детали.

        Args:
            name: Исходное название.

        Returns:
            Нормализованное название.

        Examples:
            >>> NameNormalizer.normalize("  Деталь  assembly  ")
            'Деталь assembly'
            >>> NameNormalizer.normalize("Линия\\nЖгут")
            'Линия Жгут'
        """
        if not name or not isinstance(name, str):
            return ""
        # Удаляем переносы строк и возвраты каретки
        text = name.replace("\n", " ").replace("\r", "").replace("\t", " ")
        # Сжимаем множественные пробелы
        text = " ".join(text.split())
        return text.strip()


def normalize_quantity(value: Any, default: float = 0.0) -> float:
    """Удобная функция-обёртка для нормализации количества.

    Используется в BOM и Card парсерах для единообразной обработки.
    """
    return QuantityNormalizer.normalize(value, default=default)


def normalize_part_number(part_no: str) -> str:
    """Удобная функция-обёртка для нормализации парт-номера."""
    return PartNumberNormalizer.normalize(part_no)


def clean_part_number(part_no: str) -> str:
    """Очистить парт-номер от спецсимволов, привести к верхнему регистру.

    Алиас для normalize_part_number() — единая точка правды для всех модулей.
    """
    return PartNumberNormalizer.normalize(part_no)


def is_valid_part_number(part_no: str, strict: bool = True) -> bool:
    """Проверить, похожа ли строка на парт-номер.

    Алиас для PartNumberNormalizer.is_valid() — единая точка правды.
    strict=True  — строгая проверка (для BOM и Card парсеров)
    strict=False — мягкая проверка (для fuzzy matcher)
    """
    return PartNumberNormalizer.is_valid(part_no, strict=strict)
