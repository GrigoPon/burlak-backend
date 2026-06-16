"""Модуль безопасного нечеткого сравнения (fuzzy matching) каталожных номеров.

Ключевое правило:
  - Совпадением считается ТОЛЬКО различие в дефисах, пробелах и спецсимволах.
  - Различие даже в ОДНУ цифру (напр. "ABCD123" vs "ABCD124") НЕ считается совпадением.
    Это абсолютно разные детали на производстве.

Алгоритм:
  1. Нормализовать оба номера: удалить все пробелы, дефисы, спецсимволы.
  2. Сравнить нормализованные строки.
  3. Если они идентичны — это fuzzy match.
  4. Если отличаются — НЕ match, даже на один символ.

Дополнительно:
  - Проверка целостности парт-номеров: отсеивание мусора, не похожего на каталожный номер.
"""

from __future__ import annotations

import logging
import re
from typing import Dict, List, Optional, Set

from burlak_parser.normalizer import (
    normalize_part_number,
    is_valid_part_number as _is_valid_part_number_strict,
)

logger = logging.getLogger(__name__)


def is_fuzzy_match(part_a: str, part_b: str) -> bool:
    """Проверить, являются ли два парт-номера нечетким совпадением.

    Считает совпадением только номера, идентичные после очистки от спецсимволов.
    Различие в цифрах/буквах НЕ допускается.

    Args:
        part_a: Первый парт-номер.
        part_b: Второй парт-номер.

    Returns:
        True если номера совпадают после нормализации.

    Example:
        is_fuzzy_match("ABCD-123", "ABCD123") -> True
        is_fuzzy_match("ABCD 123", "ABCD-123") -> True
        is_fuzzy_match("ABCD123", "ABCD124") -> False  # различие в цифре!
        is_fuzzy_match("ABCD-123", "ABCD-123") -> True  # точное совпадение
    """
    return normalize_part_number(part_a) == normalize_part_number(part_b)


def is_valid_part_number(part_no: str) -> bool:
    """Проверить, похожа ли строка на каталожный номер детали.

    Использует МЯГКУЮ проверку (lenient) — допускает буквы-only или цифры-only
    строки достаточной длины (для fuzzy matching).

    Для строгой проверки используйте burlak_parser.normalizer.is_valid_part_number().

    Args:
        part_no: Строка для проверки.

    Returns:
        True если строка похожа на каталожный номер.
    """
    return _is_valid_part_number_strict(part_no, strict=False)


class FuzzyMatcher:
    """Сервис нечеткого сопоставления парт-номеров.

    Строит индекс нормализованных номеров для быстрого поиска.
    """

    def __init__(self, bom_part_numbers: Set[str]):
        """Инициализировать матчер.

        Args:
            bom_part_numbers: Множество парт-номеров из BOM.
        """
        # Индекс: normalized -> список оригинальных номеров
        self._normalized_index: Dict[str, List[str]] = {}

        for pn in bom_part_numbers:
            norm = normalize_part_number(pn)
            if norm not in self._normalized_index:
                self._normalized_index[norm] = []
            self._normalized_index[norm].append(pn)

    def find_fuzzy_match(self, cards_part_no: str) -> Optional[str]:
        """Найти нечеткое совпадение для номера из карт в BOM.

        Args:
            cards_part_no: Парт-номер из операционной карты.

        Returns:
            Оригинальный парт-номер из BOM, или None если совпадений нет.
        """
        norm = normalize_part_number(cards_part_no)
        matches = self._normalized_index.get(norm, [])
        if matches:
            # Возвращаем первый (обычно он один)
            return matches[0]
        return None

    def get_normalized(self, part_no: str) -> str:
        """Получить нормализованную форму парт-номера."""
        return normalize_part_number(part_no)
