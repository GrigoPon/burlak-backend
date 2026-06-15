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
from typing import Dict, List, Optional, Set, Tuple

logger = logging.getLogger(__name__)

# Паттерн для каталожного номера детали:
# Должен содержать хотя бы одну букву и хотя бы одну цифру,
# или быть чисто цифровым достаточной длины (>= 5 символов).
# Исключаем очевидный мусор: одиночные буквы, спецсимволы, слишком короткие строки.
PART_NUMBER_PATTERN = re.compile(
    r"^(?=.*[A-Za-z])[A-Za-z0-9\-\.\/\_\s]{3,}$"  # минимум 3 символа, хотя бы 1 буква
    r"|^\d{3,}$"  # или чисто цифровой, минимум 3 цифры
)

# Символы, которые удаляются при нормализации для fuzzy matching
NORMALIZE_STRIP_CHARS = re.compile(r"[\s\-\—\.\/\_\,\;\:\'\"\(\)\[\]\{\}\|\\]+")


def normalize_part_number(part_no: str) -> str:
    """Нормализовать парт-номер для сравнения.

    Удаляет ВСЕ пробелы, дефисы, точки, слеши и прочие спецсимволы.
    Приводит к верхнему регистру.

    Args:
        part_no: Исходный парт-номер.

    Returns:
        Нормализованный парт-номер (только буквы и цифры, upper case).

    Example:
        "ABCD-123" -> "ABCD123"
        "ABCD 123" -> "ABCD123"
        "ABCD.1.2.3" -> "ABCD123"
    """
    cleaned = NORMALIZE_STRIP_CHARS.sub("", part_no.strip())
    return cleaned.upper()


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

    Отсеивает:
      - Слишком короткие строки (< 3 символов)
      - Чисто буквенные строки без цифр (кроме очень длинных)
      - Строки, состоящие только из спецсимволов
      - Явный мусор: "N/A", "-", "无", "None" и т.д.

    Args:
        part_no: Строка для проверки.

    Returns:
        True если строка похожа на каталожный номер.
    """
    if not part_no or len(part_no.strip()) < 3:
        return False

    cleaned = part_no.strip()

    # Явный мусор
    garbage = {"n/a", "na", "none", "无", "null", "-", "--", "---", "/", ".", ".."}
    if cleaned.lower() in garbage:
        return False

    # Проверка паттерном
    return bool(PART_NUMBER_PATTERN.match(cleaned))


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
        # Обратный индекс: оригинальный -> normalized
        self._reverse_index: Dict[str, str] = {}

        for pn in bom_part_numbers:
            norm = normalize_part_number(pn)
            if norm not in self._normalized_index:
                self._normalized_index[norm] = []
            self._normalized_index[norm].append(pn)
            self._reverse_index[pn] = norm

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
