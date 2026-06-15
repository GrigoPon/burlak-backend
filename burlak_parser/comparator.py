"""Модуль сверки (матчинга) BOM и операционных карт.

Система автоматически сопоставляет извлечённый список деталей
для выбранной комплектации (из BOM) с агрегированным списком всех деталей,
найденных в операционных картах, по Каталожному номеру (Парт-номеру).

Виды расхождений:
  - Только в BOM: деталь есть в BOM, но не используется ни в одной карте.
  - Только в Картах: деталь есть в картах, но отсутствует в BOM для данной комплектации.
  - Конфликт количества: количество в картах не совпадает с количеством в BOM.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from typing import Dict, List, Optional, Tuple

from burlak_parser.bom_parser import BOMData, PartInfo
from burlak_parser.card_parser import CardsData

logger = logging.getLogger(__name__)


# ─── Нечёткий поиск парт-номеров ─────────────────────────────────────────

# Минимальная длина парт-номера для нечёткого поиска (короткие — ложные срабатывания)
_FUZZY_MIN_LENGTH = 4

# Порог сходства: выше → точнее, но меньше находок; ниже → больше находок, но риск ошибок
_FUZZY_THRESHOLD_SHORT = 0.90   # для номеров короче 10 символов (строже!)
_FUZZY_THRESHOLD_MEDIUM = 0.88  # для номеров 10-16 символов
_FUZZY_THRESHOLD_LONG = 0.85    # для номеров длиннее 16 символов

# Максимальная разница в длине нормализованных номеров (25%)
_FUZZY_MAX_LEN_DIFF_RATIO = 0.25


def _normalize_for_fuzzy(pn: str) -> str:
    """Нормализовать парт-номер для сравнения.

    Удаляет ВСЕ разделители (пробелы, дефисы, подчёркивания, точки),
    приводит к верхнему регистру — остаётся «алфавитно-цифровое ядро».

    Примеры:
      'F26-1101051'  → 'F261101051'
      'F26 - 1101051' → 'F261101051'
      'F26_1101051'   → 'F261101051'
    """
    pn = pn.upper().strip()
    # Удаляем все разделители и пробельные символы
    pn = re.sub(r'[\s\-–—_\./,;:]+', '', pn)
    return pn


def _fuzzy_similarity(a: str, b: str) -> float:
    """Вычислить сходство двух НОРМАЛИЗОВАННЫХ парт-номеров.

    Возвращает 0.0..1.0. Использует многоступенчатую проверку
    для предотвращения ложных срабатываний.

    Критерии:
      1. Первый символ должен совпадать (префикс семейства деталей).
      2. Разница длин не более 25%.
      3. Общий префикс >= 2 символов.
      4. SequenceMatcher ratio + бонус за длинный общий префикс.
    """
    if a == b:
        return 1.0
    if not a or not b:
        return 0.0

    # Критерий 1: первый символ
    if a[0] != b[0]:
        return 0.0

    # Критерий 2: разница длин
    len_a, len_b = len(a), len(b)
    max_len = max(len_a, len_b)
    if max_len == 0:
        return 0.0
    len_diff_ratio = abs(len_a - len_b) / max_len
    if len_diff_ratio > _FUZZY_MAX_LEN_DIFF_RATIO:
        return 0.0

    # Критерий 3: общий префикс
    common_prefix = 0
    for ca, cb in zip(a, b):
        if ca == cb:
            common_prefix += 1
        else:
            break
    if common_prefix < 2:
        return 0.0

    # Критерий 4: SequenceMatcher
    ratio = SequenceMatcher(None, a, b).ratio()

    # Бонус за длинный общий префикс (повышает уверенность)
    if common_prefix >= 4 and ratio > 0.75:
        boost = min(0.08, common_prefix * 0.01)
        ratio = min(1.0, ratio + boost)

    return ratio


def _find_fuzzy_matches(
    source_parts: Dict[str, float],
    target_parts: Dict[str, PartInfo],
) -> Dict[str, Tuple[str, float]]:
    """Найти нечёткие совпадения между двумя множествами парт-номеров.

    Для каждого номера из source ищет лучший похожий номер в target.
    Использует несколько порогов в зависимости от длины номера:
      - Короткие (< 10 симв.): threshold = 0.90 (строже — риск ложных срабатываний)
      - Средние (10-16): threshold = 0.88
      - Длинные (> 16): threshold = 0.85

    Args:
        source_parts: {part_number: qty} — откуда ищем (например, карты).
        target_parts: {part_number: PartInfo} — где ищем (например, BOM).

    Returns:
        {source_part: (matched_target_part, similarity)}
        Только совпадения выше порога.
    """
    if not source_parts or not target_parts:
        return {}

    # Предварительно нормализуем target для быстрого поиска
    target_normalized: Dict[str, List[Tuple[str, str]]] = {}  # norm → [(original, norm)]
    for tpn in target_parts:
        norm = _normalize_for_fuzzy(tpn)
        if len(norm) >= _FUZZY_MIN_LENGTH:
            target_normalized.setdefault(norm, []).append((tpn, norm))

    # Для быстрого исключения — индексируем по первому символу
    target_by_first: Dict[str, List[Tuple[str, str]]] = {}
    for orig, norm in [(tpn, _normalize_for_fuzzy(tpn)) for tpn in target_parts]:
        if len(norm) >= _FUZZY_MIN_LENGTH:
            target_by_first.setdefault(norm[0], []).append((orig, norm))

    fuzzy_matches: Dict[str, Tuple[str, float]] = {}

    for spn, sqty in source_parts.items():
        s_norm = _normalize_for_fuzzy(spn)
        if len(s_norm) < _FUZZY_MIN_LENGTH:
            continue

        # Выбираем порог по длине
        if len(s_norm) < 10:
            threshold = _FUZZY_THRESHOLD_SHORT
        elif len(s_norm) <= 16:
            threshold = _FUZZY_THRESHOLD_MEDIUM
        else:
            threshold = _FUZZY_THRESHOLD_LONG

        # Сначала проверяем точное совпадение нормализованных номеров
        # НО только если оригинальные строки совпадают.
        # Если нормализация одинаковая, но оригиналы разные (разделители) -
        # это fuzzy-совпадение, не пропускаем.
        if s_norm in target_normalized:
            # Проверяем: есть ли среди target с такой нормализацией
            # оригинал, идентичный source?
            all_same = all(t_orig == spn for t_orig, _ in target_normalized[s_norm])
            if all_same:
                # Это не fuzzy — оригиналы совпадают (уже обработано)
                continue
            # Иначе: нормализация совпала, но оригиналы разные → fuzzy!

        # Ищем кандидатов по первому символу (быстрое отсечение)
        candidates = target_by_first.get(s_norm[0], [])
        if not candidates:
            continue

        best_match = ""
        best_sim = 0.0
        for t_orig, t_norm in candidates:
            sim = _fuzzy_similarity(s_norm, t_norm)
            if sim > best_sim:
                best_sim = sim
                best_match = t_orig

        if best_match and best_sim >= threshold:
            fuzzy_matches[spn] = (best_match, best_sim)

    return fuzzy_matches


class DiscrepancyType:
    """Типы расхождений."""
    ONLY_IN_BOM = "Только в BOM"
    ONLY_IN_CARDS = "Только в Картах"
    QUANTITY_MISMATCH = "Конфликт количества"
    FUZZY_IN_BOM = "Возможное совпадение (→ BOM)"
    FUZZY_IN_CARDS = "Возможное совпадение (→ Карты)"


@dataclass
class Discrepancy:
    """Одно расхождение между BOM и операционными картами."""
    part_number: str
    name_cn: str
    name_en: str
    qty_bom: float
    qty_cards: float
    card_numbers: List[str]  # Номера карт, где встречается деталь
    discrepancy_type: str
    fuzzy_match: str = ""  # Парт-номер, с которым сработал нечёткий поиск

    def __str__(self) -> str:
        fuzzy = f" (≈ {self.fuzzy_match})" if self.fuzzy_match else ""
        return (
            f"[{self.discrepancy_type}] {self.part_number}{fuzzy}: "
            f"BOM={self.qty_bom}, Карты={self.qty_cards}"
        )


@dataclass
class ComparisonResult:
    """Результат сверки BOM и операционных карт."""
    discrepancies: List[Discrepancy]
    total_bom_parts: int = 0
    total_cards_parts: int = 0
    matched_parts: int = 0
    bom_config_name: str = ""
    fuzzy_matches_found: int = 0


def compare(bom_parts: Dict[str, PartInfo],
            cards_data: CardsData,
            config_name: str = "",
            use_fuzzy: bool = True) -> ComparisonResult:
    """Произвести сверку BOM и операционных карт.

    Args:
        bom_parts: Детали для выбранной комплектации (из BOM).
        cards_data: Агрегированные данные из операционных карт.
        config_name: Название выбранной комплектации (для отчёта).
        use_fuzzy: Включить нечёткий поиск парт-номеров.

    Returns:
        ComparisonResult со списком расхождений.
    """
    discrepancies: List[Discrepancy] = []
    bom_part_numbers = set(bom_parts.keys())
    cards_part_numbers = set(cards_data.all_parts.keys())
    fuzzy_matches_found = 0

    # ── Нечёткий поиск (до основной сверки) ──
    fuzzy_cards_to_bom: Dict[str, Tuple[str, float]] = {}
    fuzzy_bom_to_cards: Dict[str, Tuple[str, float]] = {}
    if use_fuzzy:
        # Детали из карт, которых нет в BOM → ищем похожие в BOM
        only_in_cards_set = cards_part_numbers - bom_part_numbers
        cards_parts_for_fuzzy = {pn: cards_data.all_parts[pn] for pn in only_in_cards_set}
        fuzzy_cards_to_bom = _find_fuzzy_matches(cards_parts_for_fuzzy, bom_parts)

        # Детали из BOM, которых нет в картах → ищем похожие в картах
        only_in_bom_set = bom_part_numbers - cards_part_numbers
        # Конвертируем PartInfo → Dict для _find_fuzzy_matches
        bom_for_fuzzy = {pn: pi.quantity for pn, pi in bom_parts.items() if pn in only_in_bom_set}
        cards_for_fuzzy = {pn: qty for pn, qty in cards_data.all_parts.items()}
        # _find_fuzzy_matches ожидает target_parts: Dict[str, PartInfo],
        # но для обратного поиска нам нужны просто парт-номера.
        # Создаём фиктивные PartInfo для карт (только part_number).
        cards_as_partinfo = {pn: PartInfo(part_number=pn, name_cn="", name_en="") 
                             for pn in cards_data.all_parts}
        fuzzy_bom_to_cards = _find_fuzzy_matches(bom_for_fuzzy, cards_as_partinfo)

    # 1. Только в BOM (есть в BOM, но нет в картах)
    only_in_bom = bom_part_numbers - cards_part_numbers
    for part_no in sorted(only_in_bom):
        part = bom_parts[part_no]
        matched = fuzzy_bom_to_cards.get(part_no)
        if matched:
            matched_pn, _ = matched
            discrepancies.append(Discrepancy(
                part_number=part_no,
                name_cn=part.name_cn,
                name_en=part.name_en,
                qty_bom=part.quantity,
                qty_cards=cards_data.all_parts.get(matched_pn, 0.0),
                card_numbers=_get_card_numbers(matched_pn, cards_data),
                discrepancy_type=DiscrepancyType.FUZZY_IN_BOM,
                fuzzy_match=matched_pn,
            ))
            fuzzy_matches_found += 1
        else:
            discrepancies.append(Discrepancy(
                part_number=part_no,
                name_cn=part.name_cn,
                name_en=part.name_en,
                qty_bom=part.quantity,
                qty_cards=0.0,
                card_numbers=[],
                discrepancy_type=DiscrepancyType.ONLY_IN_BOM,
            ))

    # 2. Только в Картах (есть в картах, но нет в BOM для этой комплектации)
    only_in_cards = cards_part_numbers - bom_part_numbers
    for part_no in sorted(only_in_cards):
        qty = cards_data.all_parts[part_no]
        card_numbers = _get_card_numbers(part_no, cards_data)
        matched = fuzzy_cards_to_bom.get(part_no)
        if matched:
            matched_pn, _ = matched
            matched_part = bom_parts.get(matched_pn)
            discrepancies.append(Discrepancy(
                part_number=part_no,
                name_cn=matched_part.name_cn if matched_part else "",
                name_en=matched_part.name_en if matched_part else "",
                qty_bom=matched_part.quantity if matched_part else 0.0,
                qty_cards=qty,
                card_numbers=card_numbers,
                discrepancy_type=DiscrepancyType.FUZZY_IN_CARDS,
                fuzzy_match=matched_pn,
            ))
            fuzzy_matches_found += 1
        else:
            discrepancies.append(Discrepancy(
                part_number=part_no,
                name_cn="",
                name_en="",
                qty_bom=0.0,
                qty_cards=qty,
                card_numbers=card_numbers,
                discrepancy_type=DiscrepancyType.ONLY_IN_CARDS,
            ))

    # 3. Конфликт количества (количество не совпадает)
    common_parts = bom_part_numbers & cards_part_numbers
    matched = 0
    for part_no in sorted(common_parts):
        bom_qty = bom_parts[part_no].quantity
        cards_qty = cards_data.all_parts[part_no]

        if abs(bom_qty - cards_qty) > 0.001:  # Допуск на погрешность float
            part = bom_parts[part_no]
            card_numbers = _get_card_numbers(part_no, cards_data)
            discrepancies.append(Discrepancy(
                part_number=part_no,
                name_cn=part.name_cn,
                name_en=part.name_en,
                qty_bom=bom_qty,
                qty_cards=cards_qty,
                card_numbers=card_numbers,
                discrepancy_type=DiscrepancyType.QUANTITY_MISMATCH,
            ))
        else:
            matched += 1

    # Сортируем: конфликты → fuzzy → только BOM → только карты
    type_order = {
        DiscrepancyType.QUANTITY_MISMATCH: 0,
        DiscrepancyType.FUZZY_IN_BOM: 1,
        DiscrepancyType.FUZZY_IN_CARDS: 2,
        DiscrepancyType.ONLY_IN_BOM: 3,
        DiscrepancyType.ONLY_IN_CARDS: 4,
    }
    discrepancies.sort(key=lambda d: (type_order.get(d.discrepancy_type, 99), d.part_number))

    logger.info(f"Сверка завершена:")
    logger.info(f"  Деталей в BOM: {len(bom_parts)}")
    logger.info(f"  Деталей в картах: {len(cards_part_numbers)}")
    logger.info(f"  Совпало: {matched}")
    logger.info(f"  Расхождений: {len(discrepancies)}")
    logger.info(f"    - Только в BOM: {len(only_in_bom)}")
    logger.info(f"    - Только в картах: {len(only_in_cards)}")
    logger.info(f"    - Конфликт количества: {len(common_parts) - matched}")
    if use_fuzzy:
        logger.info(f"    - Fuzzy-совпадений: {fuzzy_matches_found}")

    return ComparisonResult(
        discrepancies=discrepancies,
        total_bom_parts=len(bom_parts),
        total_cards_parts=len(cards_part_numbers),
        matched_parts=matched,
        bom_config_name=config_name,
        fuzzy_matches_found=fuzzy_matches_found,
    )


def _get_card_numbers(part_no: str, cards_data: CardsData) -> List[str]:
    """Получить уникальные номера карт, где встречается деталь."""
    sources = cards_data.part_sources.get(part_no, [])
    seen: set = set()
    card_nums: List[str] = []
    for card_number, _, _ in sources:
        if card_number not in seen:
            seen.add(card_number)
            # Берём короткое имя файла
            card_nums.append(card_number)
    return card_nums


def format_discrepancy_report(comparison: ComparisonResult) -> str:
    """Сформировать текстовый отчёт о расхождениях.

    Args:
        comparison: Результат сверки.

    Returns:
        Текстовый отчёт.
    """
    lines = [
        "=" * 80,
        f"ОТЧЁТ О РАСХОЖДЕНИЯХ",
        f"Комплектация: {comparison.bom_config_name}",
        "=" * 80,
        "",
        f"Всего деталей в BOM: {comparison.total_bom_parts}",
        f"Всего деталей в картах: {comparison.total_cards_parts}",
        f"Совпало: {comparison.matched_parts}",
        f"Расхождений: {len(comparison.discrepancies)}",
        "",
    ]

    if not comparison.discrepancies:
        lines.append("✅ Расхождений не найдено!")
        return "\n".join(lines)

    # Группируем по типу
    for dtype in [DiscrepancyType.QUANTITY_MISMATCH,
                  DiscrepancyType.FUZZY_IN_BOM,
                  DiscrepancyType.FUZZY_IN_CARDS,
                  DiscrepancyType.ONLY_IN_BOM,
                  DiscrepancyType.ONLY_IN_CARDS]:
        type_disc = [d for d in comparison.discrepancies if d.discrepancy_type == dtype]
        if not type_disc:
            continue
        lines.append(f"\n--- {dtype} ({len(type_disc)} шт.) ---")
        lines.append(f"{'Парт-номер':<20} {'Название (CN)':<30} {'BOM':<8} {'Карты':<8} {'Карты':<30}")
        lines.append("-" * 96)
        for d in type_disc:
            cards_str = ", ".join(d.card_numbers[:3])
            if len(d.card_numbers) > 3:
                cards_str += f" (+{len(d.card_numbers) - 3})"
            lines.append(f"{d.part_number:<20} {d.name_cn[:28]:<30} {d.qty_bom:<8} {d.qty_cards:<8} {cards_str:<30}")

    return "\n".join(lines)
