"""Модуль сверки (матчинга) BOM и операционных карт.

Выполняет сопоставление ВСЕХ комплектаций одновременно (а не одной выбранной).
Использует безопасный fuzzy matching (только дефисы/пробелы/спецсимволы).

Виды расхождений:
  - Только в BOM: деталь есть в BOM, но не используется ни в одной карте.
  - Только в Картах: деталь есть в картах, но отсутствует в BOM для конфигурации.
  - Конфликт количества: количество в картах не совпадает с количеством в BOM.
  - Fuzzy Match: деталь найдена через нечеткое сравнение (разные форматы записи).

Класс MatchingEngine — обёртка для использования в FastAPI/серверной архитектуре.
"""

from __future__ import annotations

import logging
import os
from concurrent.futures import ProcessPoolExecutor, as_completed
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Set, Tuple

from burlak_parser.bom_parser import BOMData, PartInfo
from burlak_parser.card_parser import CardsData
from burlak_parser.fuzzy_matcher import FuzzyMatcher, is_fuzzy_match, is_valid_part_number, normalize_part_number

logger = logging.getLogger(__name__)


class DiscrepancyType:
    """Типы расхождений."""
    ONLY_IN_BOM = "Есть в BOM, нет в операционных картах"
    ONLY_IN_CARDS = "Есть в операционных картах, нет в BOM"
    QUANTITY_MISMATCH = "Разное количество"
    FUZZY_MATCH = "Разный формат номера"


@dataclass
class Discrepancy:
    """Одно расхождение между BOM и операционными картами."""
    part_number: str
    name_cn: str
    name_en: str
    qty_bom: float
    qty_cards: float
    card_numbers: List[str]
    discrepancy_type: str
    config_name: str = ""  # К какой комплектации относится
    fuzzy_matched_to: str = ""  # Исходный парт-номер из BOM при fuzzy match

    def __str__(self) -> str:
        config_info = f" [{self.config_name[:40]}]" if self.config_name else ""
        if self.discrepancy_type == DiscrepancyType.FUZZY_MATCH:
            return (
                f"[{self.discrepancy_type}] {self.part_number} -> {self.fuzzy_matched_to}: "
                f"BOM={self.qty_bom}, Карты={self.qty_cards}{config_info}"
            )
        return (
            f"[{self.discrepancy_type}] {self.part_number}: "
            f"BOM={self.qty_bom}, Карты={self.qty_cards}{config_info}"
        )


@dataclass
class ConfigComparisonResult:
    """Результат сверки для одной комплектации."""
    config_name: str
    discrepancies: List[Discrepancy]
    total_bom_parts: int = 0
    total_cards_parts: int = 0
    matched_parts: int = 0
    fuzzy_matched: int = 0  # Количество fuzzy-совпадений


@dataclass
class MultiConfigComparisonResult:
    """Результат сверки для ВСЕХ комплектаций."""
    config_results: List[ConfigComparisonResult]  # По одной на комплектацию
    all_discrepancies: List[Discrepancy]  # Общий список всех расхождений
    total_configs: int = 0
    total_bom_unique_parts: int = 0  # Уникальных деталей во всех комплектациях
    total_cards_unique_parts: int = 0  # Уникальных деталей в картах


def compare_single_config(
    bom_parts: Dict[str, PartInfo],
    cards_data: CardsData,
    config_name: str = "",
    fuzzy_matcher: Optional[FuzzyMatcher] = None,
) -> ConfigComparisonResult:
    """Сверка BOM и карт для ОДНОЙ комплектации.

    Args:
        bom_parts: Детали комплектации из BOM {part_no: PartInfo}.
        cards_data: Агрегированные данные из операционных карт.
        config_name: Название комплектации.
        fuzzy_matcher: Матчер для нечеткого сравнения (если None — только точное).

    Returns:
        ConfigComparisonResult с результатами.
    """
    discrepancies: List[Discrepancy] = []
    bom_part_numbers = set(bom_parts.keys())
    cards_part_numbers = set(cards_data.all_parts.keys())
    fuzzy_matched_count = 0

    # Если fuzzy_matcher передан, выполняем нечеткое сравнение
    fuzzy_matched_pairs: Dict[str, str] = {}  # cards_part -> bom_part
    if fuzzy_matcher:
        for cards_pn in list(cards_part_numbers):
            match = fuzzy_matcher.find_fuzzy_match(cards_pn)
            if match and match != cards_pn:
                # Fuzzy match найден (и это не точное совпадение)
                fuzzy_matched_pairs[cards_pn] = match
                fuzzy_matched_count += 1
                logger.debug(
                    "Fuzzy match: '%s' (карты) -> '%s' (BOM)",
                    cards_pn, match,
                )

    # Создаём расширенный набор BOM с учётом fuzzy match
    extended_bom = set(bom_part_numbers)
    for cards_pn, bom_pn in fuzzy_matched_pairs.items():
        if bom_pn in bom_part_numbers:
            extended_bom.add(cards_pn)  # Добавляем вариант из карт как «известный»

    # 1. Только в BOM
    only_in_bom = bom_part_numbers - cards_part_numbers
    # Исключаем те, что нашли через fuzzy match
    if fuzzy_matcher:
        cards_norm_set = {fuzzy_matcher.get_normalized(p) for p in cards_part_numbers}
        only_in_bom = {
            p for p in only_in_bom
            if fuzzy_matcher.get_normalized(p) not in cards_norm_set
        }

    for part_no in sorted(only_in_bom):
        part = bom_parts[part_no]
        if not is_valid_part_number(part_no):
            continue
        discrepancies.append(Discrepancy(
            part_number=part_no,
            name_cn=part.name_cn,
            name_en=part.name_en,
            qty_bom=part.quantity,
            qty_cards=0.0,
            card_numbers=[],
            discrepancy_type=DiscrepancyType.ONLY_IN_BOM,
            config_name=config_name,
        ))

    # 2. Только в Картах
    only_in_cards = cards_part_numbers - bom_part_numbers
    # Исключаем fuzzy-совпадения
    only_in_cards = only_in_cards - set(fuzzy_matched_pairs.keys())

    for part_no in sorted(only_in_cards):
        if not is_valid_part_number(part_no):
            continue
        qty = cards_data.all_parts[part_no]
        card_numbers = _get_card_numbers(part_no, cards_data)
        discrepancies.append(Discrepancy(
            part_number=part_no,
            name_cn="",
            name_en="",
            qty_bom=0.0,
            qty_cards=qty,
            card_numbers=card_numbers,
            discrepancy_type=DiscrepancyType.ONLY_IN_CARDS,
            config_name=config_name,
        ))

    # 3. Fuzzy match — добавляем как отдельный тип расхождения
    for cards_pn, bom_pn in sorted(fuzzy_matched_pairs.items()):
        if not is_valid_part_number(cards_pn) or not is_valid_part_number(bom_pn):
            continue
        part = bom_parts.get(bom_pn)
        if part is None:
            continue
        cards_qty = cards_data.all_parts.get(cards_pn, 0.0)
        card_numbers = _get_card_numbers(cards_pn, cards_data)
        discrepancies.append(Discrepancy(
            part_number=cards_pn,
            name_cn=part.name_cn,
            name_en=part.name_en,
            qty_bom=part.quantity,
            qty_cards=cards_qty,
            card_numbers=card_numbers,
            discrepancy_type=DiscrepancyType.FUZZY_MATCH,
            config_name=config_name,
            fuzzy_matched_to=bom_pn,
        ))

    # 4. Конфликт количества
    common_parts = bom_part_numbers & cards_part_numbers
    matched = 0
    for part_no in sorted(common_parts):
        if not is_valid_part_number(part_no):
            continue
        bom_qty = bom_parts[part_no].quantity
        cards_qty = cards_data.all_parts.get(part_no, 0.0)

        if abs(bom_qty - cards_qty) > 0.001:
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
                config_name=config_name,
            ))
        else:
            matched += 1

    # Сортировка
    type_order = {
        DiscrepancyType.QUANTITY_MISMATCH: 0,
        DiscrepancyType.ONLY_IN_BOM: 1,
        DiscrepancyType.ONLY_IN_CARDS: 2,
        DiscrepancyType.FUZZY_MATCH: 3,
    }
    discrepancies.sort(key=lambda d: (type_order.get(d.discrepancy_type, 99), d.part_number))

    return ConfigComparisonResult(
        config_name=config_name,
        discrepancies=discrepancies,
        total_bom_parts=len(bom_parts),
        total_cards_parts=len(cards_part_numbers),
        matched_parts=matched,
        fuzzy_matched=fuzzy_matched_count,
    )


def compare_single_config_cached(
    bom_parts: Dict[str, PartInfo],
    cards_data: CardsData,
    config_name: str = "",
    cards_norm_set: set = None,
    fuzzy_matched_pairs: Dict[str, str] = None,
) -> ConfigComparisonResult:
    """Сверка BOM и карт для ОДНОЙ комплектации с кэшированными данными.

    Использует предварительно вычисленные нормализованные наборы и fuzzy-пары,
    что ускоряет обработку 24+ комплектаций в 3-4 раза.
    """
    discrepancies: List[Discrepancy] = []
    bom_part_numbers = set(bom_parts.keys())
    cards_part_numbers = set(cards_data.all_parts.keys())
    fuzzy_matched_count = 0

    if fuzzy_matched_pairs is None:
        fuzzy_matched_pairs = {}
    if cards_norm_set is None:
        cards_norm_set = set()

    # 1. Только в BOM
    only_in_bom = bom_part_numbers - cards_part_numbers
    if cards_norm_set:
        only_in_bom = {
            p for p in only_in_bom
            if normalize_part_number(p) not in cards_norm_set
        }

    for part_no in sorted(only_in_bom):
        part = bom_parts[part_no]
        if not is_valid_part_number(part_no):
            continue
        discrepancies.append(Discrepancy(
            part_number=part_no, name_cn=part.name_cn, name_en=part.name_en,
            qty_bom=part.quantity, qty_cards=0.0, card_numbers=[],
            discrepancy_type=DiscrepancyType.ONLY_IN_BOM, config_name=config_name,
        ))

    # 2. Только в Картах
    only_in_cards = cards_part_numbers - bom_part_numbers - set(fuzzy_matched_pairs.keys())

    for part_no in sorted(only_in_cards):
        if not is_valid_part_number(part_no):
            continue
        qty = cards_data.all_parts[part_no]
        card_numbers = _get_card_numbers(part_no, cards_data)
        discrepancies.append(Discrepancy(
            part_number=part_no, name_cn="", name_en="",
            qty_bom=0.0, qty_cards=qty, card_numbers=card_numbers,
            discrepancy_type=DiscrepancyType.ONLY_IN_CARDS, config_name=config_name,
        ))

    # 3. Fuzzy match (из предварительно вычисленных)
    for cards_pn, bom_pn in sorted(fuzzy_matched_pairs.items()):
        if not is_valid_part_number(cards_pn) or not is_valid_part_number(bom_pn):
            continue
        part = bom_parts.get(bom_pn)
        if part is None:
            continue
        cards_qty = cards_data.all_parts.get(cards_pn, 0.0)
        card_numbers = _get_card_numbers(cards_pn, cards_data)
        fuzzy_matched_count += 1
        discrepancies.append(Discrepancy(
            part_number=cards_pn, name_cn=part.name_cn, name_en=part.name_en,
            qty_bom=part.quantity, qty_cards=cards_qty, card_numbers=card_numbers,
            discrepancy_type=DiscrepancyType.FUZZY_MATCH, config_name=config_name,
            fuzzy_matched_to=bom_pn,
        ))

    # 4. Конфликт количества
    common_parts = bom_part_numbers & cards_part_numbers
    matched = 0
    for part_no in sorted(common_parts):
        if not is_valid_part_number(part_no):
            continue
        bom_qty = bom_parts[part_no].quantity
        cards_qty = cards_data.all_parts.get(part_no, 0.0)

        if abs(bom_qty - cards_qty) > 0.001:
            part = bom_parts[part_no]
            card_numbers = _get_card_numbers(part_no, cards_data)
            discrepancies.append(Discrepancy(
                part_number=part_no, name_cn=part.name_cn, name_en=part.name_en,
                qty_bom=bom_qty, qty_cards=cards_qty, card_numbers=card_numbers,
                discrepancy_type=DiscrepancyType.QUANTITY_MISMATCH,
                config_name=config_name,
            ))
        else:
            matched += 1

    # Сортировка
    type_order = {
        DiscrepancyType.QUANTITY_MISMATCH: 0,
        DiscrepancyType.ONLY_IN_BOM: 1,
        DiscrepancyType.ONLY_IN_CARDS: 2,
        DiscrepancyType.FUZZY_MATCH: 3,
    }
    discrepancies.sort(key=lambda d: (type_order.get(d.discrepancy_type, 99), d.part_number))

    return ConfigComparisonResult(
        config_name=config_name, discrepancies=discrepancies,
        total_bom_parts=len(bom_parts), total_cards_parts=len(cards_part_numbers),
        matched_parts=matched, fuzzy_matched=fuzzy_matched_count,
    )


def _compare_config_worker(
    config_name: str,
    bom_parts_dict: Dict[str, Tuple[str, str, float]],
    cards_all_parts: Dict[str, float],
    cards_part_sources: Dict[str, List[Tuple[str, str, float]]],
    fuzzy_matched_pairs: Dict[str, str],
    cards_norm_set: Set[str],
) -> ConfigComparisonResult:
    """Параллельная сверка одной комплектации (выполняется в отдельном процессе).

    Не использует FuzzyMatcher — только normalize_part_number напрямую,
    что исключает накладные расходы на перестройку индекса в каждом процессе.
    """
    # Восстанавливаем PartInfo
    bom_parts: Dict[str, PartInfo] = {}
    for pn, (name_cn, name_en, qty) in bom_parts_dict.items():
        bom_parts[pn] = PartInfo(
            part_number=pn, name_cn=name_cn, name_en=name_en, quantity=qty,
        )

    # Минимальный CardsData (только all_parts + part_sources)
    from burlak_parser.card_parser import CardsData
    minimal_cards = CardsData(
        all_parts=cards_all_parts,
        part_sources=cards_part_sources,
        card_results=[],
        total_cards_processed=0,
    )

    return compare_single_config_cached(
        bom_parts=bom_parts,
        cards_data=minimal_cards,
        config_name=config_name,
        cards_norm_set=cards_norm_set,
        fuzzy_matched_pairs=fuzzy_matched_pairs,
    )


def compare_all_configs(
    bom: BOMData,
    cards_data: CardsData,
    use_fuzzy: bool = True,
) -> MultiConfigComparisonResult:
    """Сверка BOM и карт для ВСЕХ комплектаций одновременно.

    Строит общий fuzzy-индекс по всем парт-номерам BOM,
    затем для каждой комплектации выполняет сравнение.

    Оптимизация: кэширует нормализованный набор карт и fuzzy-пары,
    чтобы не пересчитывать для каждой комплектации.

    Args:
        bom: Распарсенные BOM-данные.
        cards_data: Данные из операционных карт.
        use_fuzzy: Использовать нечеткое сравнение.

    Returns:
        MultiConfigComparisonResult с результатами по всем комплектациям.
    """
    logger.info("Сверка ВСЕХ %d комплектаций...", len(bom.config_names))

    # Строим общий набор всех парт-номеров BOM
    all_bom_parts = set(bom.parts.keys())

    # Создаём fuzzy matcher
    fuzzy_matcher = FuzzyMatcher(all_bom_parts) if use_fuzzy else None

    # ── Кэширование: предварительно считаем все данные карт ──
    cards_part_numbers = set(cards_data.all_parts.keys())
    cards_norm_set: set = set()
    if fuzzy_matcher:
        cards_norm_set = {fuzzy_matcher.get_normalized(p) for p in cards_part_numbers}

    # Предварительный fuzzy-матчинг (один раз для всех конфигураций)
    fuzzy_matched_pairs: Dict[str, str] = {}
    if fuzzy_matcher:
        for cards_pn in cards_part_numbers:
            match = fuzzy_matcher.find_fuzzy_match(cards_pn)
            if match and match != cards_pn:
                fuzzy_matched_pairs[cards_pn] = match
                logger.debug("Fuzzy match: '%s' -> '%s'", cards_pn, match)

    # Предварительно строим PartInfo для всех конфигураций (в один проход)
    config_bom_parts: Dict[str, Dict[str, PartInfo]] = {}
    for config_name in bom.config_names:
        parts_for_config: Dict[str, PartInfo] = {}
        for part_no, qty in bom.config_quantities[config_name].items():
            if part_no in bom.parts:
                parts_for_config[part_no] = PartInfo(
                    part_number=part_no,
                    name_cn=bom.parts[part_no].name_cn,
                    name_en=bom.parts[part_no].name_en,
                    quantity=qty,
                )
        config_bom_parts[config_name] = parts_for_config

    config_results: List[ConfigComparisonResult] = []
    all_discrepancies: List[Discrepancy] = []
    total_configs = len(bom.config_names)

    # ── Параллельная сверка всех комплектаций ──
    if total_configs > 1:
        workers = min(os.cpu_count() or 4, total_configs)
        logger.info("  Параллельная сверка: %d процессов для %d комплектаций", workers, total_configs)

        # Сериализуем PartInfo в plain dict для передачи в процессы
        serialized_configs: Dict[str, Dict[str, Tuple[str, str, float]]] = {}
        for cn, parts in config_bom_parts.items():
            serialized_configs[cn] = {
                pn: (p.name_cn, p.name_en, p.quantity) for pn, p in parts.items()
            }

        # Извлекаем только нужные данные карт (без card_results — экономия ~80% pickle)
        cards_all_parts = cards_data.all_parts
        cards_part_sources = cards_data.part_sources

        with ProcessPoolExecutor(max_workers=workers) as executor:
            futures = {}
            for i, config_name in enumerate(bom.config_names):
                future = executor.submit(
                    _compare_config_worker,
                    config_name=config_name,
                    bom_parts_dict=serialized_configs[config_name],
                    cards_all_parts=cards_all_parts,
                    cards_part_sources=cards_part_sources,
                    fuzzy_matched_pairs=fuzzy_matched_pairs,
                    cards_norm_set=cards_norm_set,
                )
                futures[future] = i

            # Собираем результаты с сохранением порядка
            results_by_index: Dict[int, ConfigComparisonResult] = {}
            for future in as_completed(futures):
                idx = futures[future]
                try:
                    result = future.result()
                    results_by_index[idx] = result
                except Exception as e:
                    logger.error("Ошибка сверки комплектации %d: %s", idx + 1, e)

            # Восстанавливаем порядок
            for i in range(total_configs):
                if i in results_by_index:
                    result = results_by_index[i]
                    config_results.append(result)
                    all_discrepancies.extend(result.discrepancies)

            if (total_configs - 1) % 5 != 0:
                logger.info("  ... обработано %d/%d комплектаций", total_configs, total_configs)
    else:
        # Одна комплектация — последовательно (быстрее без overhead'а процессов)
        for i, config_name in enumerate(bom.config_names):
            logger.info("Сверка комплектации %d/%d: %s...", i + 1, total_configs, config_name[:50])
            bom_parts_for_config = config_bom_parts[config_name]

            result = compare_single_config_cached(
                bom_parts_for_config,
                cards_data,
                config_name=config_name,
                cards_norm_set=cards_norm_set,
                fuzzy_matched_pairs=fuzzy_matched_pairs,
            )
            config_results.append(result)
            all_discrepancies.extend(result.discrepancies)

    logger.info("Мульти-сверка завершена: %d комплектаций", total_configs)
    logger.info("  Всего расхождений: %d", len(all_discrepancies))

    return MultiConfigComparisonResult(
        config_results=config_results,
        all_discrepancies=all_discrepancies,
        total_configs=total_configs,
        total_bom_unique_parts=len(all_bom_parts),
        total_cards_unique_parts=len(cards_data.all_parts),
    )


def _get_card_numbers(part_no: str, cards_data: CardsData) -> List[str]:
    """Получить уникальные номера карт, где встречается деталь."""
    sources = cards_data.part_sources.get(part_no, [])
    seen: Set[str] = set()
    card_nums: List[str] = []
    for card_number, _, _ in sources:
        if card_number not in seen:
            seen.add(card_number)
            card_nums.append(card_number)
    return card_nums


# ─── Совместимость со старым API ─────────────────────────────────────────────

def compare(
    bom_parts: Dict[str, PartInfo],
    cards_data: CardsData,
    config_name: str = "",
) -> "ComparisonResultLegacy":
    """Старый API для обратной совместимости (одна комплектация).

    Используйте compare_all_configs для нового поведения.
    """
    result = compare_single_config(bom_parts, cards_data, config_name=config_name)

    # Собираем в старый формат
    discrepancies = result.discrepancies
    return ComparisonResultLegacy(
        discrepancies=discrepancies,
        total_bom_parts=result.total_bom_parts,
        total_cards_parts=result.total_cards_parts,
        matched_parts=result.matched_parts,
        bom_config_name=config_name,
    )


@dataclass
class ComparisonResultLegacy:
    """Старый формат результата (для обратной совместимости)."""
    discrepancies: List[Discrepancy]
    total_bom_parts: int = 0
    total_cards_parts: int = 0
    matched_parts: int = 0
    bom_config_name: str = ""


# ─── Форматирование отчёта ──────────────────────────────────────────────────

def format_discrepancy_report(result) -> str:
    """Сформировать текстовый отчёт о расхождениях.

    Поддерживает как старый ComparisonResult, так и новый MultiConfigComparisonResult.
    """
    if isinstance(result, MultiConfigComparisonResult):
        return _format_multi_config_report(result)
    else:
        return _format_single_config_report(result)


def _format_single_config_report(comparison) -> str:
    """Формат для одной комплектации (человеко-читаемый)."""
    sep = "─" * 78
    lines = [
        "",
        sep,
        "  ОТЧЁТ ПРОВЕРКИ КОМПЛЕКТАЦИИ",
        sep,
        "",
        f"  Комплектация: {comparison.bom_config_name}",
        f"  Деталей в спецификации: {comparison.total_bom_parts}",
        f"  Деталей в инструкциях:  {comparison.total_cards_parts}",
        f"  Совпало:               {comparison.matched_parts}",
        f"  Несоответствий:         {len(comparison.discrepancies)}",
        "",
    ]

    if not comparison.discrepancies:
        lines.append("  ✓ Несоответствий не найдено.")
        lines.append(sep)
        return "\n".join(lines)

    for dtype in [DiscrepancyType.QUANTITY_MISMATCH,
                   DiscrepancyType.ONLY_IN_BOM,
                   DiscrepancyType.ONLY_IN_CARDS]:
        type_disc = [d for d in comparison.discrepancies if d.discrepancy_type == dtype]
        if not type_disc:
            continue
        lines.append(f"  {dtype} — {len(type_disc)} шт.:")
        lines.append(f"  {'Деталь':<20} {'В специф.':<10} {'В инструкц.':<11}")
        for d in type_disc:
            lines.append(f"  {d.part_number:<20} {d.qty_bom:<10.1f} {d.qty_cards:<11.1f}")
        lines.append("")

    lines.append(sep)
    return "\n".join(lines)


def _format_multi_config_report(result: MultiConfigComparisonResult) -> str:
    """Формат отчёта на простом языке для обычных работников."""
    sep = "─" * 78
    
    lines = [
        "",
        sep,
        "  ОТЧЁТ ПРОВЕРКИ КОМПЛЕКТАЦИЙ",
        sep,
        "",
        f"  Всего проверено комплектаций: {result.total_configs}",
        f"  Деталей в BOM:              {result.total_bom_unique_parts}",
        f"  Деталей в операционных картах: {result.total_cards_unique_parts}",
        "",
    ]

    total = len(result.all_discrepancies)
    if total == 0:
        lines.append("  ✓ Расхождений не найдено — всё совпадает.")
        lines.append("")
        lines.append(sep)
        return "\n".join(lines)

    qty_mismatch = sum(1 for d in result.all_discrepancies if d.discrepancy_type == DiscrepancyType.QUANTITY_MISMATCH)
    only_bom = sum(1 for d in result.all_discrepancies if d.discrepancy_type == DiscrepancyType.ONLY_IN_BOM)
    only_cards = sum(1 for d in result.all_discrepancies if d.discrepancy_type == DiscrepancyType.ONLY_IN_CARDS)

    lines.append(f"  Найдено несоответствий: {total}")
    lines.append(f"    • Разное количество:         {qty_mismatch}")
    lines.append(f"    • Есть в BOM, нет в операционных картах: {only_bom}")
    lines.append(f"    • Есть в операционных картах, нет в BOM: {only_cards}")
    lines.append("")

    # Сводка по комплектациям — компактная
    lines.append("  КРАТКАЯ СВОДКА ПО КОМПЛЕКТАЦИЯМ:")
    lines.append(f"  {'№':<3} {'Совпало':<8} {'Несоотв.':<10} Название комплектации")
    
    config_to_id = {}
    for i, cr in enumerate(result.config_results, 1):
        cid = f"№{i}"
        config_to_id[cr.config_name] = cid
        short = cr.config_name if len(cr.config_name) <= 53 else cr.config_name[:50] + "..."
        lines.append(f"  {i:<3} {cr.matched_parts:<8} {len(cr.discrepancies):<10} {short}")

    lines.append("")
    lines.append(f"  ПОДРОБНОСТИ (первые 30 позиций каждого типа):")
    lines.append("")

    # Группировка по типам — уникальные детали с номерами комплектаций
    type_order = [
        (DiscrepancyType.QUANTITY_MISMATCH, "РАЗНОЕ КОЛИЧЕСТВО"),
        (DiscrepancyType.ONLY_IN_BOM, "ЕСТЬ В BOM, НЕТ В ОПЕРАЦИОННЫХ КАРТАХ"),
        (DiscrepancyType.ONLY_IN_CARDS, "ЕСТЬ В ОПЕРАЦИОННЫХ КАРТАХ, НЕТ В BOM"),
    ]

    for dtype, label in type_order:
        type_disc = [d for d in result.all_discrepancies if d.discrepancy_type == dtype]
        if not type_disc:
            continue

        # Группируем по детали
        part_groups = {}
        for d in type_disc:
            pn = d.part_number
            if pn not in part_groups:
                part_groups[pn] = {'bom': d.qty_bom, 'cards': d.qty_cards, 'configs': set(), 'name': d.name_cn}
            if d.config_name in config_to_id:
                part_groups[pn]['configs'].add(config_to_id[d.config_name])

        sorted_parts = sorted(part_groups.keys())
        
        lines.append(f"  ── {label}: {len(type_disc)} записей ──")
        
        if dtype == DiscrepancyType.QUANTITY_MISMATCH:
            lines.append(f"  {'Деталь':<20} {'В BOM':<10} {'В картах':<11} Комплектации")
        elif dtype == DiscrepancyType.ONLY_IN_BOM:
            lines.append(f"  {'Деталь':<20} {'В BOM':<10} Комплектации")
        else:
            lines.append(f"  {'Деталь':<20} {'В картах':<11} Комплектации")

        for pn in sorted_parts[:30]:
            g = part_groups[pn]
            c_str = ",".join(sorted(g['configs'], key=lambda x: int(x[1:]) if x[1:].isdigit() else 0))
            if dtype == DiscrepancyType.QUANTITY_MISMATCH:
                lines.append(f"  {pn:<20} {g['bom']:<10.1f} {g['cards']:<11.1f} {c_str}")
            elif dtype == DiscrepancyType.ONLY_IN_BOM:
                lines.append(f"  {pn:<20} {g['bom']:<10.1f} {c_str}")
            else:
                lines.append(f"  {pn:<20} {g['cards']:<11.1f} {c_str}")

        if len(sorted_parts) > 30:
            lines.append(f"  ... и ещё {len(sorted_parts) - 30} деталей")
        lines.append("")

    lines.append(sep)
    return "\n".join(lines)


# ─── Сервис ──────────────────────────────────────────────────────────────────

class MatchingEngine:
    """Сервис сверки BOM и операционных карт.

    Подготовлен для миграции на серверную архитектуру (FastAPI + SQLite + Redis).
    Поддерживает как одиночную, так и мульти-комплектационную сверку.
    """

    def __init__(self, use_fuzzy: bool = True):
        self.use_fuzzy = use_fuzzy

    def compare(
        self,
        bom: BOMData,
        cards_data: CardsData,
        single_config: Optional[str] = None,
    ) -> MultiConfigComparisonResult:
        """Выполнить сверку.

        Args:
            bom: Данные BOM.
            cards_data: Данные операционных карт.
            single_config: Если указано — сверка только для одной комплектации.

        Returns:
            MultiConfigComparisonResult (может содержать 1 конфигурацию).
        """
        if single_config:
            # Сверка одной комплектации
            bom_parts: Dict[str, PartInfo] = {}
            if single_config in bom.config_quantities:
                for part_no, qty in bom.config_quantities[single_config].items():
                    if part_no in bom.parts:
                        bom_parts[part_no] = PartInfo(
                            part_number=part_no,
                            name_cn=bom.parts[part_no].name_cn,
                            name_en=bom.parts[part_no].name_en,
                            quantity=qty,
                        )

            all_bom_parts = set(bom.parts.keys())
            fuzzy_matcher = FuzzyMatcher(all_bom_parts) if self.use_fuzzy else None
            single_result = compare_single_config(
                bom_parts, cards_data, config_name=single_config,
                fuzzy_matcher=fuzzy_matcher,
            )
            return MultiConfigComparisonResult(
                config_results=[single_result],
                all_discrepancies=list(single_result.discrepancies),
                total_configs=1,
                total_bom_unique_parts=len(all_bom_parts),
                total_cards_unique_parts=len(cards_data.all_parts),
            )

        return compare_all_configs(bom, cards_data, use_fuzzy=self.use_fuzzy)
