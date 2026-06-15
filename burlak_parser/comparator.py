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
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from burlak_parser.bom_parser import BOMData, PartInfo
from burlak_parser.card_parser import CardsData

logger = logging.getLogger(__name__)


class DiscrepancyType:
    """Типы расхождений."""
    ONLY_IN_BOM = "Только в BOM"
    ONLY_IN_CARDS = "Только в Картах"
    QUANTITY_MISMATCH = "Конфликт количества"


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

    def __str__(self) -> str:
        return (
            f"[{self.discrepancy_type}] {self.part_number}: "
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


def compare(bom_parts: Dict[str, PartInfo],
            cards_data: CardsData,
            config_name: str = "") -> ComparisonResult:
    """Произвести сверку BOM и операционных карт.

    Args:
        bom_parts: Детали для выбранной комплектации (из BOM).
        cards_data: Агрегированные данные из операционных карт.
        config_name: Название выбранной комплектации (для отчёта).

    Returns:
        ComparisonResult со списком расхождений.
    """
    discrepancies: List[Discrepancy] = []
    bom_part_numbers = set(bom_parts.keys())
    cards_part_numbers = set(cards_data.all_parts.keys())

    # 1. Только в BOM (есть в BOM, но нет в картах)
    only_in_bom = bom_part_numbers - cards_part_numbers
    for part_no in sorted(only_in_bom):
        part = bom_parts[part_no]
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
        # Названия нет в BOM, берём из карт
        card_numbers = _get_card_numbers(part_no, cards_data)
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

    # Сортируем: сначала конфликты количества, потом только в BOM, потом только в картах
    type_order = {
        DiscrepancyType.QUANTITY_MISMATCH: 0,
        DiscrepancyType.ONLY_IN_BOM: 1,
        DiscrepancyType.ONLY_IN_CARDS: 2,
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

    return ComparisonResult(
        discrepancies=discrepancies,
        total_bom_parts=len(bom_parts),
        total_cards_parts=len(cards_part_numbers),
        matched_parts=matched,
        bom_config_name=config_name,
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
