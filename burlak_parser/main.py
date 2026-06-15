#!/usr/bin/env python3
"""Точка входа в систему автоматического разбора и сверки ведомостей материалов (BOM).

Использование:
  python -m burlak_parser.main --bom <file.xlsx> --cards <path> [--config <name>]
"""

from __future__ import annotations

import argparse
import logging
import os
import sys
import time
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from tqdm import tqdm

from burlak_parser.bom_parser import BOMData, PartInfo, get_config_quantities, parse_bom
from burlak_parser.card_parser import CardsData, parse_cards, split_cards_to_files
from burlak_parser.comparator import (
    ComparisonResult,
    DiscrepancyType,
    compare,
    format_discrepancy_report,
)
from burlak_parser.report_generator import (
    create_split_cards_archive,
    generate_discrepancy_report,
)

logger = logging.getLogger(__name__)


def setup_logging(verbose: bool = False) -> None:
    """Настроить логирование."""
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )


def select_config_interactive(bom: BOMData) -> str:
    """Интерактивный выбор комплектации из списка.

    Если доступна только одна комплектация, выбирает её автоматически.
    """
    configs = bom.config_names
    if not configs:
        print("❌ Нет доступных комплектаций в BOM-файле.")
        sys.exit(1)

    if len(configs) == 1:
        print(f"✅ Автоматически выбрана единственная комплектация: {configs[0][:60]}")
        return configs[0]

    print(f"\n{'=' * 60}")
    print(f"Доступные комплектации ({len(configs)} шт.):")
    print(f"{'=' * 60}")

    # Показываем только первые 30 для выбора, остальные скрываем
    display_configs = configs[:30]
    for i, name in enumerate(display_configs, 1):
        # Укорачиваем для отображения
        display_name = name if len(name) <= 70 else name[:67] + "..."
        print(f"  {i:3d}. {display_name}")

    while True:
        try:
            choice = input(f"\nВыберите комплектацию (1-{len(display_configs)}): ").strip()
            idx = int(choice) - 1
            if 0 <= idx < len(display_configs):
                return configs[idx]
            else:
                print(f"❌ Введите число от 1 до {len(display_configs)}")
        except ValueError:
            print("❌ Введите корректное число")


def run_pipeline(bom_path: str,
                 cards_path: str,
                 config_name: Optional[str] = None,
                 output_dir: Optional[str] = None,
                 auto_split: bool = True) -> None:
    """Запустить полный конвейер обработки.

    Args:
        bom_path: Путь к BOM-файлу.
        cards_path: Путь к папке/ZIP-архиву с операционными картами.
        config_name: Название комплектации (если не указана, будет интерактивный выбор).
        output_dir: Директория для результатов.
        auto_split: Автоматически разделять многолистовые карты.
    """
    start_time = time.time()

    if output_dir is None:
        output_dir = os.path.join(os.getcwd(), "output")
    os.makedirs(output_dir, exist_ok=True)

    print(f"\n{'=' * 60}")
    print(f"🚀 Burlak Parser — Система сверки BOM и операционных карт")
    print(f"{'=' * 60}")
    print(f"BOM файл: {bom_path}")
    print(f"Карты:    {cards_path}")
    print(f"Результат: {output_dir}")
    print()

    # Шаг 1: Загрузка BOM и выбор комплектации
    print("📋 Шаг 1: Загрузка BOM-файла...")
    with tqdm(total=1, desc="Парсинг BOM", unit="файл") as pbar:
        bom = parse_bom(bom_path)
        pbar.update(1)

    if config_name:
        if config_name not in bom.config_quantities:
            print(f"\n❌ Комплектация '{config_name}' не найдена!")
            print(f"Доступные варианты (первые 5):")
            for c in bom.config_names[:5]:
                print(f"  - {c}")
            sys.exit(1)
        selected_config = config_name
    else:
        selected_config = select_config_interactive(bom)

    bom_config_parts = get_config_quantities(bom, selected_config)
    print(f"\n✅ Выбрана комплектация: {selected_config[:60]}")
    print(f"   Деталей в комплектации: {len(bom_config_parts)}")
    print()

    # Шаг 2: Обработка операционных карт
    print("📂 Шаг 2: Обработка операционных карт...")
    cards_extract_dir = os.path.join(output_dir, "_extracted_cards")
    cards = parse_cards(cards_path, extract_dir=cards_extract_dir, show_progress=True)

    print(f"\n✅ Обработано карт: {cards.total_cards_processed}")
    print(f"   Всего листов: {cards.total_sheets_processed + cards.total_sheets_skipped}")
    print(f"   Из них непустых: {cards.total_sheets_processed}")
    print(f"   Пропущено (пустых): {cards.total_sheets_skipped}")
    print(f"   Уникальных деталей найдено: {len(cards.all_parts)}")
    print()

    # Шаг 2b: Разделение многолистовых файлов (опционально)
    split_dir = ""
    if auto_split:
        print("✂️  Разделение многолистовых карт на отдельные файлы...")
        split_dir = os.path.join(output_dir, "split_cards")
        created_files = split_cards_to_files(cards, split_dir)
        print(f"   Создано отдельных файлов: {len(created_files)}")

        # Создаём ZIP-архив
        print("📦 Создание ZIP-архива...")
        zip_path = os.path.join(output_dir, "split_cards.zip")
        create_split_cards_archive(split_dir, zip_path)
        print()

    # Шаг 3: Сверка
    print("🔍 Шаг 3: Сверка BOM и операционных карт...")
    comparison = compare(bom_config_parts, cards, config_name=selected_config)

    # Вывод сводки
    print(f"\n📊 Результаты сверки:")
    print(f"{'─' * 50}")
    print(f"  Деталей в BOM:          {comparison.total_bom_parts:>6}")
    print(f"  Деталей в картах:       {comparison.total_cards_parts:>6}")
    print(f"  Совпало:                {comparison.matched_parts:>6}")
    print(f"  Расхождений:            {len(comparison.discrepancies):>6}")
    print(f"    ├ Только в BOM:       {sum(1 for d in comparison.discrepancies if d.discrepancy_type == DiscrepancyType.ONLY_IN_BOM):>6}")
    print(f"    ├ Только в картах:    {sum(1 for d in comparison.discrepancies if d.discrepancy_type == DiscrepancyType.ONLY_IN_CARDS):>6}")
    print(f"    └ Конфликт количества:{sum(1 for d in comparison.discrepancies if d.discrepancy_type == DiscrepancyType.QUANTITY_MISMATCH):>6}")

    # Шаг 4: Формирование отчёта
    print("\n📄 Шаг 4: Формирование отчётов...")

    # Текстовый отчёт
    report_text = format_discrepancy_report(comparison)
    text_report_path = os.path.join(output_dir, "report.txt")
    with open(text_report_path, "w", encoding="utf-8") as f:
        f.write(report_text)
    print(f"   Текстовый отчёт: {text_report_path}")

    # Excel-отчёт
    excel_report_path = os.path.join(output_dir, "discrepancy_report.xlsx")
    generate_discrepancy_report(comparison, excel_report_path, bom_parts=bom_config_parts)
    print(f"   Excel-отчёт: {excel_report_path}")

    elapsed = time.time() - start_time
    print(f"\n{'=' * 60}")
    print(f"✅ Обработка завершена за {elapsed:.1f} сек.")
    print(f"   Результаты сохранены в: {output_dir}")
    print(f"{'=' * 60}")

    # Выводим первые несколько расхождений в консоль
    if comparison.discrepancies:
        print(f"\n📋 Первые расхождения ({min(10, len(comparison.discrepancies))} из {len(comparison.discrepancies)}):")
        print(f"{'─' * 80}")
        for disc in comparison.discrepancies[:10]:
            print(f"  {disc}")
        if len(comparison.discrepancies) > 10:
            print(f"  ... и ещё {len(comparison.discrepancies) - 10}")
    else:
        print("\n🎉 Расхождений не найдено!")


def main() -> None:
    """Точка входа CLI."""
    parser = argparse.ArgumentParser(
        description="Burlak Parser — Система сверки BOM и операционных карт",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Примеры:
  python -m burlak_parser.main --bom BOM.xlsx --cards ./cards/
  python -m burlak_parser.main --bom BOM.xlsx --cards ./cards.zip --config "T1L自由者..."
  python -m burlak_parser.main --bom BOM.xlsx --cards ./cards/ --output ./results/
        """,
    )

    parser.add_argument(
        "--bom", "-b",
        required=True,
        help="Путь к BOM-файлу (.xlsx)",
    )
    parser.add_argument(
        "--cards", "-c",
        required=True,
        help="Путь к папке с операционными картами или ZIP-архиву",
    )
    parser.add_argument(
        "--config", "-k",
        default=None,
        help="Название комплектации (если не указана, будет интерактивный выбор)",
    )
    parser.add_argument(
        "--output", "-o",
        default=None,
        help="Директория для результатов (по умолчанию: ./output)",
    )
    parser.add_argument(
        "--no-split",
        action="store_true",
        help="Не разделять многолистовые карты на отдельные файлы",
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Подробный вывод (debug)",
    )

    args = parser.parse_args()

    setup_logging(args.verbose)

    if not os.path.exists(args.bom):
        print(f"❌ BOM-файл не найден: {args.bom}")
        sys.exit(1)
    if not os.path.exists(args.cards):
        print(f"❌ Путь к картам не найден: {args.cards}")
        sys.exit(1)

    try:
        run_pipeline(
            bom_path=args.bom,
            cards_path=args.cards,
            config_name=args.config,
            output_dir=args.output,
            auto_split=not args.no_split,
        )
    except KeyboardInterrupt:
        print("\n\n⚠️  Прервано пользователем.")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Критическая ошибка: {e}")
        logger.exception("Pipeline завершился с ошибкой")
        print("\n💡 Подсказка: проверьте пути к файлам и их формат.")
        print("   Операционные карты: .xlsx или .xls")
        print("   BOM-файл: .xlsx")
        sys.exit(1)


if __name__ == "__main__":
    main()
