#!/usr/bin/env python3
"""Точка входа в систему автоматического разбора и сверки ведомостей материалов (BOM).

Использование:
  python -m burlak_parser.main --bom <file.xlsx> --cards <path> [--config <name>]

Новый режим (по умолчанию): обработка ВСЕХ комплектаций одновременно.
Старый режим (--single-config): обработка одной выбранной комплектации.
"""

from __future__ import annotations

import argparse
import logging
import os
import shutil
import sys
import time
from pathlib import Path
from typing import List, Optional

from tqdm import tqdm

from burlak_parser.bom_parser import (
    BOMData,
    BOMService,
    PartInfo,
    get_all_config_quantities,
    get_config_quantities,
    parse_bom,
)
from burlak_parser.card_parser import (
    CardService,
    CardsData,
    parse_cards,
    split_cards_to_files,
)
from burlak_parser.comparator import (
    Discrepancy,
    DiscrepancyType,
    IntegrityCheck,
    MatchingEngine,
    MultiConfigComparisonResult,
    compare_all_configs,
    compare_single_config,
    verify_integrity,
)
from burlak_parser.fuzzy_matcher import FuzzyMatcher
from burlak_parser.report_generator import (
    Reporter,
    create_split_cards_archive,
)
from burlak_parser.diagnostic import (
    DiagnosticDumper,
    create_diagnostic_from_bom,
    create_diagnostic_from_cards,
)

logger = logging.getLogger(__name__)

# Директории, которые автоматически очищаются при запуске
AUTO_CLEAN_DIRS = ["output", "split_cards", "_extracted_cards"]


def setup_logging(verbose: bool = False, quiet: bool = False) -> None:
    """Настроить логирование."""
    if quiet:
        level = logging.WARNING
    elif verbose:
        level = logging.DEBUG
    else:
        level = logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )


def clean_output_dirs(output_dir: str) -> None:
    """Автоматически очистить временные директории от предыдущих запусков.

    Очищает:
      - Основную директорию результатов (output_dir).
      - Поддиректории split_cards, _extracted_cards внутри output_dir.
      - Дополнительные директории из AUTO_CLEAN_DIRS в текущей папке.

    Args:
        output_dir: Путь к основной директории результатов.
    """
    dirs_to_clean: List[str] = []

    # Основная директория результатов
    if os.path.isdir(output_dir):
        dirs_to_clean.append(output_dir)

    # Дополнительные auto-clean директории в CWD
    cwd = os.getcwd()
    for dirname in AUTO_CLEAN_DIRS:
        path = os.path.join(cwd, dirname)
        if os.path.isdir(path) and path != output_dir:
            dirs_to_clean.append(path)

    for d in dirs_to_clean:
        try:
            logger.info("Очистка директории: %s", d)
            shutil.rmtree(d)
        except Exception as e:
            logger.warning("Не удалось очистить %s: %s", d, e)


def select_config_interactive(bom: BOMData) -> str:
    """Интерактивный выбор комплектации из списка."""
    configs = bom.config_names
    if not configs:
        print("\u274c Нет доступных комплектаций в BOM-файле.")
        sys.exit(1)

    if len(configs) == 1:
        print(f"\u2705 Автоматически выбрана единственная комплектация: {configs[0][:60]}")
        return configs[0]

    print(f"\n{'=' * 60}")
    print(f"Доступные комплектации ({len(configs)} шт.):")
    print(f"{'=' * 60}")

    display_configs = configs[:30]
    for i, name in enumerate(display_configs, 1):
        display_name = name if len(name) <= 70 else name[:67] + "..."
        print(f"  {i:3d}. {display_name}")

    while True:
        try:
            choice = input(f"\nВыберите комплектацию (1-{len(display_configs)}): ").strip()
            idx = int(choice) - 1
            if 0 <= idx < len(display_configs):
                return configs[idx]
            else:
                print(f"\u274c Введите число от 1 до {len(display_configs)}")
        except ValueError:
            print("\u274c Введите корректное число")


def run_pipeline(
    bom_path: str,
    cards_path: str,
    config_name: Optional[str] = None,
    output_dir: Optional[str] = None,
    auto_split: bool = True,
    use_fuzzy: bool = True,
    single_config: bool = False,
    max_workers: Optional[int] = None,
    show_split_stats: bool = False,
    diagnostic: bool = False,
) -> None:
    """Запустить полный конвейер обработки.

    Args:
        bom_path: Путь к BOM-файлу (.xlsx).
        cards_path: Путь к папке/ZIP-архиву с операционными картами.
        config_name: Название комплектации (для single_config режима).
        output_dir: Директория для результатов.
        auto_split: Автоматически разделять многолистовые карты.
        use_fuzzy: Использовать нечеткое сравнение.
        single_config: Только одна комплектация (старый режим).
        max_workers: Количество процессов для параллелизации.
        show_split_stats: Показывать детальную статистику split.
    """
    start_time = time.time()

    if output_dir is None:
        output_dir = os.path.join(os.getcwd(), "output")

    # Автоочистка
    clean_output_dirs(output_dir)
    os.makedirs(output_dir, exist_ok=True)

    print(f"\n{'=' * 60}")
    print(f"\U0001f680 Burlak Parser — Система сверки BOM и операционных карт")
    print(f"{'=' * 60}")
    print(f"BOM файл: {bom_path}")
    print(f"Карты:    {cards_path}")
    print(f"Результат: {output_dir}")
    print(f"Режим:    {'Одна комплектация' if single_config else 'ВСЕ комплектации'}")
    print(f"Fuzzy:    {'Вкл' if use_fuzzy else 'Выкл'}")
    print()

    # Шаг 1: Загрузка BOM
    print("\U0001f4cb Шаг 1: Загрузка BOM-файла...")
    with tqdm(total=1, desc="Парсинг BOM", unit="файл") as pbar:
        bom = parse_bom(bom_path)
        pbar.update(1)

    print(f"\n\u2705 BOM загружен: {len(bom.parts)} деталей, {len(bom.config_names)} комплектаций")

    # ── Диагностический дамп BOM ──
    if diagnostic:
        diag_dir = os.path.join(output_dir, "diagnostic")
        os.makedirs(diag_dir, exist_ok=True)
        bom_dump_path = create_diagnostic_from_bom(bom, diag_dir)
        print(f"   \U0001f4cb Diagnostic: BOM dump → {bom_dump_path}")
    if not single_config:
        print(f"   Будут обработаны ВСЕ {len(bom.config_names)} комплектаций одновременно.")
    else:
        if config_name:
            if config_name not in bom.config_quantities:
                print(f"\n\u274c Комплектация '{config_name}' не найдена!")
                print(f"Доступные варианты (первые 5):")
                for c in bom.config_names[:5]:
                    print(f"  - {c}")
                sys.exit(1)
            selected_config = config_name
        else:
            selected_config = select_config_interactive(bom)
        bom_config_parts = get_config_quantities(bom, selected_config)
        print(f"\n\u2705 Выбрана комплектация: {selected_config[:60]}")
        print(f"   Деталей в комплектации: {len(bom_config_parts)}")
    print()

    # Шаг 2: Обработка операционных карт
    print("\U0001f4c2 Шаг 2: Обработка операционных карт...")
    cards_extract_dir = os.path.join(output_dir, "_extracted_cards")
    cards = parse_cards(
        cards_path,
        extract_dir=cards_extract_dir,
        show_progress=True,
        max_workers=max_workers,
    )

    print(f"\n\u2705 Обработано карт: {cards.total_cards_processed}")
    print(f"   Служебных файлов пропущено: {cards.service_files_skipped}")
    if cards.corrupted_files:
        print(f"   \u26a0\ufe0f  Повреждённых файлов: {len(cards.corrupted_files)}")
    print(f"   Всего листов: {cards.total_sheets_processed + cards.total_sheets_skipped}")
    print(f"   Из них непустых: {cards.total_sheets_processed}")
    print(f"   Пропущено (пустых): {cards.total_sheets_skipped}")
    print(f"   Уникальных деталей найдено: {len(cards.all_parts)}")

    # ── Диагностический дамп карт ──
    if diagnostic:
        diag_dir = os.path.join(output_dir, "diagnostic")
        os.makedirs(diag_dir, exist_ok=True)
        oc_dump_path = create_diagnostic_from_cards(cards, diag_dir)
        print(f"   \U0001f4cb Diagnostic: OC dump → {oc_dump_path}")
    print()

    # Шаг 2b: Разделение многолистовых файлов
    split_dir = ""
    if auto_split:
        # Используем кэш для инкрементальной обработки
        from burlak_parser.cache import ProcessingCache
        cache_dir = os.path.join(output_dir, ".burlak_cache")
        cache = ProcessingCache(cache_dir)

        print("\u2702\ufe0f  Разделение многолистовых карт на отдельные файлы...")
        split_dir = os.path.join(output_dir, "split_cards")
        created_files = split_cards_to_files(
            cards, split_dir, max_workers=max_workers,
        )

        # Сохраняем кэш
        cache.save()
        logger.info("Кэш: %d записей", cache.size)

        # ── Статистика split ──
        split_stats = cards.split_stats

        print(f"   Создано отдельных файлов: {len(created_files)}")
        print(f"   .xlsx файлов: {split_stats.total_xlsx}")
        print(f"   .xls файлов (не разделяются): {split_stats.total_xls}")
        print(f"   Повреждённых при split: {split_stats.total_errors}")

        if split_stats.openpyxl_fallback_count > 0:
            print(f"   \u2705 Успешно спасены через openpyxl (fallback): {split_stats.openpyxl_fallback_count} файлов")
            for fname in split_stats.openpyxl_fallback_files:
                print(f"     - {fname}")

        if show_split_stats and split_stats is not None:
            # ── Детальная статистика split ──
            print(f"\n   📊 Детальная статистика split:")
            print(f"     - Всего файлов в картах: {len(cards.card_results)}")
            print(f"     - Служебных файлов (пропущено): {split_stats.total_service_files}")
            print(f"     - Всего листов: {split_stats.total_sheets_all}")
            print(f"       ├ Разделено: {split_stats.total_sheets_split}")
            print(f"       └ Пропущено: {split_stats.total_sheets_skipped}")

            # Причины пропуска листов
            if split_stats.total_sheets_skipped > 0:
                print(f"\n   🔍 Причины пропуска листов:")
                for reason, count in split_stats.get_top_skip_reasons():
                    print(f"     - {reason}: {count}")

            # Топ файлов по пропускам
            top_skips = split_stats.get_files_with_most_skips(5)
            if top_skips:
                print(f"\n   📁 Файлы с наибольшим числом пропущенных листов:")
                for fname, total, skipped in top_skips:
                    print(f"     - {fname[:55]:55s} всего={total} пропущено={skipped}")

            # Файлы с ошибками
            error_files = [fs for fs in split_stats.file_stats if fs.has_error]
            if error_files:
                print(f"\n   ❌ Файлы с ошибками ({len(error_files)}):")
                for fs in error_files[:5]:
                    print(f"     - {fs.file_name[:55]}: {fs.error_message[:80]}")
                if len(error_files) > 5:
                    print(f"     ... и ещё {len(error_files) - 5}")

        print(f"\n\U0001f4e6 Создание ZIP-архива...")
        zip_path = os.path.join(output_dir, "split_cards.zip")
        create_split_cards_archive(split_dir, zip_path)
        print()

    # Шаг 3: Сверка
    print("\U0001f50d Шаг 3: Сверка BOM и операционных карт...")

    if single_config:
        # Старый режим: одна комплектация
        all_bom_parts = set(bom.parts.keys())
        fuzzy_matcher = FuzzyMatcher(all_bom_parts) if use_fuzzy else None
        single_result = compare_single_config(
            bom_config_parts, cards, config_name=selected_config,
            fuzzy_matcher=fuzzy_matcher,
        )
        result = MultiConfigComparisonResult(
            config_results=[single_result],
            all_discrepancies=list(single_result.discrepancies),
            total_configs=1,
            total_bom_unique_parts=len(all_bom_parts),
            total_cards_unique_parts=len(cards.all_parts),
        )

        print(f"\n\U0001f4ca Результаты сверки:")
        print(f"{'\u2500' * 50}")
        print(f"  Деталей в BOM:          {single_result.total_bom_parts:>6}")
        print(f"  Деталей в картах:       {single_result.total_cards_parts:>6}")
        print(f"  Совпало:                {single_result.matched_parts:>6}")
        print(f"  Fuzzy match:            {single_result.fuzzy_matched:>6}")
        print(f"  Расхождений:            {len(single_result.discrepancies):>6}")
        print(f"    \u251c Только в BOM:       {sum(1 for d in single_result.discrepancies if d.discrepancy_type == DiscrepancyType.ONLY_IN_BOM):>6}")
        print(f"    \u251c Только в картах:    {sum(1 for d in single_result.discrepancies if d.discrepancy_type == DiscrepancyType.ONLY_IN_CARDS):>6}")
        print(f"    \u2514 Конфликт количества: {sum(1 for d in single_result.discrepancies if d.discrepancy_type == DiscrepancyType.QUANTITY_MISMATCH):>6}")
    else:
        # Новый режим: все комплектации
        result = compare_all_configs(bom, cards, use_fuzzy=use_fuzzy)

        print(f"\n\U0001f4ca Результаты сверки (ВСЕ {result.total_configs} комплектаций):")
        print(f"{'\u2500' * 50}")
        print(f"  Уникальных деталей BOM:  {result.total_bom_unique_parts:>6}")
        print(f"  Уникальных деталей карт: {result.total_cards_unique_parts:>6}")
        print(f"  Всего расхождений:       {len(result.all_discrepancies):>6}")

        for cr in result.config_results[:5]:
            short = cr.config_name[:45]
            print(f"  {short:45s}  BOM={cr.total_bom_parts:>4}  Карты={cr.total_cards_parts:>4}  Disc={len(cr.discrepancies):>4}")
        if result.total_configs > 5:
            print(f"  ... и ещё {result.total_configs - 5} комплектаций")

    # Шаг 4: Формирование отчётов
    print("\n\U0001f4c4 Шаг 4: Формирование отчётов...")

    reporter = Reporter()
    outputs = reporter.generate(result, output_dir, bom=bom, cards_data=cards)
    print(f"   Текстовый отчёт: {outputs.get('text_report', 'N/A')}")
    print(f"   Excel-отчёт: {outputs.get('excel_report', 'N/A')}")

    # ── Верификация целостности ──
    print(f"\n\U0001f50d Верификация целостности:")
    print(f"{'\u2500' * 60}")

    integrity = verify_integrity(result)

    if integrity.is_ok:
        print(f"  \u2705 {integrity.configs_ok}/{integrity.total_configs} конфигураций: matched + discrepancies = total_bom_parts")
        print(f"  \u2705 Сумма типов расхождений совпадает с общим количеством")
    else:
        for issue in integrity.config_issues:
            logger.warning("Нарушение целостности: %s", issue)
        if integrity.global_issue:
            logger.warning("Нарушение целостности: %s", integrity.global_issue)
        print(f"  \u26a0\ufe0f  Обнаружены нарушения целостности (см. лог)")
    print()

    elapsed = time.time() - start_time
    print(f"\n{'=' * 60}")
    print(f"\u2705 Обработка завершена за {elapsed:.1f} сек.")
    print(f"   Результаты сохранены в: {output_dir}")
    print(f"{'=' * 60}")

    # Выводим первые несколько расхождений в консоль
    if result.all_discrepancies:
        n_show = min(10, len(result.all_discrepancies))
        print(f"\n\U0001f4cb Первые расхождения ({n_show} из {len(result.all_discrepancies)}):")
        print(f"{'\u2500' * 80}")
        for disc in result.all_discrepancies[:10]:
            print(f"  {disc}")
        if len(result.all_discrepancies) > 10:
            print(f"  ... и ещё {len(result.all_discrepancies) - 10}")
    else:
        print("\n\U0001f389 Расхождений не найдено!")


def main() -> None:
    """Точка входа CLI."""
    parser = argparse.ArgumentParser(
        description="Burlak Parser — Система сверки BOM и операционных карт",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Примеры:
  # Обработка ВСЕХ комплектаций (новый режим по умолчанию)
  python -m burlak_parser.main --bom BOM.xlsx --cards ./cards/

  # Обработка одной комплектации (старый режим)
  python -m burlak_parser.main --bom BOM.xlsx --cards ./cards/ --single-config

  # С указанием конкретной комплектации
  python -m burlak_parser.main --bom BOM.xlsx --cards ./cards.zip --single-config --config "T1L..."

  # Без нечеткого сравнения
  python -m burlak_parser.main --bom BOM.xlsx --cards ./cards/ --no-fuzzy

  # С указанием количества процессов
  python -m burlak_parser.main --bom BOM.xlsx --cards ./cards/ --workers 8

  # Детальная статистика split (причины пропуска листов, топ файлов)
  python -m burlak_parser.main --bom BOM.xlsx --cards ./cards/ --split-stats

  # Без разделения многолистовых файлов
  python -m burlak_parser.main --bom BOM.xlsx --cards ./cards/ --no-split
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
        help="Название комплектации (только для --single-config режима)",
    )
    parser.add_argument(
        "--output", "-o",
        default=None,
        help="Директория для результатов (по умолчанию: ./output)",
    )
    parser.add_argument(
        "--single-config", "-s",
        action="store_true",
        help="Обработать только одну комплектацию (старый режим)",
    )
    parser.add_argument(
        "--no-split",
        action="store_true",
        help="Не разделять многолистовые карты на отдельные файлы",
    )
    parser.add_argument(
        "--no-fuzzy",
        action="store_true",
        help="Отключить нечеткое сравнение парт-номеров",
    )
    parser.add_argument(
        "--workers", "-w",
        type=int,
        default=None,
        help=f"Количество процессов (по умолчанию: количество CPU = {os.cpu_count() or 4})",
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Подробный вывод (debug)",
    )
    parser.add_argument(
        "--split-stats", "-S",
        action="store_true",
        help="Показать детальную статистику разделения файлов (топ причин, пропуски, ошибки)",
    )
    parser.add_argument(
        "--diagnostic", "-d",
        action="store_true",
        help="Режим диагностики: вывод промежуточных данных в JSON (BOM_dump, OC_dump, schema)",
    )
    parser.add_argument(
        "--quiet", "-q",
        action="store_true",
        help="Тихий режим (только ошибки и предупреждения)",
    )

    args = parser.parse_args()

    # ── Input validation ──
    if args.workers is not None and not (1 <= args.workers <= 32):
        print(f"\u274c --workers must be between 1 and 32, got {args.workers}")
        sys.exit(1)

    setup_logging(verbose=args.verbose, quiet=args.quiet)

    if not os.path.exists(args.bom):
        logger.error("BOM-файл не найден: %s", args.bom)
        sys.exit(1)
    if not os.path.exists(args.cards):
        logger.error("Путь к картам не найден: %s", args.cards)
        sys.exit(1)

    # File size validation
    bom_size_mb = os.path.getsize(args.bom) / (1024 * 1024)
    if bom_size_mb > 200:
        logger.warning("BOM file is very large (%.1f MB) — may use significant memory", bom_size_mb)
    logger.info("BOM file size: %.1f MB", bom_size_mb)

    try:
        run_pipeline(
            bom_path=args.bom,
            cards_path=args.cards,
            config_name=args.config,
            output_dir=args.output,
            auto_split=not args.no_split,
            use_fuzzy=not args.no_fuzzy,
            single_config=args.single_config,
            max_workers=args.workers,
            show_split_stats=args.split_stats,
            diagnostic=args.diagnostic,
        )
    except KeyboardInterrupt:
        logger.warning("Прервано пользователем.")
        sys.exit(1)
    except Exception as e:
        logger.error("Критическая ошибка: %s", e)
        logger.exception("Pipeline завершился с ошибкой")
        logger.info("Подсказка: проверьте пути к файлам и их формат.")
        logger.info("  Операционные карты: .xlsx или .xls")
        logger.info("  BOM-файл: .xlsx")
        sys.exit(1)


if __name__ == "__main__":
    main()
