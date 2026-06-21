"""Автоматизированные тесты целостности пайплайна (Pipeline Integrity).

Эти тесты проверяют, что:
  1. Количество разделённых файлов для каждого бренда строго постоянно.
  2. Манифест split_manifest.json корректен и детерминирован.
  3. Отчёты discrepancies.xlsx и report.txt генерируются без ошибок.

Если кто-либо изменит парсер или сплиттер, эти тесты сразу покажут,
что нарушена детерминированность или целостность вывода.

Использование:
  python -m pytest tests/test_pipeline_integrity.py -v
"""

import json
import os
import sys
from pathlib import Path

import pytest

# ── Expected file counts (deterministic, measured from verified runs) ──
EXPECTED_SPLIT_COUNTS = {
    "jetour": 823,
    "swm": 320,
    "baic": 120,
    "changan": 628,
}

PROJECT_ROOT = Path(__file__).resolve().parent.parent


def _check_brand(brand: str, output_dir: str) -> dict:
    """Проверить целостность вывода для одного бренда.

    Args:
        brand: Имя бренда (для сообщений об ошибках).
        output_dir: Путь к директории с результатами.

    Returns:
        Словарь с результатами проверки.
    """
    result = {
        "brand": brand,
        "ok": True,
        "errors": [],
        "xlsx_count": 0,
        "manifest_count": 0,
        "has_discrepancies": False,
        "has_report": False,
        "has_manifest": False,
    }

    split_dir = os.path.join(output_dir, "split_cards")
    discrepancies_path = os.path.join(output_dir, "discrepancies.xlsx")
    report_path = os.path.join(output_dir, "report.txt")
    manifest_path = os.path.join(split_dir, "split_manifest.json")

    # Проверка split_cards
    if os.path.isdir(split_dir):
        xlsx_files = [
            f for f in os.listdir(split_dir)
            if f.endswith(".xlsx") and not f.startswith("~$")
        ]
        result["xlsx_count"] = len(xlsx_files)
    else:
        result["errors"].append(f"split_cards/ not found in {output_dir}")
        result["ok"] = False

    # Проверка манифеста
    if os.path.isfile(manifest_path):
        result["has_manifest"] = True
        try:
            with open(manifest_path, "r", encoding="utf-8") as f:
                manifest = json.load(f)
            total_manifested = sum(len(v) for v in manifest.values())
            result["manifest_count"] = total_manifested
            # Манифест должен содержать JSON-объект (dict), не список
            if not isinstance(manifest, dict):
                result["errors"].append(
                    f"Manifest is not a dict (got {type(manifest).__name__})"
                )
                result["ok"] = False
            # Все значения должны быть списками строк
            for src, files in manifest.items():
                if not isinstance(src, str):
                    result["errors"].append(
                        f"Manifest key is not a string: {type(src).__name__}"
                    )
                    result["ok"] = False
                if not isinstance(files, list):
                    result["errors"].append(
                        f"Manifest value for '{src}' is not a list"
                    )
                    result["ok"] = False
                for f in files:
                    if not isinstance(f, str):
                        result["errors"].append(
                            f"Manifest file entry is not a string: {type(f).__name__}"
                        )
                        result["ok"] = False
        except json.JSONDecodeError as e:
            result["errors"].append(f"Manifest JSON parse error: {e}")
            result["ok"] = False
    else:
        result["errors"].append(f"split_manifest.json not found")
        result["ok"] = False

    # Проверка discrepancies.xlsx
    if os.path.isfile(discrepancies_path):
        result["has_discrepancies"] = True
        # Проверяем, что это валидный xlsx (сигнатура PK)
        try:
            with open(discrepancies_path, "rb") as f:
                header = f.read(2)
                if header != b"PK":
                    result["errors"].append(
                        "discrepancies.xlsx is not a valid ZIP/xlsx file"
                    )
                    result["ok"] = False
        except Exception as e:
            result["errors"].append(f"Cannot read discrepancies.xlsx: {e}")
            result["ok"] = False
    else:
        result["errors"].append("discrepancies.xlsx not found")
        result["ok"] = False

    # Проверка report.txt
    if os.path.isfile(report_path):
        result["has_report"] = True
        try:
            with open(report_path, "r", encoding="utf-8") as f:
                content = f.read()
            # Отчёт должен содержать заголовок
            if "ОТЧЁТ" not in content or "ПРОВЕРКИ" not in content:
                result["errors"].append(
                    "report.txt missing report header"
                )
                result["ok"] = False
        except Exception as e:
            result["errors"].append(f"Cannot read report.txt: {e}")
            result["ok"] = False
    else:
        result["errors"].append("report.txt not found")
        result["ok"] = False

    # Сравнение xlsx_count с manifest_count (должны совпадать)
    if result["has_manifest"]:
        if result["xlsx_count"] != result["manifest_count"]:
            result["errors"].append(
                f"xlsx_count ({result['xlsx_count']}) != manifest_count "
                f"({result['manifest_count']})"
            )
            result["ok"] = False

    return result


def _run_determinism_check(brand: str, output_dir: str) -> dict:
    """Запустить проверку детерминизма: прочитать манифест и убедиться,
    что список файлов отсортирован и стабилен.

    Args:
        brand: Имя бренда.
        output_dir: Путь к директории.

    Returns:
        Словарь с результатами проверки.
    """
    result = {"brand": brand, "ok": True, "errors": []}

    manifest_path = os.path.join(output_dir, "split_cards", "split_manifest.json")
    if not os.path.isfile(manifest_path):
        result["ok"] = False
        result["errors"].append("No manifest for determinism check")
        return result

    with open(manifest_path, "r", encoding="utf-8") as f:
        manifest = json.load(f)

    # Проверка: манифест должен быть отсортирован по ключам
    keys = list(manifest.keys())
    if keys != sorted(keys):
        result["errors"].append("Manifest keys are not sorted")
        result["ok"] = False

    # Проверка: значения (списки файлов) должны быть отсортированы
    for src, files in manifest.items():
        if files != sorted(files):
            result["errors"].append(f"Manifest files for '{src}' are not sorted")
            result["ok"] = False

    return result


# ══════════════════════════════════════════════════════════════════════════
# TESTS
# ══════════════════════════════════════════════════════════════════════════


def test_jetour_split_count():
    """Jetour T1L: ровно 823 split .xlsx файлов (детерминировано)."""
    output_dir = os.path.join(PROJECT_ROOT, "output_jetour")
    if not os.path.isdir(output_dir):
        pytest.skip(f"Output dir not found: {output_dir}")
    check = _check_brand("jetour", output_dir)
    assert check["ok"], "; ".join(check["errors"])
    assert check["xlsx_count"] == EXPECTED_SPLIT_COUNTS["jetour"], (
        f"Jetour: expected {EXPECTED_SPLIT_COUNTS['jetour']} split files, "
        f"got {check['xlsx_count']}"
    )


def test_swm_split_count():
    """SWM (G01): ровно 227 split .xlsx файлов."""
    output_dir = os.path.join(PROJECT_ROOT, "output_swm")
    if not os.path.isdir(output_dir):
        pytest.skip(f"Output dir not found: {output_dir}")
    check = _check_brand("swm", output_dir)
    assert check["ok"], "; ".join(check["errors"])
    assert check["xlsx_count"] == EXPECTED_SPLIT_COUNTS["swm"], (
        f"SWM: expected {EXPECTED_SPLIT_COUNTS['swm']} split files, "
        f"got {check['xlsx_count']}"
    )


def test_baic_split_count():
    """BAIC: ровно 120 split .xlsx файлов."""
    output_dir = os.path.join(PROJECT_ROOT, "output_baic")
    if not os.path.isdir(output_dir):
        pytest.skip(f"Output dir not found: {output_dir}")
    check = _check_brand("baic", output_dir)
    assert check["ok"], "; ".join(check["errors"])
    assert check["xlsx_count"] == EXPECTED_SPLIT_COUNTS["baic"], (
        f"BAIC: expected {EXPECTED_SPLIT_COUNTS['baic']} split files, "
        f"got {check['xlsx_count']}"
    )


def test_changan_split_count():
    """Changan: ровно 628 split .xlsx файлов."""
    output_dir = os.path.join(PROJECT_ROOT, "output_changan")
    if not os.path.isdir(output_dir):
        pytest.skip(f"Output dir not found: {output_dir}")
    check = _check_brand("changan", output_dir)
    assert check["ok"], "; ".join(check["errors"])
    assert check["xlsx_count"] == EXPECTED_SPLIT_COUNTS["changan"], (
        f"Changan: expected {EXPECTED_SPLIT_COUNTS['changan']} split files, "
        f"got {check['xlsx_count']}"
    )


def test_all_brands_manifest_integrity():
    """Проверить манифесты всех 4 брендов на корректность."""
    brands = {
        "jetour": "output_jetour",
        "swm": "output_swm",
        "baic": "output_baic",
        "changan": "output_changan",
    }
    for brand, out_dir in brands.items():
        output_dir = os.path.join(PROJECT_ROOT, out_dir)
        if not os.path.isdir(output_dir):
            continue  # Skip if not generated yet
        check = _check_brand(brand, output_dir)
        assert check["has_manifest"], f"{brand}: missing split_manifest.json"
        assert check["has_discrepancies"], (
            f"{brand}: missing discrepancies.xlsx"
        )
        assert check["has_report"], f"{brand}: missing report.txt"
        # xlsx_count must match manifest_count
        assert check["xlsx_count"] == check["manifest_count"], (
            f"{brand}: xlsx_count ({check['xlsx_count']}) != "
            f"manifest_count ({check['manifest_count']})"
        )


def test_all_brands_determinism():
    """Проверить, что манифесты всех брендов отсортированы детерминированно."""
    brands = {
        "jetour": "output_jetour",
        "swm": "output_swm",
        "baic": "output_baic",
        "changan": "output_changan",
    }
    for brand, out_dir in brands.items():
        output_dir = os.path.join(PROJECT_ROOT, out_dir)
        if not os.path.isdir(output_dir):
            continue
        det = _run_determinism_check(brand, output_dir)
        assert det["ok"], (
            f"{brand}: determinism check failed: "
            f"{'; '.join(det['errors'])}"
        )


def test_split_manifest_json_structure():
    """Проверить JSON-структуру манифеста Jetour (самого большого набора)."""
    manifest_path = os.path.join(
        PROJECT_ROOT, "output_jetour", "split_cards", "split_manifest.json"
    )
    if not os.path.isfile(manifest_path):
        return  # Skip if not generated

    with open(manifest_path, "r", encoding="utf-8") as f:
        manifest = json.load(f)

    # Тип — dict
    assert isinstance(manifest, dict), "Manifest must be a dict"

    # Должно быть > 0 источников
    assert len(manifest) > 0, "Manifest must have at least 1 source file"

    # Каждый источник должен иметь список файлов
    for src, files in manifest.items():
        assert isinstance(src, str), f"Source key must be string: {src}"
        assert isinstance(files, list), (
            f"Files for {src} must be list"
        )
        assert len(files) > 0, f"Source {src} must have at least 1 file"
        for f in files:
            assert isinstance(f, str), (
                f"File entry must be string: {f}"
            )
            assert f.endswith(".xlsx"), (
                f"File entry must end with .xlsx: {f}"
            )


def test_report_txt_contains_results():
    """Проверить, что report.txt содержит результаты сверки для Jetour."""
    report_path = os.path.join(PROJECT_ROOT, "output_jetour", "report.txt")
    if not os.path.isfile(report_path):
        return  # Skip if not generated

    with open(report_path, "r", encoding="utf-8") as f:
        content = f.read()

    # Должен содержать количество комплектаций
    assert "Всего проверено комплектаций: 78" in content, (
        "report.txt должно содержать информацию о 78 конфигурациях"
    )
    # Должен содержать информацию о несоответствиях (discrepancies)
    assert "Найдено несоответствий" in content, (
        "report.txt должно содержать информацию о несоответствиях"
    )
    # Должен содержать заголовок отчёта
    assert "ОТЧЁТ ПРОВЕРКИ КОМПЛЕКТАЦИЙ" in content, (
        "report.txt должно содержать заголовок отчёта"
    )
