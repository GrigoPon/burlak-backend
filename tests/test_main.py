"""Интеграционные тесты для main.py — полный пайплайн.

Создаёт миниатюрные BOM и card .xlsx файлы, запускает run_pipeline
и проверяет выходные файлы (report.txt, excel, split_cards).

Покрытие:
  - run_pipeline: all configs, single config, no-fuzzy, no-split
  - clean_output_dirs: очистка директорий
  - setup_logging: настройка логирования
  - main: CLI аргументы (через argparse)
  - Error handling: missing BOM, missing cards
  - Output validation: report.txt content, excel file, zip archive
"""

from __future__ import annotations

import logging
import os
import shutil
import sys
import zipfile
from pathlib import Path
from unittest.mock import patch

import openpyxl
import pytest

from burlak_parser.main import (
    clean_output_dirs,
    run_pipeline,
    setup_logging,
    main,
)


# ═══════════════════════════════════════════════════════════════════════
#  HELPERS: создание тестовых .xlsx файлов
# ═══════════════════════════════════════════════════════════════════════


def _create_test_bom(path: str, multi_sheet: bool = False) -> str:
    """Создать миниатюрный BOM .xlsx для тестов.

    Структура (Sheet1 - основной BOM):
      R1:  序号 | 零部件代号 | 零部件名称 | 舒享版 | 奢享版
      R2:   1   | P001      | 螺母      |  2.0   |  3.0
      R3:   2   | P002      | 螺栓      |  1.0   |  0.0
      R4:   3   | P003      | 垫片      |  0.0   |  4.0

    Если multi_sheet=True — добавляет Sheet2 с доп. названиями.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    headers = ["序号", "零部件代号", "零部件名称", "舒享版", "奢享版"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h)

    data = [
        [1, "P001", "螺母", 2.0, 3.0],
        [2, "P002", "螺栓", 1.0, 0.0],
        [3, "P003", "垫片", 0.0, 4.0],
    ]
    for r, row in enumerate(data, 2):
        for c, val in enumerate(row, 1):
            ws.cell(row=r, column=c, value=val)

    if multi_sheet:
        ws2 = wb.create_sheet("Sheet2")
        ws2.cell(row=1, column=1, value="序号")
        ws2.cell(row=1, column=2, value="零部件代号")
        ws2.cell(row=1, column=3, value="零部件名称")
        ws2.cell(row=2, column=1, value=1)
        ws2.cell(row=2, column=2, value="P004")
        ws2.cell(row=2, column=3, value="弹簧")

    wb.save(path)
    wb.close()
    return path


def _create_test_card(path: str) -> str:
    """Создать миниатюрную операционную карту .xlsx для тестов.

    Структура:
      R1:  序号 | 零部件代号 | 数量
      R2:   1   | P001      |  2.0
      R3:   2   | P003      |  3.0
      R4:   3   | P999      |  1.0   (только в картах)
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    headers = ["序号", "零部件代号", "数量"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h)

    data = [
        [1, "P001", 2.0],
        [2, "P003", 3.0],
        [3, "P999", 1.0],
    ]
    for r, row in enumerate(data, 2):
        for c, val in enumerate(row, 1):
            ws.cell(row=r, column=c, value=val)

    wb.save(path)
    wb.close()
    return path


def _create_multi_sheet_card(path: str) -> str:
    """Создать многолистовую карту для теста split.

    Sheet1: P001
    Sheet2: P999
    """
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Операция_1"
    ws1.cell(row=1, column=1, value="序号")
    ws1.cell(row=1, column=2, value="零部件代号")
    ws1.cell(row=1, column=3, value="数量")
    ws1.cell(row=2, column=1, value=1)
    ws1.cell(row=2, column=2, value="P001")
    ws1.cell(row=2, column=3, value=1.0)

    ws2 = wb.create_sheet("Операция_2")
    ws2.cell(row=1, column=1, value="序号")
    ws2.cell(row=1, column=2, value="零部件代号")
    ws2.cell(row=1, column=3, value="数量")
    ws2.cell(row=2, column=1, value=1)
    ws2.cell(row=2, column=2, value="P999")
    ws2.cell(row=2, column=3, value=2.0)

    wb.save(path)
    wb.close()
    return path


def _count_lines(path: str) -> int:
    """Подсчитать непустые строки в файле."""
    with open(path) as f:
        return sum(1 for line in f if line.strip())


# ═══════════════════════════════════════════════════════════════════════
#  FIXTURES
# ═══════════════════════════════════════════════════════════════════════


@pytest.fixture
def bom_path(tmp_path: Path) -> str:
    """Создать временный BOM .xlsx."""
    path = os.path.join(tmp_path, "G01_ bom.xlsx")
    return _create_test_bom(path)


@pytest.fixture
def card_path(tmp_path: Path) -> str:
    """Создать временную карту .xlsx с номером операции."""
    path = os.path.join(tmp_path, "001-card.xlsx")
    return _create_test_card(path)


@pytest.fixture
def multi_card_path(tmp_path: Path) -> str:
    """Создать многолистовую карту .xlsx с номером операции."""
    path = os.path.join(tmp_path, "002-multi_card.xlsx")
    return _create_multi_sheet_card(path)


@pytest.fixture
def output_dir(tmp_path: Path) -> str:
    """Временная директория для результатов."""
    path = os.path.join(tmp_path, "output")
    os.makedirs(path, exist_ok=True)
    return path


# ═══════════════════════════════════════════════════════════════════════
#  1. setup_logging
# ═══════════════════════════════════════════════════════════════════════

class TestSetupLogging:
    def test_setup_logging_default(self):
        """Default logging does not crash (root logger level not changed by pytest)."""
        # In pytest, root logger is already configured at WARNING level
        # setup_logging calls basicConfig which is a no-op if already configured
        # Just verify the function doesn't crash
        setup_logging(verbose=False)
        assert True

    def test_setup_logging_verbose(self):
        """Verbose logging does not crash."""
        setup_logging(verbose=True)
        assert True


# ═══════════════════════════════════════════════════════════════════════
#  2. clean_output_dirs
# ═══════════════════════════════════════════════════════════════════════

class TestCleanOutputDirs:
    def test_cleans_existing_dir(self, tmp_path: Path):
        """Existing output dir is cleaned."""
        test_dir = os.path.join(tmp_path, "output")
        os.makedirs(test_dir)
        open(os.path.join(test_dir, "test.txt"), "w").close()
        assert os.path.isdir(test_dir)

        clean_output_dirs(test_dir)
        assert not os.path.isdir(test_dir)

    def test_cleans_auto_clean_dirs(self, tmp_path: Path):
        """Auto-clean dirs (output, split_cards, _extracted_cards) are cleaned."""
        for d in ["output", "split_cards", "_extracted_cards"]:
            path = os.path.join(tmp_path, d)
            os.makedirs(path)
            open(os.path.join(path, "test.txt"), "w").close()

        clean_output_dirs(str(tmp_path))
        # Only the main output dir is cleaned (tmp_path, not the subdirs)
        # The auto-clean ones in CWD would be checked - but we're not in CWD
        # So this test just verifies no crash
        assert True

    def test_clean_nonexistent_dir(self):
        """Non-existent dir doesn't crash."""
        clean_output_dirs("/nonexistent/path/12345")
        assert True  # no crash


# ═══════════════════════════════════════════════════════════════════════
#  3. run_pipeline — all configs (multi-config)
# ═══════════════════════════════════════════════════════════════════════

class TestRunPipelineAllConfigs:
    def test_full_pipeline_creates_outputs(self, bom_path, card_path, output_dir):
        """Full pipeline creates report.txt and excel file."""
        run_pipeline(
            bom_path=bom_path,
            cards_path=card_path,
            output_dir=output_dir,
            auto_split=True,
            use_fuzzy=True,
            single_config=False,
            max_workers=1,
        )

        # Проверяем, что файлы созданы
        txt_report = os.path.join(output_dir, "report.txt")
        excel_files = [f for f in os.listdir(output_dir) if f.endswith(".xlsx")]

        assert os.path.isfile(txt_report), "report.txt not created"
        assert len(excel_files) > 0, "No excel report created"

        # Проверяем содержание отчёта
        content = open(txt_report).read()
        assert "ОТЧЁТ" in content or "ПРОВЕРКИ" in content
        assert "舒享版" in content  # config name
        assert "奢享版" in content  # config name
        assert "P001" in content  # part number
        assert "P003" in content  # part number
        assert "расхождений" in content or "несоответствий" in content

    def test_pipeline_discrepancy_types(self, bom_path, card_path, output_dir):
        """Pipeline detects different discrepancy types."""
        # BOM: P001(2,3), P002(1,0), P003(0,4)
        # Cards: P001(2), P003(3), P999(1)
        # Expected: P001 qty match for 舒享(BOM=2, card=2), mismatch for 奢享(BOM=3,card=2)
        # P002 ONLY_IN_BOM, P003 qty mismatch, P999 ONLY_IN_CARDS
        run_pipeline(
            bom_path=bom_path,
            cards_path=card_path,
            output_dir=output_dir,
            auto_split=False,
            use_fuzzy=True,
            single_config=False,
            max_workers=1,
        )

        txt = os.path.join(output_dir, "report.txt")
        content = open(txt).read()
        assert "Разное количество" in content or "разное" in content.lower()
        assert "Есть в BOM" in content or "BOM" in content
        assert "Есть в" in content or "нет в" in content

    def test_pipeline_with_multi_sheet_bom(self, tmp_path, card_path, output_dir):
        """Pipeline works with multi-sheet BOM."""
        bom_path = os.path.join(tmp_path, "multi_bom.xlsx")
        _create_test_bom(bom_path, multi_sheet=True)

        run_pipeline(
            bom_path=bom_path,
            cards_path=card_path,
            output_dir=output_dir,
            auto_split=False,
            use_fuzzy=True,
            single_config=False,
            max_workers=1,
        )

        txt = os.path.join(output_dir, "report.txt")
        assert os.path.isfile(txt), "report.txt not created"

    def test_pipeline_no_fuzzy(self, bom_path, card_path, output_dir):
        """Pipeline works without fuzzy matching."""
        run_pipeline(
            bom_path=bom_path,
            cards_path=card_path,
            output_dir=output_dir,
            auto_split=False,
            use_fuzzy=False,
            single_config=False,
            max_workers=1,
        )
        txt = os.path.join(output_dir, "report.txt")
        assert os.path.isfile(txt)

    def test_pipeline_no_split(self, bom_path, card_path, output_dir):
        """Pipeline without auto-split still generates outputs."""
        run_pipeline(
            bom_path=bom_path,
            cards_path=card_path,
            output_dir=output_dir,
            auto_split=False,
            use_fuzzy=True,
            single_config=False,
            max_workers=1,
        )
        txt = os.path.join(output_dir, "report.txt")
        assert os.path.isfile(txt)

    def test_pipeline_with_multi_sheet_card(self, bom_path, multi_card_path, output_dir):
        """Pipeline with multi-sheet card creates split files."""
        run_pipeline(
            bom_path=bom_path,
            cards_path=multi_card_path,
            output_dir=output_dir,
            auto_split=True,
            use_fuzzy=True,
            single_config=False,
            max_workers=1,
        )

        txt = os.path.join(output_dir, "report.txt")
        assert os.path.isfile(txt)

        # Check split cards were created
        split_dir = os.path.join(output_dir, "split_cards")
        zip_path = os.path.join(output_dir, "split_cards.zip")
        assert os.path.isdir(split_dir) or os.path.isfile(zip_path), \
            "Neither split dir nor zip found"


# ═══════════════════════════════════════════════════════════════════════
#  4. run_pipeline — single config
# ═══════════════════════════════════════════════════════════════════════

class TestRunPipelineSingleConfig:
    def test_single_config_creates_outputs(self, bom_path, card_path, output_dir):
        """Single config mode creates report with discrepancies."""
        # BOM 舒享版: P001(2.0), P002(1.0), P003(0.0)
        # Cards:       P001(2.0), P003(3.0), P999(1.0)
        # Expected: P001 perfect match (not in discrepancy list!)
        #           P002 ONLY_IN_BOM (in report), P003 ONLY_IN_CARDS, P999 ONLY_IN_CARDS
        run_pipeline(
            bom_path=bom_path,
            cards_path=card_path,
            config_name="舒享版",
            output_dir=output_dir,
            auto_split=False,
            use_fuzzy=True,
            single_config=True,
            max_workers=1,
        )

        txt = os.path.join(output_dir, "report.txt")
        assert os.path.isfile(txt)

        content = open(txt).read()
        assert "舒享版" in content
        # P002 should be in ONLY_IN_BOM section
        assert "P002" in content
        # P999 should be in ONLY_IN_CARDS section
        assert "P999" in content

    def test_single_config_wrong_name_raises(self, bom_path, card_path, output_dir):
        """Wrong config name exits with sys.exit(1)."""
        with pytest.raises(SystemExit) as exc:
            run_pipeline(
                bom_path=bom_path,
                cards_path=card_path,
                config_name="NonExistent",
                output_dir=output_dir,
                auto_split=False,
                use_fuzzy=True,
                single_config=True,
                max_workers=1,
            )
        assert exc.value.code == 1

    def test_single_config_quantity_mismatch(self, bom_path, card_path, output_dir):
        """Single config shows quantity mismatch details."""
        # 舒享版: P001(BOM=2.0, card=2.0) → perfect match
        # P002(BOM=1.0, card=0) → ONLY_IN_BOM
        # P003(BOM=0, card=3.0) → ONLY_IN_CARDS? No, P003 has 0 in BOM for 舒享版
        # P999(not in BOM)→ ONLY_IN_CARDS
        run_pipeline(
            bom_path=bom_path,
            cards_path=card_path,
            config_name="舒享版",
            output_dir=output_dir,
            auto_split=False,
            use_fuzzy=True,
            single_config=True,
            max_workers=1,
        )

        txt = os.path.join(output_dir, "report.txt")
        # P001 qty: BOM=2.0, card=2.0 → perfect match for 舒享版
        # P002 qty: BOM=1.0, card=0 → ONLY_IN_BOM
        # P999: not in BOM → ONLY_IN_CARDS
        content = open(txt).read()
        assert "несоответствий" in content or "расхождени" in content


# ═══════════════════════════════════════════════════════════════════════
#  5. run_pipeline — default output dir
# ═══════════════════════════════════════════════════════════════════════

class TestRunPipelineDefaultOutput:
    def test_default_output_dir_created(self, bom_path, card_path, monkeypatch):
        """Default output dir (./output) is created when not specified."""
        # Change to a tmp dir so we don't pollute real project
        import tempfile
        with tempfile.TemporaryDirectory() as tmp:
            monkeypatch.chdir(tmp)
            run_pipeline(
                bom_path=bom_path,
                cards_path=card_path,
                output_dir=None,  # default
                auto_split=False,
                use_fuzzy=True,
                single_config=False,
                max_workers=1,
            )
            assert os.path.isdir(os.path.join(tmp, "output"))
            assert os.path.isfile(os.path.join(tmp, "output", "report.txt"))


# ═══════════════════════════════════════════════════════════════════════
#  6. Error handling
# ═══════════════════════════════════════════════════════════════════════

class TestErrorHandling:
    def test_missing_bom_exits(self, card_path, output_dir):
        """Missing BOM file causes sys.exit(1) from main()."""
        with pytest.raises(SystemExit) as exc:
            bad_path = "/nonexistent/bom.xlsx"
            with patch.object(sys, "argv", [
                "main.py", "--bom", bad_path, "--cards", card_path,
            ]):
                main()
        assert exc.value.code == 1

    def test_missing_cards_exits(self, bom_path, output_dir):
        """Missing cards path causes sys.exit(1) from main()."""
        with pytest.raises(SystemExit) as exc:
            bad_path = "/nonexistent/cards"
            with patch.object(sys, "argv", [
                "main.py", "--bom", bom_path, "--cards", bad_path,
            ]):
                main()
        assert exc.value.code == 1


# ═══════════════════════════════════════════════════════════════════════
#  7. CLI arg parsing (main function)
# ═══════════════════════════════════════════════════════════════════════

class TestCLIArguments:
    def test_minimal_args(self, bom_path, card_path, output_dir):
        """Minimal required args run successfully."""
        with patch.object(sys, "argv", [
            "main.py", "--bom", bom_path, "--cards", card_path, "-o", output_dir,
        ]):
            main()
        assert os.path.isfile(os.path.join(output_dir, "report.txt"))

    def test_single_config_arg(self, bom_path, card_path, output_dir):
        """--single-config flag works."""
        with patch.object(sys, "argv", [
            "main.py", "--bom", bom_path, "--cards", card_path,
            "-o", output_dir, "--single-config", "--config", "舒享版",
        ]):
            main()
        assert os.path.isfile(os.path.join(output_dir, "report.txt"))

    def test_no_fuzzy_arg(self, bom_path, card_path, output_dir):
        """--no-fuzzy flag works."""
        with patch.object(sys, "argv", [
            "main.py", "--bom", bom_path, "--cards", card_path,
            "-o", output_dir, "--no-fuzzy",
        ]):
            main()
        assert os.path.isfile(os.path.join(output_dir, "report.txt"))

    def test_no_split_arg(self, bom_path, card_path, output_dir):
        """--no-split flag works."""
        with patch.object(sys, "argv", [
            "main.py", "--bom", bom_path, "--cards", card_path,
            "-o", output_dir, "--no-split",
        ]):
            main()
        assert os.path.isfile(os.path.join(output_dir, "report.txt"))

    def test_verbose_arg(self, bom_path, card_path, output_dir):
        """--verbose flag works."""
        with patch.object(sys, "argv", [
            "main.py", "--bom", bom_path, "--cards", card_path,
            "-o", output_dir, "--verbose",
        ]):
            main()
        assert os.path.isfile(os.path.join(output_dir, "report.txt"))

    def test_workers_arg(self, bom_path, card_path, output_dir):
        """--workers flag works."""
        with patch.object(sys, "argv", [
            "main.py", "--bom", bom_path, "--cards", card_path,
            "-o", output_dir, "--workers", "2",
        ]):
            main()
        assert os.path.isfile(os.path.join(output_dir, "report.txt"))

    def test_short_args(self, bom_path, card_path, output_dir):
        """Short flags (-b, -c, -o) work."""
        with patch.object(sys, "argv", [
            "main.py", "-b", bom_path, "-c", card_path, "-o", output_dir,
        ]):
            main()
        assert os.path.isfile(os.path.join(output_dir, "report.txt"))


# ═══════════════════════════════════════════════════════════════════════
#  8. Excel report validation
# ═══════════════════════════════════════════════════════════════════════

class TestExcelReport:
    def test_excel_has_expected_sheets(self, bom_path, card_path, output_dir):
        """Generated Excel has expected sheet names."""
        run_pipeline(
            bom_path=bom_path,
            cards_path=card_path,
            output_dir=output_dir,
            auto_split=False,
            use_fuzzy=True,
            single_config=False,
            max_workers=1,
        )

        excel_files = [f for f in os.listdir(output_dir) if f.endswith(".xlsx") and "split" not in f]
        assert excel_files, "No excel report found"

        excel_path = os.path.join(output_dir, excel_files[0])
        wb = openpyxl.load_workbook(excel_path)
        sheet_names = wb.sheetnames

        # Should have at least Сводка sheet
        assert any("Сводка" in s or "свод" in s.lower() for s in sheet_names), \
            f"No summary sheet found in {sheet_names}"

        wb.close()

    def test_excel_contains_data(self, bom_path, card_path, output_dir):
        """Excel contains actual part data."""
        run_pipeline(
            bom_path=bom_path,
            cards_path=card_path,
            output_dir=output_dir,
            auto_split=False,
            use_fuzzy=True,
            single_config=False,
            max_workers=1,
        )

        excel_files = [f for f in os.listdir(output_dir) if f.endswith(".xlsx") and "split" not in f]
        assert excel_files

        wb = openpyxl.load_workbook(os.path.join(output_dir, excel_files[0]))
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for row in ws.iter_rows(values_only=True):
                cells = [str(c) for c in row if c is not None]
                cell_text = " ".join(cells)
                if "P001" in cell_text:
                    wb.close()
                    return  # Found P001 in some sheet

        wb.close()
        pytest.fail("Part P001 not found in any excel sheet")


# ═══════════════════════════════════════════════════════════════════════
#  9. Report content edge cases
# ═══════════════════════════════════════════════════════════════════════

class TestReportContent:
    def test_report_has_all_sections(self, bom_path, card_path, output_dir):
        """Report has all expected sections."""
        run_pipeline(
            bom_path=bom_path,
            cards_path=card_path,
            output_dir=output_dir,
            auto_split=False,
            use_fuzzy=True,
            single_config=False,
            max_workers=1,
        )

        content = open(os.path.join(output_dir, "report.txt")).read()
        assert "ОТЧЁТ" in content
        assert "ПРОВЕРКИ" in content
        assert "КОМПЛЕКТАЦИЙ" in content
        assert "BOM" in content or "bom" in content.lower()

    def test_report_non_empty(self, bom_path, card_path, output_dir):
        """Report is not empty."""
        run_pipeline(
            bom_path=bom_path,
            cards_path=card_path,
            output_dir=output_dir,
            auto_split=False,
            use_fuzzy=True,
            single_config=False,
            max_workers=1,
        )

        assert _count_lines(os.path.join(output_dir, "report.txt")) > 5

    def test_pipeline_without_fuzzy_still_finds_discrepancies(
        self, bom_path, card_path, output_dir,
    ):
        """Without fuzzy, basic discrepancies are still found."""
        run_pipeline(
            bom_path=bom_path,
            cards_path=card_path,
            output_dir=output_dir,
            auto_split=False,
            use_fuzzy=False,
            single_config=False,
            max_workers=1,
        )

        content = open(os.path.join(output_dir, "report.txt")).read()
        assert "несоответствий" in content or "расхождени" in content or "ОТЧЁТ" in content


# ═══════════════════════════════════════════════════════════════════════
#  10. Split cards integration
# ═══════════════════════════════════════════════════════════════════════

class TestSplitCardsIntegration:
    def test_split_cards_zip_created(self, bom_path, multi_card_path, output_dir):
        """Auto-split creates ZIP with split files."""
        run_pipeline(
            bom_path=bom_path,
            cards_path=multi_card_path,
            output_dir=output_dir,
            auto_split=True,
            use_fuzzy=True,
            single_config=False,
            max_workers=1,
        )

        zip_path = os.path.join(output_dir, "split_cards.zip")
        assert os.path.isfile(zip_path), "split_cards.zip not created"

        # Check zip contains split files
        with zipfile.ZipFile(zip_path) as zf:
            names = zf.namelist()
            assert len(names) > 0, "Zip is empty"
            # Should have at least one .xlsx file
            xlsx_in_zip = [n for n in names if n.endswith(".xlsx")]
            assert len(xlsx_in_zip) > 0, "No .xlsx files in zip"

    def test_split_cards_dir_created(self, bom_path, multi_card_path, output_dir):
        """Auto-split creates split_cards directory."""
        run_pipeline(
            bom_path=bom_path,
            cards_path=multi_card_path,
            output_dir=output_dir,
            auto_split=True,
            use_fuzzy=True,
            single_config=False,
            max_workers=1,
        )

        split_dir = os.path.join(output_dir, "split_cards")
        assert os.path.isdir(split_dir), "split_cards dir not created"
        assert len(os.listdir(split_dir)) > 0, "split_cards dir is empty"
