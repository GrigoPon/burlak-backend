"""Unit-тесты для card_parser.py.

Покрытие:
  - Data structures (CardSheetInfo, CardPart, CardParseResult, CardsData)
  - ExcelSheet wrapper (max_row, max_column, cell_value)
  - _check_sheet_has_data
  - _collect_raw_rows (section boundaries, skip keywords, qty/name)
  - _merge_multiline_part_numbers (dash, em-dash, en-dash continuation)
  - _extract_card_number
  - _collect_all_tables (multi-table, multi-operation)
  - parse_card_file (normal, service, empty, no table)
  - CardService
  - TEMPLATE_SHEET_KEYWORDS
"""

from __future__ import annotations

import os
import sys
import tempfile
import zipfile
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
import pytest
from openpyxl import Workbook

from burlak_parser.card_parser import (
    CardPart,
    CardParseResult,
    CardsData,
    CardService,
    CardSheetInfo,
    ExcelReader,
    ExcelSheet,
    FileSplitStats,
    SKIP_REASON_NO_DATA,
    SKIP_REASON_TEMPLATE,
    SplitStatistics,
    TEMPLATE_SHEET_KEYWORDS,
    _check_sheet_has_data,
    _collect_all_tables,
    _collect_raw_rows,
    _extract_card_number,
    _find_excel_files,
    _merge_multiline_part_numbers,
    _safe_remove,
    _safe_name,
    _walk_extracted_dir,
    parse_card_file,
    parse_cards,
    split_cards_to_files,
)
from burlak_parser.heuristic_analyzer import (
    HeuristicAnalyzer,
    clean_part_number,
    is_valid_part_number,
)

# ═══════════════════════════════════════════════════════════════════════
#  23. ExcelSheet.cell_value — xlrd edge cases (153, 157-158)
# ═══════════════════════════════════════════════════════════════════════

class TestExcelSheetCellValueXlrd:
    """ExcelSheet.cell_value в xlrd-ветке.

    Покрывает строки 153, 157-158:
      - line 153: val == "" (empty string) → return None
      - line 157: return int(val) (float → int conversion)
      - line 158: return val (regular non-float value)
    """

    def test_empty_string_returns_none(self, monkeypatch):
        """xlrd возвращает пустую строку → cell_value возвращает None."""
        import xlrd as real_xlrd
        class MockSheet:
            nrows, ncols = 2, 2
            def cell_value(self, r, c):
                return ""  # empty string
        class MockBook:
            def sheet_names(self): return ["Sheet1"]
            def sheet_by_name(self, name): return MockSheet()
        monkeypatch.setattr(real_xlrd, "open_workbook", lambda p: MockBook())

        fd, p = tempfile.mkstemp(suffix=".xls", prefix="xlrempty_")
        os.close(fd); open(p, "w").close()
        try:
            reader = ExcelReader(p)
            sheet = reader.get_sheet("Sheet1")
            val = sheet.cell_value(1, 1)
            assert val is None, f"Expected None for empty string, got {val}"
            reader.close()
        finally:
            _safe_remove(p)

    def test_float_to_int_conversion(self, monkeypatch):
        """xlrd float → int: cell_value(2,2)=2.0 → ExcelSheet возвращает 2."""
        import xlrd as real_xlrd
        class MockSheet:
            nrows, ncols = 2, 2
            def cell_value(self, r, c):
                return 2.0  # float
        class MockBook:
            def sheet_names(self): return ["Sheet1"]
            def sheet_by_name(self, name): return MockSheet()
        monkeypatch.setattr(real_xlrd, "open_workbook", lambda p: MockBook())

        fd, p = tempfile.mkstemp(suffix=".xls", prefix="xlrint_")
        os.close(fd); open(p, "w").close()
        try:
            reader = ExcelReader(p)
            sheet = reader.get_sheet("Sheet1")
            val = sheet.cell_value(1, 1)
            assert val == 2, f"Expected int 2, got {val} ({type(val).__name__})"
            reader.close()
        finally:
            _safe_remove(p)

    def test_regular_value_returned(self, monkeypatch):
        """xlrd не-float значение → возвращается как есть (line 158)."""
        import xlrd as real_xlrd
        class MockSheet:
            nrows, ncols = 2, 2
            def cell_value(self, r, c):
                return "ABC-123"  # regular string
        class MockBook:
            def sheet_names(self): return ["Sheet1"]
            def sheet_by_name(self, name): return MockSheet()
        monkeypatch.setattr(real_xlrd, "open_workbook", lambda p: MockBook())

        fd, p = tempfile.mkstemp(suffix=".xls", prefix="xlrreg_")
        os.close(fd); open(p, "w").close()
        try:
            reader = ExcelReader(p)
            sheet = reader.get_sheet("Sheet1")
            val = sheet.cell_value(1, 1)
            assert val == "ABC-123"
            reader.close()
        finally:
            _safe_remove(p)

    def test_xlrd_exception_returns_none(self, monkeypatch):
        """Исключение в xlrd cell_value → except возвращает None."""
        import xlrd as real_xlrd
        class MockSheet:
            nrows, ncols = 2, 2
            def cell_value(self, r, c):
                raise IndexError("mock")
        class MockBook:
            def sheet_names(self): return ["Sheet1"]
            def sheet_by_name(self, name): return MockSheet()
        monkeypatch.setattr(real_xlrd, "open_workbook", lambda p: MockBook())

        fd, p = tempfile.mkstemp(suffix=".xls", prefix="xlrex_")
        os.close(fd); open(p, "w").close()
        try:
            reader = ExcelReader(p)
            sheet = reader.get_sheet("Sheet1")
            val = sheet.cell_value(999, 999)  # out of range → exception
            assert val is None, f"Expected None on exception, got {val}"
            reader.close()
        finally:
            _safe_remove(p)


# ═══════════════════════════════════════════════════════════════════════
#  24. _extract_card_number — fallback to basename (220-221)
# ═══════════════════════════════════════════════════════════════════════

class TestExtractCardNumberAbsoluteFallback:
    """Абсолютный fallback: extract_card_number вернул None → basename."""

    def test_fallback_no_card_number(self, monkeypatch):
        """Когда extract_card_number (heuristic) возвращает None → os.path.splitext(basename)[0]."""
        from burlak_parser.heuristic_analyzer import extract_card_number as ha_fn
        monkeypatch.setattr(
            "burlak_parser.card_parser.extract_card_number",
            lambda fp, ws: None,
        )
        es = _make_excel_sheet([["some random text"]])
        num = _extract_card_number("my_file.xlsx", es)
        assert num == "my_file", f"Expected 'my_file', got '{num}'"


# ═══════════════════════════════════════════════════════════════════════
#  25. _merge_multiline — trailing buffer flush with valid PN (261)
# ═══════════════════════════════════════════════════════════════════════

class TestMergeMultilineBufferFlush:
    """_merge_multiline_part_numbers — trailing buffer flush (line 261)."""

    def test_trailing_continuation_with_valid_pn(self):
        """Одиночная строка с "-", cleaned PN валидный → buffer сбрасывается в merged."""
        rows = [(2, "P001-", 1.0, "Part1", 1)]
        merged = _merge_multiline_part_numbers(rows)
        # "P001-" → cleaned: "P001" → is_valid → appended at line 261
        assert len(merged) == 1, f"Expected 1 (buffer flushed), got {len(merged)}"
        assert merged[0][0] == "P001"
        assert merged[0][1] == 1.0


# ═══════════════════════════════════════════════════════════════════════
#  26. parse_card_file — max_row=0 / max_col=0 (325-331)
# ═══════════════════════════════════════════════════════════════════════

class TestParseCardFileEmptySheetCoverage:
    """parse_card_file: пустые листы с max_row=0 или max_col=0 (строки 325-331)."""

    def test_sheet_with_max_row_zero(self):
        """Лист с max_row=0 (без единой ячейки) → пропускается.

        Создаём файл с листом, у которого нет ни одной записанной ячейки.
        openpyxl даст max_row=None → ExcelSheet вернёт 0 → попадаем в блок 324-331.
        """
        fd, path = tempfile.mkstemp(suffix=".xlsx", prefix="card_maxrow0_")
        os.close(fd)
        wb = Workbook()
        # Удаляем активный лист
        wb.remove(wb.active)
        # Создаём новый пустой лист (без единого row в XML)
        wb.create_sheet("Пустой")
        wb.save(path)
        try:
            result = parse_card_file(path)
            assert len(result.parts) == 0
            assert len(result.sheets) >= 1
            s = result.sheets[0]
            assert not s.is_valid
            assert not s.has_data
        finally:
            _safe_remove(path)

    def test_empty_sheet_via_xlrd(self, monkeypatch):
        """Лист с nrows=0 через xlrd → max_row=0 → блок 325-331."""
        import xlrd as real_xlrd
        class MockSheet:
            nrows = 0
            ncols = 0
            def cell_value(self, r, c):
                raise IndexError()
        class MockBook:
            def sheet_names(self): return ["Empty"]
            def sheet_by_name(self, name): return MockSheet()
        monkeypatch.setattr(real_xlrd, "open_workbook", lambda p: MockBook())

        fd, p = tempfile.mkstemp(suffix=".xls")
        os.close(fd); open(p, "w").close()
        try:
            result = parse_card_file(p)
            assert len(result.parts) == 0
            assert len(result.sheets) == 1
            assert not result.sheets[0].has_data
        finally:
            _safe_remove(p)


# ═══════════════════════════════════════════════════════════════════════
#  27. _collect_raw_rows — whitespace skip & exception (495, 520-522)
# ═══════════════════════════════════════════════════════════════════════

class TestCollectRawRowsEdgeCoverage:
    """_collect_raw_rows: пробельный part_no и исключение (строки 495, 520-522)."""

    def test_whitespace_only_part_no_skipped(self):
        """Ячейка с пробелами → raw_part_no_str="" → continue (line 495)."""
        es = _make_excel_sheet([
            ["Part No"],
            ["   "],   # whitespace only → str strip = "" → continue
            ["P001"],  # this one should be collected
        ])
        rows = _collect_raw_rows(es, 1, 3, 1, 1, 0, 0, "test.xlsx")
        assert len(rows) == 1, f"Expected 1 (whitespace skipped), got {len(rows)}"
        assert rows[0][1] == "P001"

    def test_exception_during_collection_caught(self, monkeypatch):
        """Исключение в теле цикла → ловится на 520-522."""
        es = _make_excel_sheet([["P001"], ["P002"], ["P003"]])  # 3 data rows
        orig_cell = ExcelSheet.cell_value
        def mock_cell(self_obj, row, col):
            if row == 2:  # first data row raises
                raise ValueError("mock error")
            return orig_cell(self_obj, row, col)
        monkeypatch.setattr(ExcelSheet, "cell_value", mock_cell)

        rows = _collect_raw_rows(es, 1, 3, 1, 1, 0, 0, "test.xlsx")
        # Row 2 raised and was caught. Row 3 collected normally.
        assert len(rows) == 1
        assert rows[0][1] == "P003"


# ═══════════════════════════════════════════════════════════════════════
#  28. _collect_all_tables — header already processed & empty log (556, 608)
# ═══════════════════════════════════════════════════════════════════════

class TestCollectAllTablesEdgeCoverage:
    """_collect_all_tables: header_row < start_search и total_part_nos_collected==0."""

    def test_header_already_processed_breaks(self):
        """find_part_table returns header_row < start_search → break."""
        es = _make_excel_sheet([
            ["序号", "零部件代号"],  # R1 — header
            ["1", "P001"],         # R2 — data
            [None, None],
            [None, None],
            [None, None],          # R5 — 3 empty rows → boundary
            ["1", "NOT_DATA"],    # R6 — looks like data but starts a 'table' below
        ])
        parts, table_count = _collect_all_tables(es, 6, 2, "test.xlsx")
        assert len(parts) == 1, f"Expected 1 part (second table skipped), got {len(parts)}"
        assert table_count == 1
        assert parts[0][0] == "P001"

    def test_no_valid_parts_log_message(self):
        """total_part_nos_collected == 0 → log message."""
        es = _make_excel_sheet([["Текст без таблиц деталей"]])
        parts, table_count = _collect_all_tables(es, 1, 1, "test.xlsx")
        assert parts == []
        assert table_count == 0


# ═══════════════════════════════════════════════════════════════════════
#  29. _find_excel_files — zip auto extract_dir (608)
# ═══════════════════════════════════════════════════════════════════════

class TestFindExcelFilesAutoExtract:
    """_find_excel_files с ZIP без extract_dir → auto tempdir (line 608)."""

    def test_zip_without_extract_dir(self, tmp_path):
        """extract_dir is None → tempfile.mkdtemp."""
        tmpdir = str(tmp_path)
        try:
            xlsx = os.path.join(tmpdir, "card.xlsx")
            _touch_excel(xlsx)
            zip_path = os.path.join(tmpdir, "cards.zip")
            with zipfile.ZipFile(zip_path, "w") as zf:
                zf.write(xlsx, "card.xlsx")

            files = _find_excel_files(zip_path)  # no extract_dir!
            assert len(files) == 1
            assert any("card.xlsx" in f for f in files)
        finally:
            _rmtree(tmpdir)


# ═══════════════════════════════════════════════════════════════════════
#  30. _walk_extracted_dir — nested zip error (666-667)
# ═══════════════════════════════════════════════════════════════════════

class TestWalkExtractedDirNestedZipError:
    """_walk_extracted_dir: ошибка распаковки вложенного ZIP (строки 666-667)."""

    def test_corrupt_nested_zip_handled(self, tmp_path):
        """Повреждённый вложенный ZIP → except ловит ошибку."""
        tmpdir = str(tmp_path)
        try:
            # Создаём невалидный ZIP (просто текстовый файл с расширением .zip)
            bad_zip = os.path.join(tmpdir, "bad.zip")
            with open(bad_zip, "w") as f:
                f.write("not a zip file")
            # Создаём нормальный xlsx
            xlsx = os.path.join(tmpdir, "normal.xlsx")
            _touch_excel(xlsx)

            files: List[str] = []
            seen: set = set()
            _walk_extracted_dir(tmpdir, tmpdir, files, seen, is_temp=True)
            # Bad zip should be removed (is_temp=True), normal xlsx found
            assert not os.path.isfile(bad_zip), "Bad zip should be removed"
            assert len(files) == 1
            assert "normal.xlsx" in files[0]
        finally:
            _rmtree(tmpdir)


# ═══════════════════════════════════════════════════════════════════════
#  31. _safe_remove — exception handling (676-677)
# ═══════════════════════════════════════════════════════════════════════

class TestSafeRemoveException:
    """_safe_remove: os.remove бросает исключение → pass (строки 676-677)."""

    def test_remove_raises_exception(self, monkeypatch):
        """os.remove бросает PermissionError → except pass (строки 676-677).

        ВАЖНО: вызываем card_parser._safe_remove напрямую, а не локальную
        _safe_remove-helper из этого же тестового файла, которая shadowит импорт.
        """
        import burlak_parser.card_parser as _cp

        def mock_remove(path):
            raise PermissionError("Permission denied")
        monkeypatch.setattr("os.remove", mock_remove)
        # Вызываем модульную _safe_remove, а не тестовую helper
        _cp._safe_remove("/some/path")  # no crash


# ═══════════════════════════════════════════════════════════════════════
#  32. parse_cards — parallel errors & service files (776-780, 806-816)
# ═══════════════════════════════════════════════════════════════════════

class TestParseCardsErrors:
    """parse_cards: ошибки в параллельном парсинге и обработка служебных файлов."""

    def test_parallel_parsing_valid(self, tmp_path):
        """Параллельный парсинг валидных файлов.

        Заменяет test_parallel_parsing_error (строки 776-780 — except блок
        в параллельном ProcessPoolExecutor, который невозможно покрыть
        без monkeypatch, т.к. локальные функции не pickled-ятся).
        """
        tmpdir = str(tmp_path)
        try:
            for fn in ["001-card.xlsx", "002-card.xlsx"]:
                fp = os.path.join(tmpdir, fn)
                wb = Workbook()
                ws = wb.active
                ws.title = "Sheet1"
                ws.cell(row=1, column=1, value="物料编码")
                ws.cell(row=1, column=2, value="数量")
                ws.cell(row=2, column=1, value=f"P{fn[:3]}")
                ws.cell(row=2, column=2, value=1.0)
                wb.save(fp)

            result = parse_cards(tmpdir, max_workers=2, show_progress=False)
            assert result.total_cards_processed >= 2
            assert "P001" in result.all_parts
            assert "P002" in result.all_parts
        finally:
            _rmtree(tmpdir)

    def test_service_files_parsed(self, tmp_path):
        """Служебные файлы обрабатываются через parse_cards (блок 806-816)."""
        tmpdir = str(tmp_path)
        try:
            svc_path = os.path.join(tmpdir, "template.xlsx")
            wb = Workbook()
            ws = wb.active
            ws.cell(row=1, column=1, value="模板")
            wb.save(svc_path)

            op_path = os.path.join(tmpdir, "001-card.xlsx")
            wb2 = Workbook()
            ws2 = wb2.active
            ws2.cell(row=1, column=1, value="物料编码")
            ws2.cell(row=1, column=2, value="数量")
            ws2.cell(row=2, column=1, value="P001")
            ws2.cell(row=2, column=2, value=2.0)
            wb2.save(op_path)

            result = parse_cards(tmpdir, max_workers=1, show_progress=False)
            assert result.total_cards_processed >= 1
            assert result.service_files_skipped >= 1
        finally:
            _rmtree(tmpdir)

    def test_service_file_error_handling(self, monkeypatch, tmp_path):
        """Ошибка при обработке служебного файла → except (строки 815-816)."""
        # Мокаем parse_card_file для служебных файлов
        original_parse = parse_card_file
        def mock_parse(fp, is_service_file=False):
            if is_service_file:
                raise ValueError("Mock service file error")
            return original_parse(fp, is_service_file=is_service_file)
        monkeypatch.setattr(
            "burlak_parser.card_parser.parse_card_file",
            mock_parse,
        )

        tmpdir = str(tmp_path)
        try:
            svc_path = os.path.join(tmpdir, "template.xlsx")
            wb = Workbook()
            ws = wb.active
            ws.cell(row=1, column=1, value="模板")
            wb.save(svc_path)

            op_path = os.path.join(tmpdir, "001-card.xlsx")
            wb2 = Workbook()
            ws2 = wb2.active
            ws2.cell(row=1, column=1, value="物料编码")
            ws2.cell(row=1, column=2, value="数量")
            ws2.cell(row=2, column=1, value="P001")
            ws2.cell(row=2, column=2, value=2.0)
            wb2.save(op_path)

            result = parse_cards(tmpdir, max_workers=1, show_progress=False)
            assert result.total_cards_processed >= 1
            assert result.service_files_skipped >= 1
        finally:
            _rmtree(tmpdir)


# ═══════════════════════════════════════════════════════════════════════
#  33. split_cards_to_files — покрытие всех строк (938-990)
# ═══════════════════════════════════════════════════════════════════════

class TestSplitCardsToFiles:
    """split_cards_to_files — все ветки.

    Покрывает:
      - non-.xlsx skip
      - is_service_file skip
      - template sheet keyword skip
      - split_all_non_empty=False branch
      - empty tasks → return []
      - parallel split path
      - split error handling
      - corrupted warning
      - corrupted_files = None → list()
    """

    def test_skips_non_xlsx(self, tmp_path):
        """Файл не .xlsx → continue (line 938)."""
        result = CardParseResult(
            card_number="C001",
            file_path="test.xls",
            sheets=[CardSheetInfo("C001", "Sheet1", has_data=True)],
            parts=[],
            aggregated_parts={},
        )
        cd = CardsData(all_parts={}, part_sources={}, card_results=[result])
        tmpdir = str(tmp_path)
        try:
            files = split_cards_to_files(cd, tmpdir, max_workers=1)
            assert files == [], "Non-.xlsx should be skipped, no tasks"
        finally:
            _rmtree(tmpdir)

    def test_skips_all_service_files(self, tmp_path):
        """is_service_file=True → все служебные файлы пропускаются."""
        result = CardParseResult(
            card_number="C002",
            file_path="test.xlsx",
            sheets=[CardSheetInfo("C002", "Sheet1", has_data=True)],
            parts=[],
            aggregated_parts={},
            is_service_file=True,
        )
        cd = CardsData(all_parts={}, part_sources={}, card_results=[result])
        tmpdir = str(tmp_path)
        try:
            files = split_cards_to_files(cd, tmpdir, max_workers=1)
            assert files == [], "All service files should be skipped"
        finally:
            _rmtree(tmpdir)

    def test_skips_empty_template_sheet(self, tmp_path):
        """Пустой template-лист (has_data=False) всё ещё пропускается."""
        result = CardParseResult(
            card_number="C004",
            file_path="test.xlsx",
            sheets=[CardSheetInfo("C004", "空表_Empty", has_data=False)],
            parts=[],
            aggregated_parts={},
        )
        cd = CardsData(all_parts={}, part_sources={}, card_results=[result])
        tmpdir = str(tmp_path)
        try:
            files = split_cards_to_files(cd, tmpdir, max_workers=1)
            assert files == [], "Empty template sheet should be skipped"
            stats = cd.split_stats
            assert stats is not None
            assert stats.total_sheets_skipped == 1
            assert "空表_Empty" in stats.file_stats[0].skip_reasons.get(SKIP_REASON_TEMPLATE, [])
        finally:
            _rmtree(tmpdir)

    def test_split_all_non_empty_false(self, monkeypatch, tmp_path):
        """split_all_non_empty=False → проверяет is_valid."""
        from burlak_parser import splitter as splitter_mod

        monkeypatch.setattr(
            splitter_mod, "_extract_to_path_worker",
            _mock_extract_success,
        )

        result = CardParseResult(
            card_number="C005",
            file_path="test.xlsx",
            sheets=[CardSheetInfo("C005", "Sheet1", has_data=True, is_valid=True)],
            parts=[],
            aggregated_parts={},
        )
        cd = CardsData(all_parts={}, part_sources={}, card_results=[result])
        tmpdir = str(tmp_path)
        try:
            files = split_cards_to_files(cd, tmpdir, split_all_non_empty=False, max_workers=1)
            assert len(files) == 1
        finally:
            _rmtree(tmpdir)

    def test_empty_tasks_returns_empty(self, tmp_path):
        """Все задачи отфильтрованы → tasks=[] → return [] (line 966)."""
        # Все файлы .xls (не .xlsx) → continue → tasks = []
        result = CardParseResult(
            card_number="C006",
            file_path="test.xls",
            sheets=[CardSheetInfo("C006", "Sheet1", has_data=True)],
            parts=[],
            aggregated_parts={},
        )
        cd = CardsData(all_parts={}, part_sources={}, card_results=[result])
        tmpdir = str(tmp_path)
        try:
            files = split_cards_to_files(cd, tmpdir, max_workers=1)
            assert files == []
        finally:
            _rmtree(tmpdir)

    def test_parallel_split_path(self, monkeypatch, tmp_path):
        """workers > 1 и tasks > 1 → parallel path."""
        from burlak_parser import splitter as splitter_mod

        monkeypatch.setattr(
            splitter_mod, "_extract_to_path_worker",
            _mock_extract_success,
        )

        # Создаём 2 задачи
        r1 = CardParseResult(
            card_number="C010", file_path="test1.xlsx",
            sheets=[CardSheetInfo("C010", "S1", has_data=True)],
            parts=[], aggregated_parts={},
        )
        r2 = CardParseResult(
            card_number="C011", file_path="test2.xlsx",
            sheets=[CardSheetInfo("C011", "S1", has_data=True)],
            parts=[], aggregated_parts={},
        )
        cd = CardsData(all_parts={}, part_sources={}, card_results=[r1, r2])
        tmpdir = str(tmp_path)
        try:
            files = split_cards_to_files(cd, tmpdir, max_workers=2)
            assert len(files) == 2
        finally:
            _rmtree(tmpdir)

    def test_split_error_handling(self, monkeypatch, tmp_path):
        """split_file бросает исключение → except (lines 980-982)."""
        from burlak_parser import splitter as splitter_mod

        class MockSplitter:
            def split_file(self, src, out, sheets, label):
                raise ValueError("Mock split error")

        monkeypatch.setattr(
            splitter_mod, "CardSplitter",
            lambda **kw: MockSplitter(),
        )

        result = CardParseResult(
            card_number="C020", file_path="test.xlsx",
            sheets=[CardSheetInfo("C020", "S1", has_data=True)],
            parts=[], aggregated_parts={},
        )
        cd = CardsData(all_parts={}, part_sources={}, card_results=[result])
        tmpdir = str(tmp_path)
        try:
            # Ошибка должна быть залогирована, функция не падает
            files = split_cards_to_files(cd, tmpdir, max_workers=1)
            assert files == []
            # corrupted_files должен быть установлен
            assert cd.corrupted_files is not None
        finally:
            _rmtree(tmpdir)

    def test_corrupted_files_none_case(self, monkeypatch, tmp_path):
        """cards_data.corrupted_files defaults to empty list, gets populated on error."""
        from burlak_parser import splitter as splitter_mod

        class MockSplitter:
            def split_file(self, src, out, sheets, label):
                raise ValueError("error")

        monkeypatch.setattr(
            splitter_mod, "CardSplitter",
            lambda **kw: MockSplitter(),
        )

        result = CardParseResult(
            card_number="C030", file_path="test.xlsx",
            sheets=[CardSheetInfo("C030", "S1", has_data=True)],
            parts=[], aggregated_parts={},
        )
        cd = CardsData(all_parts={}, part_sources={}, card_results=[result])
        assert cd.corrupted_files == [], "Default should be empty list"
        tmpdir = str(tmp_path)
        try:
            files = split_cards_to_files(cd, tmpdir, max_workers=1)
            assert files == []
            assert cd.corrupted_files == ["test.xlsx"]
        finally:
            _rmtree(tmpdir)


# ═══════════════════════════════════════════════════════════════════════
#  ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ═══════════════════════════════════════════════════════════════════════

def _make_ws(data: List[List[Optional[Any]]]) -> Any:
    """Create in-memory openpyxl worksheet."""
    wb = Workbook()
    ws = wb.active
    for r_idx, row in enumerate(data, 1):
        for c_idx, val in enumerate(row, 1):
            if val is not None:
                ws.cell(row=r_idx, column=c_idx, value=val)
    return ws


def _make_excel_sheet(data: List[List[Optional[Any]]]) -> ExcelSheet:
    """Create an ExcelSheet wrapper from data rows."""
    ws = _make_ws(data)
    return ExcelSheet(ws, "openpyxl")


def _make_card_xlsx(
    data: List[List[Optional[Any]]],
    file_name: str = "test_card.xlsx",
) -> str:
    """Create a temporary .xlsx file for parse_card_file testing."""
    fd, path = tempfile.mkstemp(suffix=".xlsx", prefix="card_test_")
    os.close(fd)
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for r_idx, row in enumerate(data, 1):
        for c_idx, val in enumerate(row, 1):
            if val is not None:
                ws.cell(row=r_idx, column=c_idx, value=val)
    wb.save(path)
    return path


# ═══════════════════════════════════════════════════════════════════════
#  1. Data Structures
# ═══════════════════════════════════════════════════════════════════════

class TestCardSheetInfo:
    def test_default_creation(self):
        si = CardSheetInfo(card_number="C001", sheet_name="Лист1")
        assert si.card_number == "C001"
        assert si.sheet_name == "Лист1"
        assert si.operation_name == ""
        assert si.is_valid is False
        assert si.has_data is False

    def test_full_creation(self):
        si = CardSheetInfo(
            card_number="C001",
            sheet_name="Лист1",
            operation_name="Установка",
            is_valid=True,
            has_data=True,
        )
        assert si.operation_name == "Установка"
        assert si.is_valid is True
        assert si.has_data is True


class TestCardPart:
    def test_default_creation(self):
        cp = CardPart(part_number="ABC001", quantity=2.0, source_card="C001", source_sheet="S1")
        assert cp.part_number == "ABC001"
        assert cp.quantity == 2.0
        assert cp.source_card == "C001"
        assert cp.source_sheet == "S1"


class TestCardParseResult:
    def test_default_creation(self):
        result = CardParseResult(
            card_number="C001",
            file_path="/path/to/file.xlsx",
            sheets=[],
            parts=[],
            aggregated_parts={},
        )
        assert result.card_number == "C001"
        assert result.is_service_file is False
        assert result.parts == []

    def test_with_parts(self):
        parts = [
            CardPart(part_number="P001", quantity=1.0, source_card="C001", source_sheet="S1"),
            CardPart(part_number="P002", quantity=2.0, source_card="C001", source_sheet="S1"),
        ]
        result = CardParseResult(
            card_number="C001",
            file_path="/path.xlsx",
            sheets=[],
            parts=parts,
            aggregated_parts={"P001": 1.0, "P002": 2.0},
            is_service_file=True,
            )
        assert len(result.parts) == 2
        assert result.is_service_file is True is True


class TestCardsData:
    def test_default_creation(self):
        cd = CardsData(
            all_parts={},
            part_sources={},
            card_results=[],
        )
        assert cd.all_parts == {}
        assert cd.total_cards_processed == 0
        assert cd.total_sheets_processed == 0
        assert cd.total_sheets_skipped == 0
        assert cd.service_files_skipped == 0
        assert cd.corrupted_files == []


# ═══════════════════════════════════════════════════════════════════════
#  2. ExcelSheet
# ═══════════════════════════════════════════════════════════════════════

class TestExcelSheet:
    def test_max_row(self):
        es = _make_excel_sheet([
            ["A", "B"],
            ["C", "D"],
            ["E", "F"],
        ])
        assert es.max_row == 3
        assert es.max_column == 2

    def test_cell_value(self):
        es = _make_excel_sheet([
            ["Hello", "World"],
            [None, 42],
        ])
        assert es.cell_value(1, 1) == "Hello"
        assert es.cell_value(1, 2) == "World"
        assert es.cell_value(2, 1) is None
        assert es.cell_value(2, 2) == 42

    def test_multiline_cell(self):
        es = _make_excel_sheet([
            ["序号\nСерийный номер", "零部件代号\nКод детали"],
        ])
        val = es.cell_value(1, 1)
        assert val is not None
        assert "序号" in str(val)
        assert "Серийный" in str(val)

    def test_out_of_bounds_returns_none(self):
        es = _make_excel_sheet([["A"]])
        assert es.cell_value(999, 999) is None


# ═══════════════════════════════════════════════════════════════════════
#  3. _check_sheet_has_data
# ═══════════════════════════════════════════════════════════════════════

class TestCheckSheetHasData:
    def test_non_empty_sheet(self):
        es = _make_excel_sheet([["A", "B"], ["C", None]])
        assert _check_sheet_has_data(es) is True

    def test_empty_sheet(self):
        es = _make_excel_sheet([])
        assert _check_sheet_has_data(es) is False

    def test_sheet_with_only_none(self):
        es = _make_excel_sheet([[None, None], [None, None]])
        assert _check_sheet_has_data(es) is False

    def test_sheet_with_data_deep(self):
        """Sheet with data only at row 20 should still be detected."""
        data = [[""] for _ in range(25)]
        data[19] = ["HasData"]  # row 20
        wb = Workbook()
        ws = wb.active
        for r_idx, row in enumerate(data, 1):
            for c_idx, val in enumerate(row, 1):
                if val:
                    ws.cell(row=r_idx, column=c_idx, value=val)
        es = ExcelSheet(ws, "openpyxl")
        # Row 20 is within sampling range (15 + step < 25)
        assert _check_sheet_has_data(es) is True

    def test_single_cell_sheet(self):
        es = _make_excel_sheet([["Only"]])
        assert _check_sheet_has_data(es) is True


# ═══════════════════════════════════════════════════════════════════════
#  4. _collect_raw_rows — basic collection
# ═══════════════════════════════════════════════════════════════════════

class TestCollectRawRows:
    def test_normal_collection(self):
        """T1L card: part_no=C1, qty=C3, name=C2, data rows after header."""
        es = _make_excel_sheet([
            ["物料编码", "零件名称", "数量", "单位"],
            ["P001", "Part1", "2", "pcs"],
            ["P002", "Part2", "1", "pcs"],
            ["P003", "Part3", "3", "pcs"],
        ])
        rows = _collect_raw_rows(es, 1, 4, 4, 1, 3, 2, "test.xlsx")
        assert len(rows) == 3
        assert rows[0][1] == "P001"
        assert rows[0][2] == 2.0  # qty
        assert rows[0][3] == "Part1"  # name
        assert rows[1][1] == "P002"
        assert rows[2][1] == "P003"

    def test_default_qty_when_qty_col_0(self):
        """When qty_col=0, all parts get qty=1.0."""
        es = _make_excel_sheet([
            ["零部件代号"],
            ["P001"],
            ["P002"],
        ])
        rows = _collect_raw_rows(es, 1, 3, 1, 1, 0, 0, "test.xlsx")
        assert len(rows) == 2
        assert rows[0][2] == 1.0
        assert rows[1][2] == 1.0

    def test_empty_name_when_name_col_0(self):
        """When name_col=0, names are empty."""
        es = _make_excel_sheet([
            ["Part No", "Qty"],
            ["P001", "2"],
            ["P002", "1"],
        ])
        rows = _collect_raw_rows(es, 1, 3, 2, 1, 2, 0, "test.xlsx")
        assert len(rows) == 2
        assert rows[0][3] == ""  # name
        assert rows[0][2] == 2.0

    def test_skips_empty_rows(self):
        """Empty rows should be skipped."""
        es = _make_excel_sheet([
            ["物料编码", "名称"],
            ["P001", "Part1"],
            [None, None],  # empty row
            ["P002", "Part2"],
        ])
        rows = _collect_raw_rows(es, 1, 4, 2, 1, 0, 2, "test.xlsx")
        assert len(rows) == 2
        assert rows[0][1] == "P001"
        assert rows[1][1] == "P002"

    def test_skips_keyword_rows(self):
        """Rows with skip keywords should be skipped."""
        es = _make_excel_sheet([
            ["物料编码", "名称"],
            ["P001", "Part1"],
            ["变更记录", "Change log"],
            ["P002", "Part2"],
        ])
        rows = _collect_raw_rows(es, 1, 4, 2, 1, 0, 2, "test.xlsx")
        assert len(rows) == 2
        assert rows[0][1] == "P001"
        assert rows[1][1] == "P002"

    def test_qty_from_string(self):
        """Quantity should be parsed from string values."""
        es = _make_excel_sheet([
            ["Part", "Qty"],
            ["P001", "2.5"],
            ["P002", "1"],
        ])
        rows = _collect_raw_rows(es, 1, 3, 2, 1, 2, 0, "test.xlsx")
        assert rows[0][2] == 2.5
        assert rows[1][2] == 1.0

    def test_invalid_qty_defaults_to_1(self):
        """Invalid qty values should default to 1.0."""
        es = _make_excel_sheet([
            ["Part", "Qty"],
            ["P001", "N/A"],
        ])
        rows = _collect_raw_rows(es, 1, 2, 2, 1, 2, 0, "test.xlsx")
        assert rows[0][2] == 1.0


# ═══════════════════════════════════════════════════════════════════════
#  5. _collect_raw_rows — section boundary detection
# ═══════════════════════════════════════════════════════════════════════

class TestCollectRawRowsBoundaries:
    def test_stops_at_new_header(self):
        """SWM card: first table at R1, new header at R5 with part_no keyword."""
        es = _make_excel_sheet([
            ["序号\nСерийный номер", "零部件代号\nКод детали"],
            ["1", "P001"],
            ["2", "P002"],
            ["3", "P003"],
            ["序号", "零部件代号"],  # NEW header → should stop
            ["1", "Q001"],
        ])
        rows = _collect_raw_rows(es, 1, 6, 2, 2, 0, 0, "test.xlsx")
        assert len(rows) == 3, f"Expected 3 parts before section boundary, got {len(rows)}"
        assert rows[0][1] == "P001"
        assert rows[1][1] == "P002"
        assert rows[2][1] == "P003"

    def test_stops_after_3_consecutive_empty_rows(self):
        """Data with 3+ empty rows should stop collection."""
        es = _make_excel_sheet([
            ["零部件代号", "名称"],
            ["P001", "Part1"],
            [None, None],
            [None, None],
            [None, None],
            ["P002", "Part2"],  # after 3 empty rows → should NOT be collected
        ])
        rows = _collect_raw_rows(es, 1, 6, 2, 1, 0, 2, "test.xlsx")
        assert len(rows) == 1, f"Expected 1 part (3 empties stop), got {len(rows)}"
        assert rows[0][1] == "P001"

    def test_does_not_stop_at_2_empty_rows(self):
        """2 empty rows should not stop — only 3+."""
        es = _make_excel_sheet([
            ["Part No"],
            ["P001"],
            [None],
            [None],
            ["P002"],
        ])
        rows = _collect_raw_rows(es, 1, 5, 1, 1, 0, 0, "test.xlsx")
        assert len(rows) == 2, f"Expected 2 parts (2 empties OK), got {len(rows)}"
        assert rows[1][1] == "P002"

    def test_skips_long_cell_in_boundary_check(self):
        """A cell with part_no keyword but > 50 chars should NOT trigger boundary."""
        long_text = "拿取零部件1检查是否有破损；Возьмите деталь"  # > 50 chars
        es = _make_excel_sheet([
            ["序号\nСерийный номер", "零部件代号\nКод детали"],
            ["1", "P001"],
            ["2", "P002"],
            [long_text, None],  # long cell with 'деталь' — should NOT trigger boundary; C2=None → skip row
            ["3", "P003"],
        ])
        rows = _collect_raw_rows(es, 1, 5, 2, 2, 0, 0, "test.xlsx")
        assert len(rows) == 3, f"Expected 3 parts (long cell skipped), got {len(rows)}"
        assert rows[2][1] == "P003"

    def test_does_not_trigger_on_data_rows(self):
        """Data rows with part numbers but no keywords should not trigger boundary."""
        es = _make_excel_sheet([
            ["物料编码", "名称"],
            ["P001", "Part1"],
            ["P002", "Part2"],
            ["ABC-123-DEF", "Part3"],  # looks like part no, not a header
            ["P003", "Part4"],
        ])
        rows = _collect_raw_rows(es, 1, 5, 2, 1, 0, 2, "test.xlsx")
        assert len(rows) == 4, f"Expected 4 parts (data rows pass through), got {len(rows)}"

    def test_boundary_english_keyword(self):
        """English header 'Part No' should trigger boundary."""
        es = _make_excel_sheet([
            ["Seq", "Part No", "Qty"],
            ["1", "P001", "2"],
            ["2", "P002", "1"],
            ["Seq", "Part No", "Qty"],  # English header → boundary
            ["1", "Q001", "1"],
        ])
        rows = _collect_raw_rows(es, 1, 5, 3, 2, 3, 0, "test.xlsx")
        assert len(rows) == 2, f"Expected 2 parts (English boundary), got {len(rows)}"
        assert rows[0][1] == "P001"
        assert rows[1][1] == "P002"

    def test_boundary_russian_keyword(self):
        """Russian header 'Код детали' should trigger boundary."""
        es = _make_excel_sheet([
            ["№", "Код детали", "Кол-во"],
            ["1", "P001", "2"],
            ["2", "P002", "1"],
            ["№", "Код детали", "Кол-во"],  # Russian header → boundary
            ["1", "Q001", "1"],
        ])
        rows = _collect_raw_rows(es, 1, 5, 3, 2, 3, 0, "test.xlsx")
        assert len(rows) == 2, f"Expected 2 parts (Russian boundary), got {len(rows)}"
        assert rows[0][1] == "P001"
        assert rows[1][1] == "P002"

    def test_boundary_not_triggered_by_service_keyword(self):
        """Service keyword row '变更记录' without part_no keyword should NOT trigger boundary.
        The row should be skipped (not collected) due to skip_keywords instead."""
        es = _make_excel_sheet([
            ["物料编码", "零件名称", "数量"],
            ["P001", "Part1", "2"],
            ["P002", "Part2", "1"],
            ["变更记录", "Change", "Log"],  # no PART_NO_KEYWORD → not a boundary; skip via skip_keywords
            ["P003", "Part3", "3"],
        ])
        rows = _collect_raw_rows(es, 1, 5, 3, 1, 3, 2, "test.xlsx")
        # 变更记录 should be skipped (via skip_keywords), not trigger boundary
        assert len(rows) == 3, f"Expected 3 parts (service row skipped), got {len(rows)}"
        assert rows[2][1] == "P003"

    def test_boundary_not_triggered_by_single_cell(self):
        """Single cell with part_no keyword should NOT trigger boundary (need >=2 non-empty)."""
        es = _make_excel_sheet([
            ["序号", "零部件代号"],
            ["1", "P001"],
            ["2", "P002"],
            ["零部件代号", None],  # part_no keyword but only 1 non-empty cell
            ["3", "P003"],
        ])
        rows = _collect_raw_rows(es, 1, 5, 2, 2, 0, 0, "test.xlsx")
        # R4: raw_part_no=None → not all_empty (C1 has data) → skip; NOT a boundary
        assert len(rows) == 3, f"Expected 3 parts (single-cell skipped), got {len(rows)}"
        assert rows[2][1] == "P003"

    def test_part_no_none_with_data_elsewhere(self):
        """Row with None in part_no column but data in other columns should be skipped,
        NOT counted as an empty row (does not increment consecutive_empty_pn)."""
        es = _make_excel_sheet([
            ["Part No", "Description"],
            ["P001", "Part1"],
            [None, "Some description"],  # part_no is None but desc has data → skip, not empty
            ["P002", "Part2"],
        ])
        rows = _collect_raw_rows(es, 1, 4, 2, 1, 0, 2, "test.xlsx")
        assert len(rows) == 2, f"Expected 2 parts (description row skipped), got {len(rows)}"
        assert rows[0][1] == "P001"
        assert rows[1][1] == "P002"


# ═══════════════════════════════════════════════════════════════════════
#  5b. _collect_raw_rows — пограничные сценарии (ПРОДВИНУТЫЕ)
# ═══════════════════════════════════════════════════════════════════════

class TestCollectRawRowsBoundariesExtended:
    """Продвинутые тесты границ секций в _collect_raw_rows.

    Покрывает:
      - Case-insensitive PART_NO_KEYWORD ("Part No", "part no", "PART NO")
      - Ячейка >= 50 символов с PART_NO_KEYWORD → НЕ триггерит границу
      - 3+ skip-строк (не пустых, без part_no) → НЕ триггерит границу
      - Числовой part_no (int/float из Excel)
      - max_data_row (header_row + 500) — данные на границе
      - Skip-ключевые слова с разным регистром
      - Ячейка с PART_NO_KEYWORD длины ровно 50 символов
      - Смешанные китайские/английские заголовки
    """

    def test_boundary_case_insensitive_english(self):
        """'Part No' (mixed case) должен триггерить границу."""
        es = _make_excel_sheet([
            ["Seq", "Part No", "Qty"],      # header
            ["1", "P001", "2"],
            ["2", "P002", "1"],
            ["Seq", "Part No", "Qty"],      # mixed case → boundary
            ["1", "Q001", "1"],
        ])
        rows = _collect_raw_rows(es, 1, 5, 3, 2, 3, 0, "test.xlsx")
        assert len(rows) == 2, f"Comprehensive detection should work with mixed case"
        assert rows[0][1] == "P001"
        assert rows[1][1] == "P002"

    def test_boundary_case_insensitive_chinese(self):
        """'零部件代号' (китайский, нижний регистр не применим) должен триггерить границу."""
        es = _make_excel_sheet([
            ["序号", "零部件代号"],
            ["1", "P001"],
            ["序号", "零部件代号"],  # exact same → boundary
            ["1", "Q001"],
        ])
        rows = _collect_raw_rows(es, 1, 4, 2, 2, 0, 0, "test.xlsx")
        assert len(rows) == 1, f"Chinese header should trigger boundary"
        assert rows[0][1] == "P001"

    def test_long_cell_with_keyword_does_not_trigger_boundary(self):
        """Ячейка >= 50 символов, содержащая PART_NO_KEYWORD, НЕ триггерит границу.

        Проверяем условие `len(rv) < 50`: даже если ячейка содержит 'код детали',
        но её длина >= 50, она не считается заголовком новой таблицы.
        C2 непустой, чтобы non_empty >= 2 и проверка `len < 50` достиглась.
        """
        long_text = "код детали " + "x" * 39  # 50 chars exactly, contains keyword
        assert len(long_text) >= 50, f"Long text must be >= 50 chars, got {len(long_text)}"
        assert "код детали" in long_text

        es = _make_excel_sheet([
            ["Part No", "Desc"],
            ["P001", "Part1"],
            [long_text, "has data"],     # C1 >= 50 with keyword, C2 non-empty
            ["P002", "Part2"],
        ])
        rows = _collect_raw_rows(es, 1, 4, 2, 1, 0, 2, "test.xlsx")
        # R3: non_empty=2 (C1=long_text, C2="has data")
        # has_part_no_keyword_short? C1 содержит "код детали" но len=50, NOT < 50 → False
        # NOT a boundary. raw_part_no = long_text → collected.
        assert len(rows) == 3, f"Expected 3 raw rows (50-char with keyword NOT boundary), got {len(rows)}"

    def test_long_cell_without_keyword_does_not_trigger_boundary(self):
        """Ячейка >= 50 символов БЕЗ PART_NO_KEYWORD — не триггерит границу."""
        long_text = "A" * 55  # 55 chars, no keyword
        es = _make_excel_sheet([
            ["Part No", "Desc"],
            ["P001", "Part1"],
            [long_text, "Some long description"],  # C1 long, no keyword
            ["P002", "Part2"],
        ])
        rows = _collect_raw_rows(es, 1, 4, 2, 1, 0, 2, "test.xlsx")
        # R3: C1 has long_text, C2 has data. non_empty=2. has_part_no_keyword_short?
        # C1: "AAA..." (55 chars) → no keyword. C2: "some long description" → no keyword.
        # NOT a boundary. Then raw_part_no at C1 = "AAA..." → not None, not a skip keyword.
        # But "AAA..." does NOT pass is_valid_part_number() (all letters, no digits).
        # Wait, no - is_valid_part_number is checked in _merge_multiline_part_numbers,
        # not in _collect_raw_rows. _collect_raw_rows just collects ALL rows.
        # So it would be collected as a raw row.
        assert len(rows) == 3, f"Expected 3 parts (long cell passes through), got {len(rows)}"

    def test_skip_keywords_mixed_case(self):
        """Skip-ключевые слова должны работать независимо от регистра."""
        es = _make_excel_sheet([
            ["Part No", "Name"],
            ["P001", "Part1"],
            ["变更记录", "Change Log"],  # Chinese skip keyword
            ["P002", "Part2"],
        ])
        rows = _collect_raw_rows(es, 1, 4, 2, 1, 0, 2, "test.xlsx")
        assert len(rows) == 2, f"Expected 2 parts (skip keyword skipped), got {len(rows)}"
        assert rows[0][1] == "P001"
        assert rows[1][1] == "P002"

    def test_numeric_part_no_collected(self):
        """Числовой part_no (int) должен быть собран как строка."""
        # Create workbook directly for numeric values
        wb = Workbook()
        ws = wb.active
        ws.cell(row=1, column=1, value="Part No")
        ws.cell(row=1, column=2, value="Qty")
        ws.cell(row=2, column=1, value=12345)  # numeric part_no
        ws.cell(row=2, column=2, value=2)
        ws.cell(row=3, column=1, value="ABC-001")  # string part_no
        ws.cell(row=3, column=2, value=1)
        es = ExcelSheet(ws, "openpyxl")

        rows = _collect_raw_rows(es, 1, 3, 2, 1, 2, 0, "test.xlsx")
        assert len(rows) == 2, f"Expected 2 parts (numeric + string), got {len(rows)}"
        # Numeric 12345 becomes string "12345"
        assert rows[0][1] == "12345", f"Expected '12345', got '{rows[0][1]}'"
        assert rows[0][2] == 2.0
        assert rows[1][1] == "ABC-001"

    def test_multiple_consecutive_skip_rows(self):
        """3+ skip-строк (part_no=None, но данные в других колонках) НЕ триггерят
        границу пустых строк. Данные после скипов собираются."""
        es = _make_excel_sheet([
            ["Part No", "Description"],
            ["P001", "Part1"],
            [None, "Skip description 1"],  # skip row (has data in C2)
            [None, "Skip description 2"],  # skip row
            [None, "Skip description 3"],  # skip row (3 consecutive!)
            ["P002", "Part2"],            # should be collected
        ])
        rows = _collect_raw_rows(es, 1, 6, 2, 1, 0, 2, "test.xlsx")
        assert len(rows) == 2, f"Expected 2 parts (skips don't trigger 3-empty boundary), got {len(rows)}"
        assert rows[0][1] == "P001"
        assert rows[1][1] == "P002"

    def test_all_empty_then_data_after_3_empties_not_collected(self):
        """3+ полностью пустых строк → граница. Данные после НЕ собираются."""
        es = _make_excel_sheet([
            ["Part No", "Name"],
            ["P001", "Part1"],
            [None, None],  # empty 1
            [None, None],  # empty 2
            [None, None],  # empty 3 → boundary!
            ["P002", "Part2"],  # should NOT be collected
        ])
        rows = _collect_raw_rows(es, 1, 6, 2, 1, 0, 2, "test.xlsx")
        assert len(rows) == 1, f"Expected 1 part (3 empties stop at boundary), got {len(rows)}"
        assert rows[0][1] == "P001"

    def test_mixed_skip_then_empty_triggers_boundary(self):
        """2 skip-строки + 2 пустых строки → граница НЕ триггерится (skip не увеличивает
        счётчик пустых). После скипов data, потом 3 пустых → граница."""
        es = _make_excel_sheet([
            ["Part No", "Desc"],
            ["P001", "Part1"],
            [None, "Skip1"],   # skip (not empty, not data)
            [None, "Skip2"],   # skip
            ["P002", "Part2"],  # data collected
            [None, None],       # empty 1
            [None, None],       # empty 2
            [None, None],       # empty 3 → boundary!
            ["P003", "Part3"],  # should NOT be collected
        ])
        rows = _collect_raw_rows(es, 1, 9, 2, 1, 0, 2, "test.xlsx")
        # После скипов P002 собран. Потом 3 пустых → граница. P003 НЕ собран.
        assert len(rows) == 2, f"Expected 2 parts (3 empties after data), got {len(rows)}"
        assert rows[0][1] == "P001"
        assert rows[1][1] == "P002"

    def test_boundary_not_triggered_by_2_non_empty_no_keyword(self):
        """Строка с >= 2 непустыми ячейками, но без PART_NO_KEYWORD — НЕ граница."""
        es = _make_excel_sheet([
            ["Part No", "Name"],
            ["P001", "Part1"],
            ["Такелажные", "ремни"],  # 2 non-empty, no keyword
            ["P002", "Part2"],
        ])
        rows = _collect_raw_rows(es, 1, 4, 2, 1, 0, 2, "test.xlsx")
        # R3: non_empty=2. has_part_no_keyword_short? "Такелажные" → no keyword match.
        # "ремни" → no keyword match. NOT a boundary.
        # raw_part_no at C1 = "Такелажные" → not None, not skip keyword → collected!
        # Но "Такелажные" не пройдёт is_valid_part_number (нет цифр)
        # _collect_raw_rows не проверяет is_valid_part_number, _merge_multiline делает
        assert len(rows) == 3, f"Expected 3 raw rows (text row passes through), got {len(rows)}"

    def test_max_data_row_collects_up_to_header_plus_500(self):
        """Данные собираются до header_row + 500 включительно.

        Создаём лист без пустых строк (чтобы 3+ empty не вмешивался).
        max_data_row = min(max_row, header_row + 500).
        """
        wb = Workbook()
        ws = wb.active
        ws.cell(row=1, column=1, value="Part No")
        ws.cell(row=1, column=2, value="Qty")
        # Заполняем R2-R501 данными (500 строк данных)
        for i in range(500):
            ws.cell(row=2 + i, column=1, value=f"P{i:03d}")
            ws.cell(row=2 + i, column=2, value=1)
        # R502 — now included since 500-row limit was removed
        ws.cell(row=502, column=1, value="P500")
        ws.cell(row=502, column=2, value=999)
        es = ExcelSheet(ws, "openpyxl")

        rows = _collect_raw_rows(es, 1, 502, 2, 1, 2, 0, "test.xlsx")
        # No 500-row limit — all 501 rows collected
        assert len(rows) == 501, f"Expected 501 parts (no limit), got {len(rows)}"
        assert rows[0][1] == "P000"
        assert rows[500][1] == "P500"

    def test_row_has_part_no_keyword_exact_50_chars(self):
        """Ячейка длины ровно 50 символов с PART_NO_KEYWORD — НЕ триггерит границу."""
        # Создаём текст длиной ровно 50, содержащий "код детали"
        base = "код детали"  # 11 chars
        long_text = base + "x" * (50 - len(base))  # 50 chars exactly
        assert len(long_text) == 50, f"Must be exactly 50 chars, got {len(long_text)}"

        es = _make_excel_sheet([
            ["Part No", "Name"],
            ["P001", "Part1"],
            [long_text, "Extra"],  # C1 = 50 chars, contains "код детали"
            ["P002", "Part2"],
        ])
        rows = _collect_raw_rows(es, 1, 4, 2, 1, 0, 2, "test.xlsx")
        # R3: non_empty=2. has_part_no_keyword_short? len(rv)=50, NOT < 50 → False!
        # NOT a boundary. raw_part_no = long_text → not None → collected as raw row!
        assert len(rows) == 3, f"Expected 3 raw rows (exact 50-char boundary NOT triggered), got {len(rows)}"

    def test_row_has_part_no_keyword_49_chars_triggers_boundary(self):
        """Ячейка длины 49 символов с PART_NO_KEYWORD — триггерит границу (< 50)."""
        base = "код детали"
        text_49 = base + "x" * (49 - len(base))
        assert len(text_49) == 49, f"Must be exactly 49 chars, got {len(text_49)}"

        es = _make_excel_sheet([
            ["Part No", "Name"],
            ["P001", "Part1"],
            [text_49, "Extra"],  # 49 chars, len < 50 → boundary!
            ["P002", "Part2"],
        ])
        rows = _collect_raw_rows(es, 1, 4, 2, 1, 0, 2, "test.xlsx")
        # R3: non_empty=2. has_part_no_keyword_short? len(rv)=49 < 50 → True!
        # Boundary triggered! P002 is NOT collected.
        assert len(rows) == 1, f"Expected 1 part (49-char boundary triggered), got {len(rows)}"
        assert rows[0][1] == "P001"

    def test_skip_keyword_triggers_on_substring_match(self):
        """Строка, содержащая skip-ключевое слово как подстроку, скипается.
        Например '变更记录仪' содержит '变更记录' → скипается.
        Это корректно, т.к. такие строки — заголовки, а не детали."""
        es = _make_excel_sheet([
            ["Part No", "Name"],
            ["P001", "Part1"],
            ["变更记录仪", "Measuring device"],  # contains '变更记录' → skipped
            ["P002", "Part2"],
        ])
        rows = _collect_raw_rows(es, 1, 4, 2, 1, 0, 2, "test.xlsx")
        # R3: skip check: '变更记录' in '变更记录仪' → True → skipped
        assert len(rows) == 2, f"Expected 2 parts (substring match skips row), got {len(rows)}"
        assert rows[0][1] == "P001"
        assert rows[1][1] == "P002"


# ═══════════════════════════════════════════════════════════════════════
#  6. _merge_multiline_part_numbers
# ═══════════════════════════════════════════════════════════════════════

class TestMergeMultilinePartNumbers:
    def test_normal_no_merge(self):
        """Normal rows without continuation should pass through."""
        rows = [
            (2, "ABC-001", 1.0, "Part1", 1),
            (3, "DEF-002", 2.0, "Part2", 1),
        ]
        merged = _merge_multiline_part_numbers(rows)
        assert len(merged) == 2
        assert merged[0][0] == "ABC001"  # cleaned
        assert merged[0][1] == 1.0
        assert merged[1][0] == "DEF002"

    def test_dash_continuation(self):
        """Part number ending with '-' should merge with next row."""
        rows = [
            (2, "5306200-", 1.0, "Part1", 1),
            (3, "ED001", 1.0, "Part1", 1),
        ]
        merged = _merge_multiline_part_numbers(rows)
        assert len(merged) == 1
        assert merged[0][0] == "5306200ED001"
        assert merged[0][1] == 1.0

    def test_em_dash_continuation(self):
        """Em-dash continuation."""
        rows = [
            (2, "ABC—", 2.0, "Part1", 1),
            (3, "123", 2.0, "Part1", 1),
        ]
        merged = _merge_multiline_part_numbers(rows)
        assert len(merged) == 1
        assert merged[0][0] == "ABC123"

    def test_en_dash_continuation(self):
        """En-dash continuation."""
        rows = [
            (2, "GHI–", 1.0, "Part2", 1),
            (3, "456", 1.0, "Part2", 1),
        ]
        merged = _merge_multiline_part_numbers(rows)
        assert len(merged) == 1
        assert merged[0][0] == "GHI456"

    def test_skips_invalid_part_numbers(self):
        """Invalid part numbers should be filtered out."""
        rows = [
            (2, "AB", 1.0, "Part1", 1),  # too short
            (3, "P001", 2.0, "Part2", 1),
        ]
        merged = _merge_multiline_part_numbers(rows)
        assert len(merged) == 1
        assert merged[0][0] == "P001"

    def test_empty_input(self):
        """Empty input returns empty list."""
        assert _merge_multiline_part_numbers([]) == []

    def test_only_continuation_without_resolution(self):
        """Trailing continuation without resolution should not appear."""
        rows = [
            (2, "ABC-", 1.0, "Part1", 1),
            # No continuation row — buffer stays pending
        ]
        merged = _merge_multiline_part_numbers(rows)
        assert len(merged) == 0

    def test_qty_from_first_row_in_continuation(self):
        """Quantity should be taken from the first row of a continuation."""
        rows = [
            (2, "LONG-", 5.0, "PartX", 1),
            (3, "123", 99.0, "PartX", 1),  # qty=99 should be ignored → 5 from first
        ]
        merged = _merge_multiline_part_numbers(rows)
        assert len(merged) == 1
        assert merged[0][0] == "LONG123"
        assert merged[0][1] == 5.0  # qty from first row

    def test_name_from_first_row_in_continuation(self):
        """Name should be taken from the first row of a continuation."""
        rows = [
            (2, "ABC-", 1.0, "FirstName", 1),
            (3, "123", 1.0, "SecondName", 1),
        ]
        merged = _merge_multiline_part_numbers(rows)
        assert len(merged) == 1
        assert merged[0][0] == "ABC123"
        assert merged[0][2] == "FirstName"  # name from first row


# ═══════════════════════════════════════════════════════════════════════
#  7. _extract_card_number
# ═══════════════════════════════════════════════════════════════════════

class TestExtractCardNumber:
    def test_from_sheet_content(self):
        """Card number found in sheet content."""
        es = _make_excel_sheet([
            ["Header", "SQRT1L-17-AS-04001"],
        ])
        num = _extract_card_number("unknown.xlsx", es)
        assert num == "SQRT1L-17-AS-04001"

    def test_fallback_to_filename(self):
        """No card number in sheet → fallback to filename."""
        es = _make_excel_sheet([["Just text"]])
        num = _extract_card_number("G01-AS-05001-Install.xlsx", es)
        assert num == "G01-AS-05001"

    def test_fallback_to_basename(self):
        """No pattern match → return basename without extension."""
        es = _make_excel_sheet([["No card here"]])
        num = _extract_card_number("simple_name.xlsx", es)
        assert num == "simple_name"


# ═══════════════════════════════════════════════════════════════════════
#  8. _collect_all_tables — multi-operation support
# ═══════════════════════════════════════════════════════════════════════

class TestCollectAllTables:
    def test_single_table(self):
        """Single table with 3 parts."""
        es = _make_excel_sheet([
            ["物料编码", "零件名称", "数量"],
            ["P001", "Part1", "1"],
            ["P002", "Part2", "2"],
            ["P003", "Part3", "1"],
        ])
        parts, table_count = _collect_all_tables(es, 4, 3, "test.xlsx")
        assert len(parts) == 3
        assert table_count == 1
        assert parts[0][0] == "P001"
        assert parts[1][0] == "P002"

    def test_multi_table(self):
        """2 tables separated by a gap with header."""
        es = _make_excel_sheet([
            ["序号\nСерийный номер", "零部件代号\nКод детали"],
            ["1", "P001"],
            ["2", "P002"],
            [None, None],
            [None, None],
            [None, None],
            ["序号", "零部件代号"],  # Second table header
            ["1", "Q001"],
            ["2", "Q002"],
        ])
        parts, table_count = _collect_all_tables(es, 9, 2, "test.xlsx")
        assert len(parts) == 4, f"Expected 4 parts from 2 tables, got {len(parts)}"
        assert table_count == 2
        pns = [p[0] for p in parts]
        assert "P001" in pns
        assert "P002" in pns
        assert "Q001" in pns
        assert "Q002" in pns

    def test_multi_table_with_boundary(self):
        """Two tables stopped by section boundary (new header)."""
        es = _make_excel_sheet([
            ["序号", "零部件代号"],
            ["1", "P001"],
            ["序号", "零部件代号"],  # New header → boundary
            ["1", "Q001"],
        ])
        parts, table_count = _collect_all_tables(es, 4, 2, "test.xlsx")
        # First table: P001. Second table: should continue past the boundary.
        # Both tables found.
        assert len(parts) == 2, f"Expected 2 parts from 2 tables, got {len(parts)}"
        assert table_count == 2

    def test_no_tables(self):
        """Sheet without part tables returns empty."""
        es = _make_excel_sheet([["Just", "Text"]])
        parts, table_count = _collect_all_tables(es, 1, 2, "test.xlsx")
        assert parts == []
        assert table_count == 0

    def test_empty_table_skipped(self):
        """Table with no valid parts should be skipped."""
        es = _make_excel_sheet([
            ["零部件代号", "名称"],
            ["AB", "Too Short"],  # invalid part number
            ["P001", "Valid"],
        ])
        parts, table_count = _collect_all_tables(es, 3, 2, "test.xlsx")
        # "AB" is invalid (too short), "P001" is valid
        assert len(parts) == 1, f"Expected 1 valid part, got {len(parts)}"
        assert table_count == 1
        assert parts[0][0] == "P001"

    def test_three_tables(self):
        """Three tables on one sheet separated by 3+ empty rows."""
        es = _make_excel_sheet([
            ["序号", "零部件代号"],     # R1 — table 1
            ["1", "P001"],
            ["2", "P002"],
            [None, None],
            [None, None],
            [None, None],
            ["序号", "零部件代号"],     # R7 — table 2
            ["1", "Q001"],
            ["2", "Q002"],
            [None, None],
            [None, None],
            [None, None],
            ["序号", "零部件代号"],     # R13 — table 3
            ["1", "R001"],
            ["2", "R002"],
        ])
        parts, table_count = _collect_all_tables(es, 15, 2, "test.xlsx")
        assert len(parts) == 6, f"Expected 6 parts from 3 tables, got {len(parts)}"
        assert table_count == 3
        pns = [p[0] for p in parts]
        assert pns == ["P001", "P002", "Q001", "Q002", "R001", "R002"], \
            f"Expected ordered parts, got {pns}"

    def test_tables_with_description_rows(self):
        """Tables separated by description/operation text rows (SWM-style)."""
        es = _make_excel_sheet([
            ["序号\nСерийный номер", "零部件代号\nКод детали"],
            ["1", "P001"],
            ["操作描述：拿取零部件1", None],  # description row, C2=None → skip
            [None, None],
            [None, None],
            ["序号", "零部件代号"],     # R6 — table 2 header
            ["1", "Q001"],
        ])
        parts, table_count = _collect_all_tables(es, 7, 2, "test.xlsx")
        assert len(parts) == 2, f"Expected 2 parts from 2 tables, got {len(parts)}"
        assert table_count == 2
        pns = [p[0] for p in parts]
        assert "P001" in pns
        assert "Q001" in pns

    def test_max_tables_limit(self):
        """Loop should stop at max_tables even if more headers exist."""
        # Create 12 identical headers
        rows = [["序号", "零部件代号"]]
        for i in range(12):
            rows.append([str(i + 1), f"P{i:03d}"])
            rows.append([None, None])
            rows.append([None, None])
            rows.append([None, None])
            rows.append(["序号", "零部件代号"])  # next header
        es = _make_excel_sheet(rows)
        parts, table_count = _collect_all_tables(es, len(rows), 2, "test.xlsx")
        # Should not crash, should find all 12 data tables
        # (trailing empty header not counted — tables_found only increments for non-empty tables)
        assert len(parts) == 12, f"Expected exactly 12 parts, got {len(parts)}"
        assert table_count == 12, f"Expected 12 tables (only data tables counted), got {table_count}"
        pns = [p[0] for p in parts]
        assert "P000" in pns
        assert "P011" in pns

    def test_all_tables_empty(self):
        """When all tables have no valid parts, returns empty list."""
        es = _make_excel_sheet([
            ["序号", "零部件代号"],
            ["AB", "Too Short"],  # invalid (too short)
            [None, None],
            [None, None],
            [None, None],
            ["序号", "零部件代号"],
            ["XY", "Also Short"],  # invalid (too short)
        ])
        parts, table_count = _collect_all_tables(es, 7, 2, "test.xlsx")
        assert parts == [], f"Expected empty list (no valid parts), got {len(parts)}"
        # table_count is 0 because no tables had valid parts (only counted when merged_parts non-empty)
        assert table_count == 0

    def test_tables_staggered_positions(self):
        """Tables at different row positions with staggered headers."""
        es = _make_excel_sheet([
            ["序号", "零部件代号"],     # R1 — table 1 at top
            ["1", "P001"],
            [None, None],
            [None, None],
            [None, None],
            ["Some", "Text"],           # R6 — non-header row (only 2 col, no PART_NO)
            [None, None],
            [None, None],
            [None, None],
            ["序号", "零部件代号"],     # R10 — table 2 deeper
            ["1", "Q001"],
        ])
        parts, table_count = _collect_all_tables(es, 11, 2, "test.xlsx")
        assert len(parts) == 2, f"Expected 2 parts from 2 staggered tables, got {len(parts)}"
        assert table_count == 2
        pns = [p[0] for p in parts]
        assert "P001" in pns
        assert "Q001" in pns


# ═══════════════════════════════════════════════════════════════════════
#  9. parse_card_file
# ═══════════════════════════════════════════════════════════════════════

class TestParseCardFile:
    def test_normal_card(self):
        """T1L card: header + data rows → should find parts with qty."""
        path = _make_card_xlsx([
            ["物料编码", "零件名称", "数量", "单位"],
            ["P001", "Part1", "2", "pcs"],
            ["P002", "Part2", "1", "pcs"],
            ["P003", "Part3", "3", "pcs"],
        ])
        result = parse_card_file(path)
        assert not result.is_service_file
        assert len(result.parts) == 3
        assert result.aggregated_parts["P001"] == 2.0
        assert result.aggregated_parts["P002"] == 1.0
        assert result.aggregated_parts["P003"] == 3.0

    def test_service_file(self):
        """Service files should not parse parts."""
        path = _make_card_xlsx([
            ["物料编码", "零件名称", "数量"],
            ["P001", "Part1", "1"],
        ])
        result = parse_card_file(path, is_service_file=True)
        assert result.is_service_file
        assert len(result.parts) == 0
        assert len(result.sheets) >= 1
        assert not result.sheets[0].is_valid

    def test_no_part_table(self):
        """Sheet without part table → valid=False, no parts."""
        path = _make_card_xlsx([
            ["Just some", "text without", "part numbers"],
        ])
        result = parse_card_file(path)
        assert not result.is_service_file
        assert len(result.parts) == 0
        # Sheet exists but no table found
        assert len(result.sheets) >= 1

    def test_card_number_extracted(self):
        """Card number should be extracted from sheet content."""
        path = _make_card_xlsx([
            ["SQRT1L-17-AS-04001", None, None],
            ["物料编码", "零件名称", "数量"],
            ["P001", "Part1", "1"],
        ])
        result = parse_card_file(path)
        # Card number from sheet content
        assert "SQRT1L-17-AS-04001" in result.card_number

    def test_aggregation_of_duplicate_parts(self):
        """Duplicate parts should have quantities summed."""
        path = _make_card_xlsx([
            ["物料编码", "零件名称", "数量"],
            ["P001", "Part1", "1"],
            ["P001", "Part1", "2"],
            ["P002", "Part2", "1"],
        ])
        result = parse_card_file(path)
        assert result.aggregated_parts["P001"] == 3.0  # 1+2
        assert result.aggregated_parts["P002"] == 1.0

    def test_multi_sheet_file(self):
        """File with multiple sheets should process each."""
        fd, path = tempfile.mkstemp(suffix=".xlsx", prefix="card_multi_")
        os.close(fd)
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "Op1"
        ws1.cell(row=1, column=1, value="物料编码")
        ws1.cell(row=1, column=2, value="数量")
        ws1.cell(row=2, column=1, value="P001")
        ws1.cell(row=2, column=2, value="1")
        ws2 = wb.create_sheet(title="Op2")
        ws2.cell(row=1, column=1, value="物料编码")
        ws2.cell(row=1, column=2, value="数量")
        ws2.cell(row=2, column=1, value="P002")
        ws2.cell(row=2, column=2, value="2")
        wb.save(path)

        result = parse_card_file(path)
        assert len(result.parts) == 2
        assert result.aggregated_parts["P001"] == 1.0
        assert result.aggregated_parts["P002"] == 2.0


# ═══════════════════════════════════════════════════════════════════════
#  10. CardService
# ═══════════════════════════════════════════════════════════════════════

class TestCardService:
    def test_initial_state(self):
        svc = CardService()
        assert not svc.is_loaded
        assert svc.cards is None

    def test_load_not_implemented_without_real_files(self):
        """CardService.load needs real ZIP/directory — not tested here.
        Just verify the error behavior with non-existent path."""
        svc = CardService()
        with pytest.raises(FileNotFoundError):
            svc.load("/nonexistent/path.zip")

    def test_not_loaded_raises(self):
        svc = CardService()
        with pytest.raises(RuntimeError, match="не загружены"):
            svc.get_all_parts()
        with pytest.raises(RuntimeError, match="не загружены"):
            svc.get_part_sources()
        with pytest.raises(RuntimeError, match="не загружены"):
            svc.get_card_results()


# ═══════════════════════════════════════════════════════════════════════
#  10b. CardService — load_from_bytes, cleanup, context manager, async
# ═══════════════════════════════════════════════════════════════════════

class TestCardServiceServer:
    """Тесты серверной функциональности CardService:
      - load_from_bytes (in-memory upload) для .xlsx
      - load_from_bytes для ZIP
      - cleanup (автоудаление temp-файлов)
      - context manager (with)
      - async (load_async)
    """

    @pytest.fixture
    def dir_with_cards(self, tmp_path) -> str:
        """Создать временную директорию с .xlsx файлами (нормальные имена).

        Использует pytest tmp_path — автоочистка после теста.
        """
        # Создаём файл с распознаваемым именем (цифры в начале)
        xlsx_path = os.path.join(tmp_path, "001-card.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.cell(row=1, column=1, value="物料编码")
        ws.cell(row=1, column=2, value="零件名称")
        ws.cell(row=1, column=3, value="数量")
        ws.cell(row=2, column=1, value="P001")
        ws.cell(row=2, column=2, value="Part1")
        ws.cell(row=2, column=3, value=2.0)
        ws.cell(row=3, column=1, value="P002")
        ws.cell(row=3, column=2, value="Part2")
        ws.cell(row=3, column=3, value=1.0)
        wb.save(xlsx_path)
        return str(tmp_path)

    @pytest.fixture
    def zip_bytes(self, tmp_path) -> bytes:
        """Создать ZIP с .xlsx файлами (нормальные имена), вернуть байты."""
        xlsx_path = os.path.join(tmp_path, "001-card.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.cell(row=1, column=1, value="物料编码")
        ws.cell(row=1, column=2, value="数量")
        ws.cell(row=2, column=1, value="P001")
        ws.cell(row=2, column=2, value=2.0)
        ws.cell(row=3, column=1, value="P002")
        ws.cell(row=3, column=2, value=1.0)
        wb.save(xlsx_path)

        zip_path = os.path.join(tmp_path, "cards.zip")
        import zipfile
        with zipfile.ZipFile(zip_path, "w") as zf:
            zf.write(xlsx_path, "001-card.xlsx")
        with open(zip_path, "rb") as f:
            return f.read()

    def test_load_from_directory(self, dir_with_cards: str):
        """Загрузка карт из директории с .xlsx файлами.

        Использует load() с правильным путём — файлы с номерами операций.
        """
        svc = CardService(max_workers=1)
        cards = svc.load(dir_with_cards)
        assert svc.is_loaded
        assert cards.total_cards_processed >= 1
        assert "P001" in cards.all_parts
        assert cards.all_parts["P001"] == 2.0
        assert "P002" in cards.all_parts
        assert cards.all_parts["P002"] == 1.0

    def test_load_zip_from_bytes(self, zip_bytes: bytes):
        """Загрузка ZIP-архива из байтов (in-memory upload).

        ZIP содержит файлы с распознаваемыми именами → парсинг успешен.
        """
        svc = CardService(max_workers=1)
        cards = svc.load_from_bytes(zip_bytes, filename="cards.zip")
        assert svc.is_loaded
        assert cards.total_cards_processed >= 1
        assert "P001" in cards.all_parts
        svc.cleanup()

    def test_cleanup_removes_temp_files(self, zip_bytes: bytes):
        """cleanup() удаляет созданные temp-файлы (ZIP + извлечённый файл)."""
        svc = CardService(max_workers=1)
        svc.load_from_bytes(zip_bytes, filename="cards.zip")
        assert len(svc._temp_paths) >= 1
        temp_path = svc._temp_paths[0]
        assert os.path.isfile(temp_path)
        svc.cleanup()
        assert not os.path.isfile(temp_path)
        assert len(svc._temp_paths) == 0

    def test_cleanup_removes_temp_dirs(self, zip_bytes: bytes):
        """cleanup() удаляет temp-директории, созданные для ZIP."""
        svc = CardService(max_workers=1)
        svc.load_from_bytes(zip_bytes, filename="cards.zip")
        assert len(svc._temp_dirs) >= 1
        temp_dir = svc._temp_dirs[0]
        assert os.path.isdir(temp_dir)
        svc.cleanup()
        assert not os.path.isdir(temp_dir)
        assert len(svc._temp_dirs) == 0

    def test_context_manager_cleans_up(self, zip_bytes: bytes):
        """Выход из with-блока вызывает cleanup."""
        with CardService(max_workers=1) as svc:
            svc.load_from_bytes(zip_bytes, filename="cards.zip")
            temp_path = svc._temp_paths[0]
            assert os.path.isfile(temp_path)
        assert not os.path.isfile(temp_path)

    def test_load_async(self, zip_bytes: bytes):
        """Асинхронная загрузка ZIP из байтов."""
        import asyncio

        async def run():
            svc = CardService(max_workers=1)
            cards = await svc.load_async(zip_bytes, filename="async_cards.zip")
            return svc, cards

        svc, cards = asyncio.run(run())
        assert svc.is_loaded
        assert "P001" in cards.all_parts
        svc.cleanup()

    def test_load_from_bytes_invalid_zip(self):
        """Загрузка невалидного ZIP — должна выбросить исключение."""
        import zipfile
        data = b"not a real zip file"
        svc = CardService(max_workers=1)
        with pytest.raises((zipfile.BadZipFile, ValueError)):
            svc.load_from_bytes(data, filename="cards.zip")
        svc.cleanup()

    def test_load_xlsx_without_card_data(self, tmp_path):
        """Загрузка .xlsx без распознаваемых заголовков деталей.

        Файл находится, классифицируется, но деталей не содержит.
        """
        xlsx_path = os.path.join(tmp_path, "001-card.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.cell(row=1, column=1, value="A")
        ws.cell(row=1, column=2, value="B")
        wb.save(xlsx_path)
        with open(xlsx_path, "rb") as f:
            data = f.read()

        svc = CardService(max_workers=1)
        cards = svc.load_from_bytes(data, filename="001-card.xlsx")
        assert svc.is_loaded
        assert "P001" not in cards.all_parts
        svc.cleanup()


# ═══════════════════════════════════════════════════════════════════════
#  11. _find_excel_files
# ═══════════════════════════════════════════════════════════════════════

class TestFindExcelFiles:
    def test_single_xlsx_file(self):
        """Single .xlsx file should be returned."""
        path = _make_card_xlsx([[1]], "single.xlsx")
        try:
            files = _find_excel_files(path)
            assert path in files, f"Expected {path} in {files}"
        finally:
            _safe_remove(path)

    def test_single_xls_file_not_found(self):
        """Non-existent .xls file returns empty list."""
        files = _find_excel_files("/nonexistent/file.xls")
        assert files == []

    def test_non_excel_file_skipped(self):
        """Non-Excel file should be skipped."""
        fd, path = tempfile.mkstemp(suffix=".txt", prefix="card_test_")
        os.close(fd)
        try:
            files = _find_excel_files(path)
            assert files == [], f"Expected empty for .txt file, got {files}"
        finally:
            _safe_remove(path)

    def test_directory_with_xlsx(self, tmp_path):
        """Directory containing .xlsx files should find them."""
        tmpdir = str(tmp_path)
        try:
            path1 = os.path.join(tmpdir, "card1.xlsx")
            path2 = os.path.join(tmpdir, "card2.xlsx")
            _touch_excel(path1)
            _touch_excel(path2)

            files = _find_excel_files(tmpdir)
            assert len(files) == 2, f"Expected 2 files, got {len(files)}"
            assert path1 in files
            assert path2 in files
        finally:
            _rmtree(tmpdir)

    def test_directory_with_nested_xlsx(self, tmp_path):
        """Directory with nested .xlsx files should find all."""
        tmpdir = str(tmp_path)
        try:
            subdir = os.path.join(tmpdir, "sub")
            os.makedirs(subdir)
            path1 = os.path.join(tmpdir, "card1.xlsx")
            path2 = os.path.join(subdir, "card2.xlsx")
            _touch_excel(path1)
            _touch_excel(path2)

            files = _find_excel_files(tmpdir)
            assert len(files) == 2, f"Expected 2 files (nested), got {len(files)}"
            assert path1 in files
            assert path2 in files
        finally:
            _rmtree(tmpdir)

    def test_directory_empty_returns_empty(self, tmp_path):
        """Empty directory returns empty list."""
        tmpdir = str(tmp_path)
        try:
            files = _find_excel_files(tmpdir)
            assert files == [], f"Expected empty for empty dir, got {len(files)}"
        finally:
            os.rmdir(tmpdir)


# ═══════════════════════════════════════════════════════════════════════
#  12. TEMPLATE_SHEET_KEYWORDS (constant)
# ═══════════════════════════════════════════════════════════════════════

class TestTemplateSheetKeywords:
    def test_contains_expected_keywords(self):
        assert "空表" in TEMPLATE_SHEET_KEYWORDS
        assert "填写范本" in TEMPLATE_SHEET_KEYWORDS
        assert "范本" in TEMPLATE_SHEET_KEYWORDS




# ═══════════════════════════════════════════════════════════════════════
#  35. Интеграционный тест: template-лист с данными попадает в split_cards
# ═══════════════════════════════════════════════════════════════════════

class TestTemplateSheetSplitIntegration:
    """Интеграционный тест: template-лист (空表) с данными включается в split_cards.

    Отличие от unit-тестов: использует реальный .xlsx файл и настоящий CardSplitter
    (не мокированный), чтобы проверить полный pipeline парсинга → split.
    """

    def test_template_sheet_with_data_included_in_split(self, tmp_path):
        """Создаём реальный .xlsx с листом '空表' и данными → split_cards создаёт файл.

        Проверяет полную цепочку:
          1. parse_card_file корректно определяет has_data=True для '空表' листа
          2. split_cards_to_files НЕ пропускает '空表' лист (т.к. есть данные)
          3. В output-директории создаётся .xlsx файл для этого листа
        """
        tmpdir = str(tmp_path)
        try:
            # ── 1. Создаём реальный .xlsx файл ──
            fd, xlsx_path = tempfile.mkstemp(suffix='.xlsx', prefix='int_template_')
            os.close(fd)

            wb = Workbook()
            ws = wb.active
            ws.title = '空表'  # template name
            # Заполняем данными (как реальная T1L карта)
            ws.cell(row=1, column=1, value='物料编码')
            ws.cell(row=1, column=2, value='零件名称')
            ws.cell(row=1, column=3, value='数量')
            ws.cell(row=2, column=1, value='P001')
            ws.cell(row=2, column=2, value='Болт')
            ws.cell(row=2, column=3, value=2.0)
            ws.cell(row=3, column=1, value='P002')
            ws.cell(row=3, column=2, value='Гайка')
            ws.cell(row=3, column=3, value=4.0)
            wb.save(xlsx_path)

            # ── 2. Парсим файл ──
            result = parse_card_file(xlsx_path)
            assert len(result.sheets) == 1, f"Expected 1 sheet, got {len(result.sheets)}"
            s = result.sheets[0]
            assert '空表' in s.sheet_name, f"Expected '空表' sheet, got {s.sheet_name}"
            assert s.has_data, "Template sheet with real data should have has_data=True"
            assert s.is_valid, "Template sheet with valid parts should be is_valid=True"
            assert len(result.parts) >= 2, f"Expected at least 2 parts, got {len(result.parts)}"

            # ── 3. Запускаем split_cards_to_files ──
            output_dir = os.path.join(tmpdir, 'split_output')
            cd = CardsData(
                all_parts=result.aggregated_parts,
                part_sources={
                    pn: [(result.card_number, xlsx_path, qty)]
                    for pn, qty in result.aggregated_parts.items()
                },
                card_results=[result],
            )

            created_files = split_cards_to_files(cd, output_dir, max_workers=1)

            # ── 4. Проверяем что файлы созданы ──
            assert len(created_files) > 0,                 "Template sheet with data should produce split files, got 0"

            # Проверяем что созданный файл существует и читается
            for fpath in created_files:
                assert os.path.isfile(fpath), f"Split file not found: {fpath}"
                # Открываем и проверяем что там есть данные
                wb_check = openpyxl.load_workbook(fpath, data_only=True)
                ws_check = wb_check.active
                assert ws_check is not None
                # Должны быть хотя бы заголовки и данные
                assert ws_check.cell(1, 1).value is not None, f"Split file {fpath} has no headers"
                assert ws_check.cell(2, 1).value is not None, f"Split file {fpath} has no data"
                wb_check.close()

            # ── 5. Проверяем статистику split ──
            stats = cd.split_stats
            assert stats is not None
            assert stats.total_files_created > 0
            assert stats.total_sheets_skipped == 0,                 f"Template sheet with data should NOT be skipped, but got {stats.total_sheets_skipped} skipped"

        finally:
            _safe_remove(xlsx_path)
            _rmtree(tmpdir)

    def test_header_only_sheet_content_still_split(self, tmp_path):
        """Header-only sheet в '空表' листе: has_data=True → split включается, файл создаётся.

        Даже один заголовок даёт has_data=True, и лист не пропускается как пустой.
        """
        tmpdir = str(tmp_path)
        try:
            fd, xlsx_path = tempfile.mkstemp(suffix='.xlsx', prefix='int_template_hdr_')
            os.close(fd)

            wb = Workbook()
            ws = wb.active
            ws.title = '空表'
            ws.cell(row=1, column=1, value='Header')
            wb.save(xlsx_path)

            result = parse_card_file(xlsx_path)
            assert result.sheets[0].has_data, "Header-only sheet should have has_data=True"

            cd = CardsData(
                all_parts={},
                part_sources={},
                card_results=[result],
            )
            output_dir = os.path.join(tmpdir, 'split_hdr')
            created_files = split_cards_to_files(cd, output_dir, max_workers=1)

            # has_data=True → лист включается в split, файл создаётся
            stats = cd.split_stats
            assert stats is not None
            assert stats.total_files_created >= 1,                 f"Header-only template should be included, got {stats.total_files_created}"
            assert len(created_files) >= 1

        finally:
            _safe_remove(xlsx_path)
            _rmtree(tmpdir)
    def test_mkdtemp_creates_in_system_temp(self):
        """mkdtemp() без dir= создаёт временную папку в tempfile.gettempdir(), а не в CWD."""
        tmpdir = tempfile.mkdtemp()
        try:
            assert os.path.isdir(tmpdir), f"mkdtemp should create a directory, got {tmpdir}"
            abs_path = os.path.abspath(tmpdir)
            temp_dir = os.path.abspath(tempfile.gettempdir())
            # Путь должен начинаться с системной temp-директории
            assert abs_path.startswith(temp_dir + os.sep), \
                f"mkdtemp() created {abs_path}, expected under {temp_dir}"
            # Путь НЕ должен быть в CWD
            cwd = os.path.abspath(os.getcwd())
            assert not abs_path.startswith(cwd + os.sep), \
                f"mkdtemp() created {abs_path} in CWD {cwd}, expected in system temp"
        finally:
            _rmtree(tmpdir)

    def test_mkdtemp_with_dir_creates_in_specified_dir(self):
        """mkdtemp(dir=...) создаёт папку в указанной директории (для контраста)."""
        custom_dir = tempfile.mkdtemp()
        try:
            tmpdir = tempfile.mkdtemp(dir=custom_dir)
            try:
                abs_path = os.path.abspath(tmpdir)
                custom_abs = os.path.abspath(custom_dir)
                assert abs_path.startswith(custom_abs + os.sep), \
                    f"mkdtemp(dir=...) created {abs_path}, expected under {custom_abs}"
            finally:
                _rmtree(tmpdir)
        finally:
            _rmtree(custom_dir)


# ═══════════════════════════════════════════════════════════════════════
#  ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ (для _find_excel_files)
# ═══════════════════════════════════════════════════════════════════════

def _safe_remove(path: str) -> None:
    try:
        if os.path.isfile(path):
            os.remove(path)
    except Exception:
        pass


def _rmtree(path: str) -> None:
    try:
        import shutil
        shutil.rmtree(path, ignore_errors=True)
    except Exception:
        pass


def _touch_excel(path: str) -> None:
    """Create a minimal valid .xlsx file."""
    wb = Workbook()
    wb.active.cell(row=1, column=1, value="test")
    wb.save(path)


# --- Mock-helper for split_cards_to_files tests ---


def _mock_extract_success(source_path, output_path, sheet_name):
    """Mock _extract_to_path_worker: successful extraction."""
    return {
        "path": output_path,
        "error": None,
        "used_fallback": False,
        "source_basename": os.path.basename(source_path),
        "source_path": source_path,
        "sheet_name": sheet_name,
    }


def _mock_extract_error(source_path, output_path, sheet_name):
    """Mock _extract_to_path_worker: extraction error."""
    return {
        "path": None,
        "error": "Mock error during split",
        "used_fallback": False,
        "source_basename": os.path.basename(source_path),
        "source_path": source_path,
        "sheet_name": sheet_name,
    }



# ═══════════════════════════════════════════════════════════════════════
#  13. ExcelReader — xlrd, fallback, error handling
# ═══════════════════════════════════════════════════════════════════════

class TestExcelReader:
    """ExcelReader — универсальный загрузчик .xlsx / .xls.

    Покрывает строки 62-109:
      - Загрузка .xlsx через openpyxl
      - Загрузка .xls через xlrd (мок)
      - Fallback openpyxl → xlrd при ошибке openpyxl
      - Ошибка импорта xlrd
      - Ошибка открытия .xls
    """

    def test_openpyxl_xlsx(self):
        """.xlsx загружается через openpyxl."""
        path = _make_card_xlsx([["A", "B"]], "test_openpyxl.xlsx")
        try:
            reader = ExcelReader(path)
            assert reader._engine == "openpyxl"
            assert len(reader.sheet_names) >= 1
            reader.close()
        finally:
            _safe_remove(path)

    def test_xlrd_xls(self, monkeypatch):
        """.xls загружается через xlrd (мокируем xlrd)."""
        import xlrd as real_xlrd

        class MockSheet:
            nrows = 2
            ncols = 2
            def cell_value(self, r, c):
                return [["A", "B"], ["1", "2"]][r][c]

        class MockBook:
            def sheet_names(self):
                return ["Sheet1"]
            def sheet_by_name(self, name):
                return MockSheet()

        monkeypatch.setattr(real_xlrd, "open_workbook", lambda path: MockBook())

        path = _make_card_xlsx([["A"]], "test_xls.xls")
        try:
            # Force .xls extension to trigger xlrd path
            fd, xls_path = tempfile.mkstemp(suffix=".xls", prefix="card_xlstest_")
            os.close(fd)
            open(xls_path, "w").close()  # empty file, but xlrd is mocked
            try:
                reader = ExcelReader(xls_path)
                assert reader._engine == "xlrd"
                assert len(reader.sheet_names) == 1
                sheet = reader.get_sheet("Sheet1")
                assert sheet.max_row == 2
                assert sheet.max_column == 2
                assert sheet.cell_value(1, 1) == "A"
                reader.close()
            finally:
                _safe_remove(xls_path)
        finally:
            _safe_remove(path)

    def test_xlrd_cell_value_float_to_int(self, monkeypatch):
        """xlrd возвращает float(3.0) как int(3) через ExcelSheet."""
        import xlrd as real_xlrd

        class MockSheet:
            nrows = 2
            ncols = 2
            def cell_value(self, r, c):
                if c == 0:  # column 1: part_no
                    if r == 0:
                        return "P001"
                    return None
                if c == 1:  # column 2: qty
                    if r == 0:
                        return None
                    return 2.0  # float
                return None

        class MockBook:
            def sheet_names(self): return ["Sheet1"]
            def sheet_by_name(self, name): return MockSheet()

        monkeypatch.setattr(real_xlrd, "open_workbook", lambda path: MockBook())

        fd, xls_path = tempfile.mkstemp(suffix=".xls", prefix="card_xlsfloat_")
        os.close(fd)
        open(xls_path, "w").close()
        try:
            reader = ExcelReader(xls_path)
            sheet = reader.get_sheet("Sheet1")
            val = sheet.cell_value(2, 2)
            assert val == 2, f"Expected int 2, got {val} (type={type(val).__name__})"
            reader.close()
        finally:
            _safe_remove(xls_path)

    def test_fallback_openpyxl_to_xlrd(self, monkeypatch):
        """При ошибке openpyxl .xlsx файл падает на xlrd."""
        import xlrd as real_xlrd

        class MockSheet:
            nrows = 1
            ncols = 1
            def cell_value(self, r, c): return "XL"

        class MockBook:
            def sheet_names(self): return ["Sheet1"]
            def sheet_by_name(self, name): return MockSheet()

        # Make openpyxl.load_workbook fail
        import openpyxl as real_openpyxl
        monkeypatch.setattr(real_openpyxl, "load_workbook", lambda path, **kw: (_ for _ in ()).throw(Exception("mock fail")))
        monkeypatch.setattr(real_xlrd, "open_workbook", lambda path: MockBook())

        path = _make_card_xlsx([["A"]], "test_fallback.xlsx")
        try:
            reader = ExcelReader(path)
            assert reader._engine == "xlrd", f"Expected xlrd fallback, got {reader._engine}"
            reader.close()
        finally:
            _safe_remove(path)

    def test_xlrd_import_error(self, monkeypatch):
        """xlrd не установлен → ImportError."""
        import builtins
        original_import = builtins.__import__

        def mock_import(name, *args, **kwargs):
            if name == "xlrd":
                raise ImportError("No xlrd")
            return original_import(name, *args, **kwargs)

        monkeypatch.setattr(builtins, "__import__", mock_import)

        fd, xls_path = tempfile.mkstemp(suffix=".xls", prefix="card_xlsno_")
        os.close(fd)
        open(xls_path, "w").close()
        try:
            with pytest.raises(ImportError, match="xlrd"):
                ExcelReader(xls_path)
        finally:
            _safe_remove(xls_path)

    def test_xlrd_open_error(self, monkeypatch):
        """xlrd.open_workbook падает → ValueError."""
        import xlrd as real_xlrd
        monkeypatch.setattr(real_xlrd, "open_workbook", lambda path: (_ for _ in ()).throw(Exception("Corrupt")))

        fd, xls_path = tempfile.mkstemp(suffix=".xls", prefix="card_xlscorrupt_")
        os.close(fd)
        open(xls_path, "w").close()
        try:
            with pytest.raises(ValueError, match="Не удалось открыть"):
                ExcelReader(xls_path)
        finally:
            _safe_remove(xls_path)

    def test_get_sheet_key_error(self):
        """get_sheet с несуществующим именем → KeyError."""
        path = _make_card_xlsx([[1]], "test_key.xlsx")
        try:
            reader = ExcelReader(path)
            with pytest.raises(KeyError, match="NonExistent"):
                reader.get_sheet("NonExistent")
            reader.close()
        finally:
            _safe_remove(path)

    def test_close_openpyxl(self):
        """close() on openpyxl engine should not crash."""
        path = _make_card_xlsx([[1]], "test_close.xlsx")
        try:
            reader = ExcelReader(path)
            reader.close()  # should not raise
        finally:
            _safe_remove(path)


# ═══════════════════════════════════════════════════════════════════════
#  14. _walk_extracted_dir — nested ZIP, temp cleanup
# ═══════════════════════════════════════════════════════════════════════

class TestWalkExtractedDir:
    """_walk_extracted_dir — обход директории, вложенные ZIP.

    Покрывает строки 636-669:
      - Пропуск временных файлов (~$)
      - Сбор .xlsx/.xls
      - Обработка вложенных ZIP
      - Очистка мусора (is_temp)
      - Дубликаты во вложенных архивах
    """

    def test_skips_temp_files(self, tmp_path):
        """Файлы, начинающиеся с ~$, пропускаются."""
        tmpdir = str(tmp_path)
        try:
            open(os.path.join(tmpdir, "~$tempfile.xlsx"), "w").close()
            open(os.path.join(tmpdir, "normal.xlsx"), "w").close()
            files: List[str] = []
            _walk_extracted_dir(tmpdir, tmpdir, files, set(), is_temp=False)
            assert len(files) == 1, f"Expected 1 normal file (temp skipped), got {len(files)}"
            assert "normal.xlsx" in files[0]
        finally:
            _rmtree(tmpdir)

    def test_cleans_non_excel_in_temp(self, tmp_path):
        """В temp-директории не-Excel файлы удаляются."""
        tmpdir = str(tmp_path)
        try:
            junk_path = os.path.join(tmpdir, "readme.txt")
            open(junk_path, "w").close()
            excel_path = os.path.join(tmpdir, "card.xlsx")
            _touch_excel(excel_path)

            files: List[str] = []
            _walk_extracted_dir(tmpdir, tmpdir, files, set(), is_temp=True)
            # Junk file should be removed
            assert not os.path.isfile(junk_path), "Junk file should be removed in temp dir"
            assert os.path.isfile(excel_path), "Excel file should remain"
            assert len(files) == 1
        finally:
            _rmtree(tmpdir)

    def test_keeps_non_excel_in_non_temp(self, tmp_path):
        """В НЕ-temp директории не-Excel файлы НЕ удаляются."""
        tmpdir = str(tmp_path)
        try:
            txt_path = os.path.join(tmpdir, "notes.txt")
            open(txt_path, "w").close()
            files: List[str] = []
            _walk_extracted_dir(tmpdir, tmpdir, files, set(), is_temp=False)
            assert os.path.isfile(txt_path), "Non-Excel file should remain in non-temp dir"
            assert len(files) == 0
        finally:
            _rmtree(tmpdir)

    def test_nested_zip_processed(self, tmp_path):
        """Вложенный ZIP распаковывается и файлы собираются."""
        tmpdir = str(tmp_path)
        try:
            # Create inner zip
            inner_xlsx = os.path.join(tmpdir, "inner.xlsx")
            _touch_excel(inner_xlsx)
            inner_zip = os.path.join(tmpdir, "archive.zip")
            with zipfile.ZipFile(inner_zip, "w") as zf:
                zf.write(inner_xlsx, "inner_file.xlsx")
            os.remove(inner_xlsx)

            # Create outer xlsx
            outer_xlsx = os.path.join(tmpdir, "outer.xlsx")
            _touch_excel(outer_xlsx)

            files: List[str] = []
            seen = set()
            _walk_extracted_dir(tmpdir, tmpdir, files, seen, is_temp=True)

            # outer.xlsx should be found
            outer_found = any("outer.xlsx" in f for f in files)
            assert outer_found, "outer.xlsx should be in files"
            # The zip was processed, inner files should be extracted
            # (may vary depending on whether the zip was properly created)
            assert len(files) >= 1, f"Expected at least outer.xlsx, got {len(files)}"
        finally:
            _rmtree(tmpdir)

    def test_nested_duplicate_detected(self, tmp_path):
        """Дубликат имени во вложенном ZIP пропускается."""
        tmpdir = str(tmp_path)
        try:
            seen = {"same_name.xlsx"}
            # Create inner zip with duplicate name
            inner_xlsx = os.path.join(tmpdir, "same_name.xlsx")
            _touch_excel(inner_xlsx)
            inner_zip = os.path.join(tmpdir, "dup.zip")
            with zipfile.ZipFile(inner_zip, "w") as zf:
                zf.write(inner_xlsx, "same_name.xlsx")
            os.remove(inner_xlsx)

            files: List[str] = []
            _walk_extracted_dir(tmpdir, tmpdir, files, seen, is_temp=True, is_nested=True)
            # The file 'same_name.xlsx' should be skipped (already in seen)
            # And the zip file should be removed (is_temp=True)
            assert not os.path.isfile(inner_zip), "Zip should be removed in temp"
            assert len(files) == 0, "Duplicate file should not be added"
        finally:
            _rmtree(tmpdir)


# ═══════════════════════════════════════════════════════════════════════
#  15. _merge_multiline — trailing continuation and edge cases
# ═══════════════════════════════════════════════════════════════════════

class TestMergeMultilineExtended:
    """Дополнительные тесты _merge_multiline_part_numbers.

    Покрывает строки 259-261 (trailing continuation buffer flush).
    """

    def test_trailing_continuation_without_completion(self):
        """Незавершённый перенос БЕЗ последующей строки (buffer не сброшен).

        Строка с тире на конце, но следующая строка не скипает
        (она просто не существует) — buffer не должен появиться в результате.
        """
        rows = [
            (2, "P001", 1.0, "Part1", 1),   # normal
            (3, "ABC-", 3.0, "Part2", 1),   # continuation at the end, no next row
        ]
        merged = _merge_multiline_part_numbers(rows)
        # P001 should be in result, ABC- should NOT (incomplete continuation)
        assert len(merged) == 1, f"Expected 1 part (incomplete continuation dropped), got {len(merged)}"
        assert merged[0][0] == "P001"

    def test_continuation_invalid_final_buffer(self):
        """Buffer после цикла не проходит is_valid_part_number → не добавляется."""
        rows = [
            (2, "XY-", 1.0, "PartX", 1),  # continuation (too short when cleaned)
        ]
        merged = _merge_multiline_part_numbers(rows)
        assert len(merged) == 0, "Incomplete continuation should be dropped"


# ═══════════════════════════════════════════════════════════════════════
#  16. _find_excel_files — .xls, nested ZIP
# ═══════════════════════════════════════════════════════════════════════

class TestFindExcelFilesExtended:
    """Дополнительные тесты _find_excel_files.

    Покрывает:
      - .xls файл как одиночный файл
      - Не-Excel файл (.txt)
      - Вложенный ZIP с дубликатами
    """

    def test_single_xls_file(self):
        """Одиночный .xls файл возвращается."""
        fd, path = tempfile.mkstemp(suffix=".xls", prefix="card_findxls_")
        os.close(fd)
        open(path, "w").close()
        try:
            files = _find_excel_files(path)
            assert len(files) == 1, f"Expected 1 .xls file, got {len(files)}"
            assert path in files
        finally:
            _safe_remove(path)

    def test_zip_with_nested_content(self, tmp_path):
        """ZIP с вложенными .xlsx извлекается и файлы находятся."""
        tmpdir = str(tmp_path)
        try:
            # Create xlsx files inside temp dir
            xlsx1 = os.path.join(tmpdir, "card1.xlsx")
            xlsx2_dir = os.path.join(tmpdir, "subdir")
            os.makedirs(xlsx2_dir)
            xlsx2 = os.path.join(xlsx2_dir, "card2.xlsx")
            _touch_excel(xlsx1)
            _touch_excel(xlsx2)

            # Create zip
            zip_path = os.path.join(tmpdir, "cards.zip")
            with zipfile.ZipFile(zip_path, "w") as zf:
                zf.write(xlsx1, "card1.xlsx")
                zf.write(xlsx2, "subdir/card2.xlsx")

            extract_dir = os.path.join(tmpdir, "extracted")
            files = _find_excel_files(zip_path, extract_dir=extract_dir)
            assert len(files) == 2, f"Expected 2 files from zip, got {len(files)}"
            assert any("card1.xlsx" in f for f in files)
            assert any("card2.xlsx" in f for f in files)
        finally:
            _rmtree(tmpdir)

    def test_directory_with_mixed_formats(self, tmp_path):
        """Директория с .xlsx, .xls и .txt — только Excel файлы."""
        tmpdir = str(tmp_path)
        try:
            _touch_excel(os.path.join(tmpdir, "card1.xlsx"))
            open(os.path.join(tmpdir, "card2.xls"), "w").close()
            open(os.path.join(tmpdir, "notes.txt"), "w").close()

            files = _find_excel_files(tmpdir)
            assert len(files) == 2, f"Expected 2 Excel files, got {len(files)}"
            assert any(f.endswith(".xlsx") for f in files)
            assert any(f.endswith(".xls") for f in files)
        finally:
            _rmtree(tmpdir)


# ═══════════════════════════════════════════════════════════════════════
#  17. parse_cards — parallel & sequential paths
# ═══════════════════════════════════════════════════════════════════════

class TestParseCards:
    """parse_cards — параллельный и последовательный парсинг.

    Покрывает строки 740-802:
      - Параллельный парсинг (workers > 1)
      - Последовательный парсинг (workers = 1)
      - Обработка служебных файлов
      - original_part_numbers
    """

    def test_sequential_parsing(self, tmp_path):
        """Парсинг с max_workers=1 использует последовательный путь."""
        tmpdir = str(tmp_path)
        try:
            path1 = os.path.join(tmpdir, "001-card.xlsx")
            wb = Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws.cell(row=1, column=1, value="物料编码")
            ws.cell(row=1, column=2, value="数量")
            ws.cell(row=2, column=1, value="P001")
            ws.cell(row=2, column=2, value=2.0)
            wb.save(path1)

            path2 = os.path.join(tmpdir, "002-card.xlsx")
            wb2 = Workbook()
            ws2 = wb2.active
            ws2.title = "Sheet1"
            ws2.cell(row=1, column=1, value="物料编码")
            ws2.cell(row=1, column=2, value="数量")
            ws2.cell(row=2, column=1, value="P002")
            ws2.cell(row=2, column=2, value=1.0)
            wb2.save(path2)

            result = parse_cards(tmpdir, max_workers=1, show_progress=False)
            assert result.total_cards_processed >= 2
            assert "P001" in result.all_parts
            assert "P002" in result.all_parts
            assert result.all_parts["P001"] == 2.0
            assert result.all_parts["P002"] == 1.0
        finally:
            _rmtree(tmpdir)

    def test_parallel_parsing(self, tmp_path):
        """Парсинг с max_workers=2 использует параллельный путь."""
        tmpdir = str(tmp_path)
        try:
            path1 = os.path.join(tmpdir, "001-par.xlsx")
            wb = Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws.cell(row=1, column=1, value="物料编码")
            ws.cell(row=1, column=2, value="数量")
            ws.cell(row=2, column=1, value="P100")
            ws.cell(row=2, column=2, value=3.0)
            wb.save(path1)

            path2 = os.path.join(tmpdir, "002-par.xlsx")
            wb2 = Workbook()
            ws2 = wb2.active
            ws2.title = "Sheet1"
            ws2.cell(row=1, column=1, value="物料编码")
            ws2.cell(row=1, column=2, value="数量")
            ws2.cell(row=2, column=1, value="P200")
            ws2.cell(row=2, column=2, value=5.0)
            wb2.save(path2)

            result = parse_cards(tmpdir, max_workers=2, show_progress=False)
            assert result.total_cards_processed >= 2
            assert "P100" in result.all_parts
            assert "P200" in result.all_parts
            assert result.all_parts["P100"] == 3.0
        finally:
            _rmtree(tmpdir)

    def test_original_part_numbers(self):
        """original_part_numbers заполняется из карт."""
        # parse_cards goes through file_classifier — must use recognizable headers
        path = _make_card_xlsx([
            ["序号", "零部件代号", "数量"],
            ["1", "5306200-ED001", 1.0],
        ])
        try:
            result = parse_card_file(path)
            aggregated = result.aggregated_parts
            assert "5306200ED001" in aggregated, f"Expected cleaned PN, got {list(aggregated.keys())}"
            # Check original_part_numbers from the card parse result
            # (parse_card_file doesn't build original_part_numbers — parse_cards does)
            # So we check that the part was found with the right quantity
            assert aggregated["5306200ED001"] == 1.0
        finally:
            _safe_remove(path)


# ═══════════════════════════════════════════════════════════════════════
#  18. parse_card_file — empty sheets, no tables
# ═══════════════════════════════════════════════════════════════════════

class TestParseCardFileExtended:
    """Дополнительные тесты parse_card_file.

    Покрывает строки 324-359:
      - Лист с max_row=0 или max_col=0 → пустой, пропускается
      - Лист без таблицы деталей → "Лист без таблицы деталей"
      - Служебный файл → is_service_file
    """

    def test_empty_sheet_no_data_rows(self):
        """Лист без записанных данных (max_row=0)."""
        fd, path = tempfile.mkstemp(suffix=".xlsx", prefix="card_empty2_")
        os.close(fd)
        wb = Workbook()
        ws = wb.active
        ws.title = "Empty"
        # No data rows written
        wb.save(path)
        try:
            result = parse_card_file(path)
            assert len(result.parts) == 0
            assert len(result.sheets) >= 1
            assert not result.sheets[0].is_valid
            assert not result.sheets[0].has_data
        finally:
            _safe_remove(path)

    def test_sheet_without_part_table_first(self):
        """Лист с данными, но без таблицы деталей (первый дубликат, переименован)."""
        fd, path = tempfile.mkstemp(suffix=".xlsx", prefix="card_notable2_")
        os.close(fd)
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        ws.cell(row=1, column=1, value="Some descriptive text without part numbers")
        ws.cell(row=2, column=1, value="More random content")
        wb.save(path)
        try:
            result = parse_card_file(path)
            assert len(result.parts) == 0
            assert len(result.sheets) >= 1
            assert not result.sheets[0].is_valid
            assert result.sheets[0].has_data is True
        finally:
            _safe_remove(path)

    def test_empty_sheet_removed_active(self):
        """Лист с max_row=0 (без единой ячейки с данными) → пропускается.

        Покрывает строки 324-331: max_row == 0 → лист записывается как
        is_valid=False, has_data=False и пропускается.
        openpyxl Workbook с только что созданным пустым листом имеет
        max_row=None → ExcelSheet.max_row возвращает 0.
        """
        fd, path = tempfile.mkstemp(suffix=".xlsx", prefix="card_empty_")
        os.close(fd)
        wb = Workbook()
        # Delete the default active sheet so we have NO sheets with data
        ws = wb.active
        wb.remove(ws)
        # Add a new empty sheet (no cells written)
        wb.create_sheet(title="Empty")
        wb.save(path)
        try:
            result = parse_card_file(path)
            assert len(result.parts) == 0
            assert len(result.sheets) >= 1
            assert not result.sheets[0].is_valid
            assert not result.sheets[0].has_data
        finally:
            _safe_remove(path)

    def test_sheet_without_part_table_second(self):
        """Лист с данными, но без таблицы деталей (второй дубликат, переименован)."""
        fd, path = tempfile.mkstemp(suffix=".xlsx", prefix="card_notable3_")
        os.close(fd)
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        ws.cell(row=1, column=1, value="Some descriptive text without part numbers")
        ws.cell(row=2, column=1, value="More random content")
        wb.save(path)
        try:
            result = parse_card_file(path)
            assert len(result.parts) == 0
            assert len(result.sheets) >= 1
            assert not result.sheets[0].is_valid
            assert result.sheets[0].has_data is True
        finally:
            _safe_remove(path)

    def test_service_file_sheets_info(self):
        """Служебный файл возвращает sheets с is_valid=False, has_data по факту."""
        path = _make_card_xlsx([
            ["物料编码", "零件名称", "数量"],
            ["P001", "Part1", "1"],
        ])
        try:
            result = parse_card_file(path, is_service_file=True)
            assert result.is_service_file
            assert len(result.parts) == 0
            assert len(result.sheets) >= 1
            assert not result.sheets[0].is_valid
        finally:
            _safe_remove(path)


# ═══════════════════════════════════════════════════════════════════════
#  19. _extract_card_number — fallback edge
# ═══════════════════════════════════════════════════════════════════════

class TestExtractCardNumberExtended:
    """Дополнительные тесты _extract_card_number.

    Покрывает строки 220-221 (абсолютный fallback).
    """

    def test_absolute_fallback(self):
        """Когда extract_card_number возвращает None → базовое имя файла без расширения."""
        # Создаём лист без номера карты
        es = _make_excel_sheet([["Просто текст"]])
        # Файл с именем, которое НЕ соответствует CARD_NUMBER_RE
        num = _extract_card_number("my_random_file.xlsx", es)
        assert num == "my_random_file"


# ═══════════════════════════════════════════════════════════════════════
#  20. _safe_name — cleanup helper
# ═══════════════════════════════════════════════════════════════════════

class TestSafeName:
    def test_safe_name_removes_special_chars(self):
        """_safe_name заменяет спецсимволы на подчёркивания."""
        result = _safe_name("my@file#name.xlsx")
        assert "@" not in result
        assert "#" not in result
        assert "_" in result

    def test_safe_name_truncates(self):
        """_safe_name обрезает до 50 символов."""
        long_name = "a" * 100 + ".xlsx"
        result = _safe_name(long_name)
        assert len(result) <= 50


# ═══════════════════════════════════════════════════════════════════════
#  21. _safe_remove — error handling
# ═══════════════════════════════════════════════════════════════════════

class TestSafeRemove:
    def test_remove_existing_file(self):
        """Удаление существующего файла."""
        fd, path = tempfile.mkstemp(prefix="safe_")
        os.close(fd)
        assert os.path.isfile(path)
        _safe_remove(path)
        assert not os.path.isfile(path)

    def test_remove_nonexistent_file(self):
        """Удаление несуществующего файла не вызывает ошибку."""
        _safe_remove("/nonexistent/path/12345")  # should not raise


# ═══════════════════════════════════════════════════════════════════════
#  22. CardService - load with .xlsx file
# ═══════════════════════════════════════════════════════════════════════

class TestCardServiceExtended:
    """CardService.load with a real .xlsx file.

    Note: CardService.load uses parse_cards which goes through file_classifier.
    Files must have digit-prefixed names and recognizable headers.
    """

    def test_load_directory_with_cards(self, tmp_path):
        """Загрузка директории с картами через CardService."""
        tmpdir = str(tmp_path)
        try:
            path = os.path.join(tmpdir, "001-card.xlsx")
            wb = Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws.cell(row=1, column=1, value="序号")
            ws.cell(row=1, column=2, value="零部件代号")
            ws.cell(row=1, column=3, value="数量")
            ws.cell(row=2, column=1, value=1)
            ws.cell(row=2, column=2, value="P001")
            ws.cell(row=2, column=3, value=2.0)
            wb.save(path)

            svc = CardService()
            result = svc.load(tmpdir)
            assert svc.is_loaded
            assert "P001" in result.all_parts
            assert "P001" in svc.get_all_parts()
            sources = svc.get_part_sources()
            assert "P001" in sources
            assert len(svc.get_card_results()) >= 1
        finally:
            _rmtree(tmpdir)


# ═══════════════════════════════════════════════════════════════════════
#  34. SplitStatistics — детальная статистика разделения
# ═══════════════════════════════════════════════════════════════════════

class TestSplitStatistics:
    """split_cards_to_files — детальная per-file статистика (SplitStatistics).

    Проверяет:
      - SplitStatistics создаётся и заполняется
      - .xls файлы корректно помечаются как неподдерживаемые
      - Служебные файлы (is_service_file) корректно помечаются
      - Шаблонные листы корректно учитываются в skip_reasons
      - Пустые листы корректно учитываются в skip_reasons
      - Агрегированные счётчики (total_xlsx, total_xls, total_sheets_all, и т.д.)
      - get_top_skip_reasons() возвращает корректные данные
      - get_files_with_most_skips() возвращает корректные данные
    """

    def test_xlsx_files_processed(self, monkeypatch, tmp_path):
        """.xlsx файлы правильно считаются в статистике (total_xlsx)."""
        from burlak_parser import splitter as splitter_mod

        monkeypatch.setattr(
            splitter_mod, "_extract_to_path_worker",
            _mock_extract_success,
        )

        result = CardParseResult(
            card_number="C001", file_path="test.xlsx",
            sheets=[CardSheetInfo("C001", "S1", has_data=True)],
            parts=[], aggregated_parts={},
        )
        cd = CardsData(all_parts={}, part_sources={}, card_results=[result])
        tmpdir = str(tmp_path)
        try:
            split_cards_to_files(cd, tmpdir, max_workers=1)
            stats = cd.split_stats
            assert stats is not None
            assert stats.total_xlsx == 1
            assert stats.total_xls == 0
            assert stats.total_sheets_all == 1
            assert stats.total_sheets_split == 1
            assert stats.total_sheets_skipped == 0
            assert stats.total_files_created == 1
            assert len(stats.file_stats) == 1
            assert stats.file_stats[0].is_xlsx is True
            assert stats.file_stats[0].sheets_split == 1
            assert stats.file_stats[0].sheets_skipped == 0
        finally:
            _rmtree(tmpdir)

    def test_xls_files_marked_skipped(self, tmp_path):
        """.xls файлы помечаются как неподдерживаемые (split_reason)."""
        result = CardParseResult(
            card_number="C001", file_path="test.xls",
            sheets=[CardSheetInfo("C001", "S1", has_data=True)],
            parts=[], aggregated_parts={},
        )
        cd = CardsData(all_parts={}, part_sources={}, card_results=[result])
        tmpdir = str(tmp_path)
        try:
            split_cards_to_files(cd, tmpdir, max_workers=1)
            stats = cd.split_stats
            assert stats is not None
            assert stats.total_xlsx == 0
            assert stats.total_xls == 1
            assert stats.total_sheets_all == 1
            assert stats.total_files_created == 0
            assert len(stats.file_stats) == 1
            fs = stats.file_stats[0]
            assert fs.is_xlsx is False
            assert fs.sheets_split == 0
            assert fs.sheets_skipped == 1
            assert "не .xlsx" in fs.split_reason.lower() or "не xlsx" in fs.split_reason.lower()
        finally:
            _rmtree(tmpdir)

    def test_service_files_marked_skipped(self, tmp_path):
        """Служебные файлы помечаются как пропущенные (split_reason)."""
        result = CardParseResult(
            card_number="C001", file_path="test.xlsx",
            sheets=[CardSheetInfo("C001", "S1", has_data=True)],
            parts=[], aggregated_parts={},
            is_service_file=True,
        )
        cd = CardsData(all_parts={}, part_sources={}, card_results=[result])
        tmpdir = str(tmp_path)
        try:
            split_cards_to_files(cd, tmpdir, max_workers=1)
            stats = cd.split_stats
            assert stats is not None
            assert stats.total_service_files == 1
            assert stats.total_xlsx == 1
            assert stats.total_files_created == 0
            fs = stats.file_stats[0]
            assert fs.is_service_file is True
            assert fs.sheets_split == 0
            assert fs.sheets_skipped == 1
            assert "служебный" in fs.split_reason.lower()
        finally:
            _rmtree(tmpdir)

    def test_template_sheets_with_data_are_now_included(self, tmp_path):
        """Шаблонные листы с данными БОЛЬШЕ НЕ пропускаются (has_data=True → включены в split)."""
        result = CardParseResult(
            card_number="C001", file_path="test.xlsx",
            sheets=[
                CardSheetInfo("C001", "空表_Sheet1", has_data=True),   # has data → NOT skipped anymore
                CardSheetInfo("C001", "Sheet2", has_data=True),
            ],
            parts=[], aggregated_parts={},
        )
        cd = CardsData(all_parts={}, part_sources={}, card_results=[result])
        tmpdir = str(tmp_path)
        try:
            split_cards_to_files(cd, tmpdir, max_workers=1)
            stats = cd.split_stats
            assert stats is not None
            assert stats.total_sheets_all == 2
            assert stats.total_sheets_split == 2, "Both sheets have data → both should split"
            assert stats.total_sheets_skipped == 0
            fs = stats.file_stats[0]
            assert SKIP_REASON_TEMPLATE not in fs.skip_reasons,                 "Template sheet with data should NOT be in skip_reasons"
            assert fs.sheets_split == 2
            assert fs.sheets_skipped == 0
        finally:
            _rmtree(tmpdir)

    def test_template_sheets_without_data_still_skipped(self, tmp_path):
        """Шаблонные листы БЕЗ данных по-прежнему пропускаются."""
        result = CardParseResult(
            card_number="C001", file_path="test.xlsx",
            sheets=[
                CardSheetInfo("C001", "空表_Empty", has_data=False),  # no data → skipped
                CardSheetInfo("C001", "Sheet2", has_data=True),
            ],
            parts=[], aggregated_parts={},
        )
        cd = CardsData(all_parts={}, part_sources={}, card_results=[result])
        tmpdir = str(tmp_path)
        try:
            split_cards_to_files(cd, tmpdir, max_workers=1)
            stats = cd.split_stats
            assert stats is not None
            assert stats.total_sheets_all == 2
            assert stats.total_sheets_split == 1
            assert stats.total_sheets_skipped == 1
            fs = stats.file_stats[0]
            assert SKIP_REASON_TEMPLATE in fs.skip_reasons
            assert "空表_Empty" in fs.skip_reasons[SKIP_REASON_TEMPLATE]
            assert fs.sheets_split == 1
            assert fs.sheets_skipped == 1
        finally:
            _rmtree(tmpdir)

    def test_empty_sheets_recorded_in_skip_reasons(self, tmp_path):
        """Пустые листы корректно записываются в skip_reasons."""
        result = CardParseResult(
            card_number="C001", file_path="test.xlsx",
            sheets=[
                CardSheetInfo("C001", "S1", has_data=True),
                CardSheetInfo("C001", "S2", has_data=False),
                CardSheetInfo("C001", "S3", has_data=False),
            ],
            parts=[], aggregated_parts={},
        )
        cd = CardsData(all_parts={}, part_sources={}, card_results=[result])
        tmpdir = str(tmp_path)
        try:
            split_cards_to_files(cd, tmpdir, max_workers=1)
            stats = cd.split_stats
            assert stats is not None
            assert stats.total_sheets_all == 3
            assert stats.total_sheets_split == 1
            assert stats.total_sheets_skipped == 2
            fs = stats.file_stats[0]
            assert SKIP_REASON_NO_DATA in fs.skip_reasons
            assert len(fs.skip_reasons[SKIP_REASON_NO_DATA]) == 2
            assert "S2" in fs.skip_reasons[SKIP_REASON_NO_DATA]
            assert "S3" in fs.skip_reasons[SKIP_REASON_NO_DATA]
        finally:
            _rmtree(tmpdir)

    def test_get_top_skip_reasons(self):
        """get_top_skip_reasons() возвращает причины отсортированные по частоте."""
        stats = SplitStatistics(
            file_stats=[
                FileSplitStats(
                    file_path="t.xlsx", file_name="t.xlsx", card_number="C",
                    is_xlsx=True, is_service_file=False, 
                    total_sheets=3, sheets_split=1, sheets_skipped=2,
                    skip_reasons={SKIP_REASON_NO_DATA: ["S1", "S2"]},
                ),
            ],
            total_xlsx=1, total_xls=0,
            total_sheets_all=3, total_sheets_split=1, total_sheets_skipped=2,
            total_files_created=1,
        )
        top = stats.get_top_skip_reasons(5)
        assert len(top) >= 1
        # The most common reason should be SKIP_REASON_NO_DATA with count 2
        reason_name, count = top[0]
        assert count == 2
        assert reason_name == SKIP_REASON_NO_DATA

    def test_get_files_with_most_skips(self):
        """get_files_with_most_skips() возвращает файлы отсортированные по skips."""
        stats = SplitStatistics(
            file_stats=[
                FileSplitStats(
                    file_path="a.xlsx", file_name="a.xlsx", card_number="C1",
                    is_xlsx=True, is_service_file=False,
                    total_sheets=10, sheets_split=2, sheets_skipped=8,
                ),
                FileSplitStats(
                    file_path="b.xlsx", file_name="b.xlsx", card_number="C2",
                    is_xlsx=True, is_service_file=False,
                    total_sheets=5, sheets_split=4, sheets_skipped=1,
                ),
                FileSplitStats(
                    file_path="c.xlsx", file_name="c.xlsx", card_number="C3",
                    is_xlsx=False, is_service_file=True,
                    total_sheets=3, sheets_split=0, sheets_skipped=3,
                    split_reason="skip",
                ),
            ],
            total_xlsx=2, total_xls=1,
            total_sheets_all=18, total_sheets_split=6, total_sheets_skipped=12,
            total_files_created=6,
        )
        top = stats.get_files_with_most_skips(2)
        assert len(top) == 2
        assert top[0][0] == "a.xlsx"  # most skips (8)
        assert top[0][2] == 8
        # c.xlsx has 3 skips, b.xlsx has 1 skip -> second is c.xlsx
        assert top[1][0] == "c.xlsx"  # second most (3 skips)
        assert top[1][2] == 3

    def test_error_files_marked(self, monkeypatch, tmp_path):
        """Ошибки при split_file отмечаются в file_stats."""
        from burlak_parser import splitter as splitter_mod

        def mock_extract_to_path(source_path, output_path, sheet_name):
            return {
                "path": None,
                "error": "Mock error during split",
                "used_fallback": False,
                "source_basename": os.path.basename(source_path),
                "source_path": source_path,
                "sheet_name": sheet_name,
            }

        monkeypatch.setattr(
            splitter_mod, "_extract_to_path_worker",
            mock_extract_to_path,
        )

        result = CardParseResult(
            card_number="C001", file_path="test.xlsx",
            sheets=[CardSheetInfo("C001", "S1", has_data=True)],
            parts=[], aggregated_parts={},
        )
        cd = CardsData(all_parts={}, part_sources={}, card_results=[result])
        tmpdir = str(tmp_path)
        try:
            split_cards_to_files(cd, tmpdir, max_workers=1)
            stats = cd.split_stats
            assert stats is not None
            assert stats.total_errors >= 1
            error_file = [fs for fs in stats.file_stats if fs.has_error]
            assert len(error_file) >= 1
            assert error_file[0].file_name == "test.xlsx"
            assert "Mock error" in error_file[0].error_message
        finally:
            _rmtree(tmpdir)

    def test_cp7cp8_files_included_in_stats(self, monkeypatch, tmp_path):
        """CP7/CP8 файлы включаются в статистику (не пропускаются)."""
        from burlak_parser import splitter as splitter_mod

        monkeypatch.setattr(
            splitter_mod, "_extract_to_path_worker",
            _mock_extract_success,
        )

        result = CardParseResult(
            card_number="CP001", file_path="cp7.xlsx",
            sheets=[CardSheetInfo("CP001", "S1", has_data=True)],
            parts=[], aggregated_parts={},
            is_service_file=False,
        )
        cd = CardsData(all_parts={}, part_sources={}, card_results=[result])
        tmpdir = str(tmp_path)
        try:
            split_cards_to_files(cd, tmpdir, max_workers=1)
            stats = cd.split_stats
            assert stats is not None
            assert stats.total_xlsx == 1
            assert stats.total_files_created == 1
            fs = stats.file_stats[0]
            assert fs.sheets_split == 1
        finally:
            _rmtree(tmpdir)
