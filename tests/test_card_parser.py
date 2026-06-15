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
import tempfile
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
    ExcelSheet,
    TEMPLATE_SHEET_KEYWORDS,
    _check_sheet_has_data,
    _collect_all_tables,
    _collect_raw_rows,
    _extract_card_number,
    _find_excel_files,
    _merge_multiline_part_numbers,
    parse_card_file,
)
from burlak_parser.heuristic_analyzer import (
    HeuristicAnalyzer,
    clean_part_number,
    is_valid_part_number,
)


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
    fd, path = tempfile.mkstemp(suffix=".xlsx", prefix="card_test_", dir="/home/rpaup/projects/burlak")
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
        assert result.is_final_check is False
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
            is_final_check=True,
        )
        assert len(result.parts) == 2
        assert result.is_service_file is True
        assert result.is_final_check is True


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
        assert cd.corrupted_files is None


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
        """T1L card: single table with 3 parts."""
        es = _make_excel_sheet([
            ["物料编码", "零件名称", "数量"],
            ["P001", "Part1", "1"],
            ["P002", "Part2", "2"],
            ["P003", "Part3", "1"],
        ])
        parts = _collect_all_tables(es, 4, 3, "test.xlsx")
        assert len(parts) == 3
        assert parts[0][0] == "P001"
        assert parts[1][0] == "P002"

    def test_multi_table(self):
        """SWM card: 2 tables separated by a gap with header."""
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
        parts = _collect_all_tables(es, 9, 2, "test.xlsx")
        assert len(parts) == 4, f"Expected 4 parts from 2 tables, got {len(parts)}"
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
        parts = _collect_all_tables(es, 4, 2, "test.xlsx")
        # First table: P001. Second table: should continue past the boundary.
        # The boundary detection stops at R3, then _collect_all_tables starts
        # from R4 (last_data_row=2, start_search=3+1=4... wait)
        # _collect_raw_rows for first table: header=R1, scans R2.
        # R2 has "1", "P001" → collected.
        # R3 is new header → boundary stop. raw_rows=[(2, "P001", ...)]
        # last_data_row=2, start_search=3
        # find_part_table(ws, start_row=3) → finds R3 (header)
        # header_row=3 >= start_search=3 → OK
        # _collect_raw_rows for second table: header=R3, scans R4
        # R4 has "1", "Q001" → collected.
        # Both tables found.
        assert len(parts) == 2, f"Expected 2 parts from 2 tables, got {len(parts)}"

    def test_no_tables(self):
        """Sheet without part tables returns empty."""
        es = _make_excel_sheet([["Just", "Text"]])
        parts = _collect_all_tables(es, 1, 2, "test.xlsx")
        assert parts == []

    def test_empty_table_skipped(self):
        """Table with no valid parts should be skipped."""
        es = _make_excel_sheet([
            ["零部件代号", "名称"],
            ["AB", "Too Short"],  # invalid part number
            ["P001", "Valid"],
        ])
        parts = _collect_all_tables(es, 3, 2, "test.xlsx")
        # "AB" is invalid (too short), "P001" is valid
        assert len(parts) == 1, f"Expected 1 valid part, got {len(parts)}"
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
        parts = _collect_all_tables(es, 15, 2, "test.xlsx")
        assert len(parts) == 6, f"Expected 6 parts from 3 tables, got {len(parts)}"
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
        parts = _collect_all_tables(es, 7, 2, "test.xlsx")
        assert len(parts) == 2, f"Expected 2 parts from 2 tables, got {len(parts)}"
        pns = [p[0] for p in parts]
        assert "P001" in pns
        assert "Q001" in pns

    def test_max_tables_limit(self):
        """Loop should stop at max_tables=10 even if more headers exist."""
        # Create 12 identical headers
        rows = [["序号", "零部件代号"]]
        for i in range(12):
            rows.append([str(i + 1), f"P{i:03d}"])
            rows.append([None, None])
            rows.append([None, None])
            rows.append([None, None])
            rows.append(["序号", "零部件代号"])  # next header
        es = _make_excel_sheet(rows)
        parts = _collect_all_tables(es, len(rows), 2, "test.xlsx")
        # Should not crash, should find at most 10 tables (10 part numbers)
        assert len(parts) <= 10, f"Expected at most 10 tables, got {len(parts)} parts"
        assert len(parts) == 10, f"Expected exactly 10 parts (max_tables=10), got {len(parts)}"
        pns = [p[0] for p in parts]
        assert "P000" in pns
        assert "P009" in pns

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
        parts = _collect_all_tables(es, 7, 2, "test.xlsx")
        assert parts == [], f"Expected empty list (no valid parts), got {len(parts)}"

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
        parts = _collect_all_tables(es, 11, 2, "test.xlsx")
        assert len(parts) == 2, f"Expected 2 parts from 2 staggered tables, got {len(parts)}"
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
        fd, path = tempfile.mkstemp(suffix=".xlsx", prefix="card_multi_", dir="/home/rpaup/projects/burlak")
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
        fd, path = tempfile.mkstemp(suffix=".txt", prefix="card_test_", dir="/home/rpaup/projects/burlak")
        os.close(fd)
        try:
            files = _find_excel_files(path)
            assert files == [], f"Expected empty for .txt file, got {files}"
        finally:
            _safe_remove(path)

    def test_directory_with_xlsx(self):
        """Directory containing .xlsx files should find them."""
        tmpdir = tempfile.mkdtemp(prefix="card_dir_", dir="/home/rpaup/projects/burlak")
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

    def test_directory_with_nested_xlsx(self):
        """Directory with nested .xlsx files should find all."""
        tmpdir = tempfile.mkdtemp(prefix="card_nest_", dir="/home/rpaup/projects/burlak")
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

    def test_directory_empty_returns_empty(self):
        """Empty directory returns empty list."""
        tmpdir = tempfile.mkdtemp(prefix="card_empty_", dir="/home/rpaup/projects/burlak")
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
