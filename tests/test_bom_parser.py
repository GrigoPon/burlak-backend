"""Unit-тесты для bom_parser.py.

Покрытие:
  - PartInfo, BOMData (data structures)
  - parse_bom с T1L-стилем (много комплектаций)
  - parse_bom с одной qty-колонкой (спец-листы附件)
  - parse_bom: мульти-листовая обработка, global_names
  - parse_bom: служебные листы (skip)
  - parse_bom: дедупликация имён комплектаций
  - parse_bom: empty/service листы
  - get_config_quantities, get_all_config_quantities
  - lookup_part_name
  - BOMService
"""

from __future__ import annotations

import os
import tempfile
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
import pytest
from openpyxl import Workbook

from burlak_parser.bom_parser import (
    BOMData,
    BOMService,
    PartInfo,
    get_all_config_quantities,
    get_config_quantities,
    lookup_part_name,
    parse_bom,
)


# ═══════════════════════════════════════════════════════════════════════
#  ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ═══════════════════════════════════════════════════════════════════════

def _save_workbook(wb: Workbook, file_path: str) -> str:
    """Сохранить Workbook во временный .xlsx файл и вернуть путь."""
    wb.save(file_path)
    return file_path


def _create_xlsx(
    sheets_data: Dict[str, List[List[Optional[Any]]]],
    suffix: str = ".xlsx",
) -> str:
    """Создать .xlsx файл с несколькими листами и вернуть путь к нему.

    sheets_data: {sheet_name: [row_data, ...]}
    """
    fd, path = tempfile.mkstemp(suffix=suffix, prefix="bom_test_", dir="/home/rpaup/projects/burlak")
    os.close(fd)

    wb = Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    for sheet_name, data in sheets_data.items():
        ws = wb.create_sheet(title=sheet_name)
        for r_idx, row in enumerate(data, 1):
            for c_idx, val in enumerate(row, 1):
                if val is not None:
                    ws.cell(row=r_idx, column=c_idx, value=val)

    wb.save(path)
    return path


# ═══════════════════════════════════════════════════════════════════════
#  1. PartInfo
# ═══════════════════════════════════════════════════════════════════════

class TestPartInfo:
    def test_default_creation(self):
        p = PartInfo(part_number="ABC001")
        assert p.part_number == "ABC001"
        assert p.name_cn == ""
        assert p.name_en == ""
        assert p.quantity == 0.0
        assert p.applicable_configs == []

    def test_full_creation(self):
        p = PartInfo(
            part_number="ABC001",
            name_cn="中文名",
            name_en="English Name",
            quantity=2.0,
            applicable_configs=["Config1", "Config2"],
        )
        assert p.part_number == "ABC001"
        assert p.name_cn == "中文名"
        assert p.name_en == "English Name"
        assert p.quantity == 2.0
        assert p.applicable_configs == ["Config1", "Config2"]

    def test_default_factory(self):
        p1 = PartInfo(part_number="P001")
        p2 = PartInfo(part_number="P002")
        assert p1.applicable_configs == []
        assert p2.applicable_configs == []
        # Check they are independent lists
        p1.applicable_configs.append("C1")
        assert len(p1.applicable_configs) == 1
        assert len(p2.applicable_configs) == 0


# ═══════════════════════════════════════════════════════════════════════
#  2. BOMData
# ═══════════════════════════════════════════════════════════════════════

class TestBOMData:
    def test_empty_creation(self):
        bom = BOMData(
            parts={},
            config_names=[],
            config_quantities={},
        )
        assert bom.parts == {}
        assert bom.config_names == []
        assert bom.config_quantities == {}
        assert bom.source_file == ""
        assert bom.global_names == {}

    def test_with_data(self):
        parts = {
            "P001": PartInfo(part_number="P001", name_cn="Деталь1", quantity=1.0),
            "P002": PartInfo(part_number="P002", name_cn="Деталь2", quantity=2.0),
        }
        config_qty = {
            "Конфиг1": {"P001": 1.0},
            "Конфиг2": {"P001": 2.0, "P002": 1.0},
        }
        global_names = {"P001": ("Деталь1", ""), "P002": ("Деталь2", "")}

        bom = BOMData(
            parts=parts,
            config_names=["Конфиг1", "Конфиг2"],
            config_quantities=config_qty,
            source_file="test.xlsx",
            global_names=global_names,
        )
        assert len(bom.parts) == 2
        assert len(bom.config_names) == 2
        assert bom.config_quantities["Конфиг1"]["P001"] == 1.0
        assert bom.source_file == "test.xlsx"


# ═══════════════════════════════════════════════════════════════════════
#  3. parse_bom — T1L-style BOM (multi-config)
# ═══════════════════════════════════════════════════════════════════════

class TestParseBomT1LStyle:
    """T1L BOM: headers at row 3, part_no=C2, name=C3, 4 config columns."""

    @pytest.fixture
    def t1l_xlsx(self) -> str:
        """Create a T1L-style BOM file with 4 configs and 5 parts."""
        data = [
            ["T1L WE地区 CKD BOM CT1260301", None, None, None, None, None, None, None],
            [None, None, None, None, None, None, None, None],
            ["序号", "零部件件号", "零部件名称", "CPAC编码",
             "舒享版-全黑内饰", "舒享版-黑米内饰", "奢享版-全黑内饰", "奢享版-黑米内饰"],
            ["1", "132000184AA", "变速箱控制单元支架", "CPAC001", "1", "1", "1", "1"],
            ["2", "551002664AA", "TCU", "CPAC002", "1", "1", "1", "1"],
            ["3", "5306200-ED001", "仪表板横梁总成 / Поперечная балка", "CPAC003", "1", "1", "2", "2"],
            ["4", "Q146Z0825F36", "螺栓", "CPAC004", "4", "4", "4", "4"],
            ["5", "G086A001", "卡扣", "CPAC005", "2", "0", "2", "0"],
        ]
        return _create_xlsx({"总装BOM": data})

    def test_t1l_bom_parsed(self, t1l_xlsx: str):
        bom = parse_bom(t1l_xlsx)
        assert len(bom.config_names) == 4, f"Expected 4 configs, got {len(bom.config_names)}: {bom.config_names}"
        assert len(bom.parts) == 5, f"Expected 5 parts, got {len(bom.parts)}"
        assert len(bom.global_names) == 5, "All 5 parts should have names"

    def test_t1l_part_numbers(self, t1l_xlsx: str):
        bom = parse_bom(t1l_xlsx)
        expected_parts = {"132000184AA", "551002664AA", "5306200ED001", "Q146Z0825F36", "G086A001"}
        assert set(bom.parts.keys()) == expected_parts, f"Got {set(bom.parts.keys())}"

    def test_t1l_config_quantities(self, t1l_xlsx: str):
        bom = parse_bom(t1l_xlsx)
        # Check a specific part quantity in a specific config
        # 5306200-ED001 should have qty=2 in 奢享版 configs, qty=1 in 舒享版
        for config_name in bom.config_names:
            if "奢享版" in config_name:
                assert bom.config_quantities[config_name]["5306200ED001"] == 2.0, \
                    f"Expected qty=2 for 5306200ED001 in {config_name}"
            if "舒享版" in config_name:
                if "黑米" in config_name:
                    assert bom.config_quantities[config_name].get("G086A001") is None or \
                           bom.config_quantities[config_name].get("G086A001", 0) == 0, \
                        "G086A001 should not exist in 舒享版-黑米"
                else:
                    assert bom.config_quantities[config_name].get("G086A001", 0) == 2.0

    def test_t1l_config_names(self, t1l_xlsx: str):
        bom = parse_bom(t1l_xlsx)
        expected_configs = {"舒享版-全黑内饰", "舒享版-黑米内饰", "奢享版-全黑内饰", "奢享版-黑米内饰"}
        config_set = set(bom.config_names)
        assert config_set == expected_configs, f"Got {config_set}"

    def test_t1l_part_info_applicable_configs(self, t1l_xlsx: str):
        bom = parse_bom(t1l_xlsx)
        pn = "132000184AA"
        part = bom.parts[pn]
        assert len(part.applicable_configs) == 4, \
            f"Part {pn} should be in all 4 configs, got {part.applicable_configs}"

    def test_t1l_global_names(self, t1l_xlsx: str):
        bom = parse_bom(t1l_xlsx)
        name_cn, name_en = bom.global_names.get("132000184AA", ("", ""))
        assert "变速箱" in name_cn, f"Expected '变速箱', got '{name_cn}'"


# ═══════════════════════════════════════════════════════════════════════
#  4. parse_bom — Russian G01-style BOM
# ═══════════════════════════════════════════════════════════════════════

class TestParseBomG01Style:
    """G01 Russian BOM: headers at row 1, Russian/Chinese headers."""

    @pytest.fixture
    def g01_xlsx(self) -> str:
        data = [
            ["序号\nСерийный номер", "零部件件号\nКод детали",
             "零部件名称\nНаименование", "系统\nСистема",
             "舒享版-全黑", "舒享版-黑米", "奢享版-全黑", "奢享版-黑米"],
            ["1", "5306200-ED001", "仪表板横梁总成 / Поперечная балка", "A", "1", "1", "2", "2"],
            ["2", "551002664AA", "TCU", "B", "1", "1", "1", "1"],
            ["3", "Q146Z0825F36", "Болт", "A", "4", "4", "4", "4"],
        ]
        return _create_xlsx({"G01 BOM": data})

    def test_g01_bom_parsed(self, g01_xlsx: str):
        bom = parse_bom(g01_xlsx)
        assert len(bom.config_names) == 4, f"Expected 4 configs, got {bom.config_names}"
        assert len(bom.parts) == 3, f"Expected 3 parts, got {len(bom.parts)}"

    def test_g01_part_no_clean(self, g01_xlsx: str):
        bom = parse_bom(g01_xlsx)
        assert "5306200ED001" in bom.parts, "Part number not cleaned correctly"
        assert "5306200-ED001" not in bom.parts, "Original part number should be cleaned"

    def test_g01_config_quantities(self, g01_xlsx: str):
        bom = parse_bom(g01_xlsx)
        # 5306200ED001: qty=1 in舒享版, 2 in奢享版
        for cn in bom.config_names:
            pn = "5306200ED001"
            if "奢享版" in cn:
                assert bom.config_quantities[cn][pn] == 2.0, f"{cn}: expected 2, got {bom.config_quantities[cn][pn]}"
            if "舒享版" in cn:
                assert bom.config_quantities[cn][pn] == 1.0


# ═══════════════════════════════════════════════════════════════════════
#  5. parse_bom — Single qty column (附件 style)
# ═══════════════════════════════════════════════════════════════════════

class TestParseBomAttachmentSheet:
    """Лист附件 (attachment) с одной qty-колонкой, без config columns.

    Attachment sheet обрабатывается ВТОРЫМ (после основного BOM-листа),
    поэтому его части добавляются только в global_names, а не в config_quantities.
    """

    @pytest.fixture
    def attachment_xlsx(self) -> str:
        """Create a BOM with: primary sheet (multi-config) + attachment sheet (1 qty).

        Both sheets need >= 3 data rows for is_sheet_bom_candidate.
        """
        data = {
            "总装BOM": [
                ["序号", "零部件件号", "零部件名称",
                 "舒享版-全黑", "舒享版-黑米", "奢享版-全黑"],
                ["1", "P001", "Part One", "1", "1", "2"],
                ["2", "P002", "Part Two", "1", "0", "1"],
                ["3", "P003", "Part Three", "2", "2", "2"],
                ["4", "P006", "Part Six", "1", "1", "0"],
            ],
            "零部件附件": [
                ["零部件件号", "组件物料描述", "组件数量"],
                ["P004", "附件部件4", "5"],
                ["P005", "附件部件5", "3"],
                ["P007", "附件部件7", "2"],
            ],
        }
        return _create_xlsx(data)

    def test_attachment_sheet_adds_global_names(self, attachment_xlsx: str):
        """Attachment sheet adds its parts to global_names (not configs).

        Второй лист (零部件附件) обрабатывается после primary_bom_found=True →
        его части попадают только в global_names, не в bom.parts.
        """
        bom = parse_bom(attachment_xlsx)
        # Attachment parts are in global_names, not bom.parts
        assert "P004" in bom.global_names, "Attachment part missing from global_names"
        assert "P005" in bom.global_names, "Attachment part missing from global_names"
        # Parts from PRIMARY sheet go to bom.parts
        assert "P001" in bom.parts, "Primary part missing from bom.parts"
        # Attachment parts NOT in bom.parts (second sheet, only global_names)
        assert "P004" not in bom.parts, "Attachment part should not be in bom.parts (second sheet)"

    def test_primary_configs_only(self, attachment_xlsx: str):
        """Only primary sheet contributes configs."""
        bom = parse_bom(attachment_xlsx)
        # Primary sheet has 3 config columns: 舒享版-全黑, 舒享版-黑米, 奢享版-全黑
        assert len(bom.config_names) == 3, f"Expected 3 configs, got {bom.config_names}"
        # Attachment sheet name should NOT become a config (processed second)
        assert "零部件附件" not in bom.config_names, "Attachment sheet should not add a config"

    def test_primary_parts_in_config(self, attachment_xlsx: str):
        """Primary sheet parts should be in config quantities."""
        bom = parse_bom(attachment_xlsx)
        cn = bom.config_names[0]
        assert "P001" in bom.config_quantities[cn], "Primary part missing from config"
        # P004 from attachment should NOT be in config quantities
        assert "P004" not in bom.config_quantities[cn], "Attachment part should not be in config"


# ═══════════════════════════════════════════════════════════════════════
#  6. parse_bom — Multi-sheet: only first BOM gives configs
# ═══════════════════════════════════════════════════════════════════════

class TestParseBomMultiSheet:
    """Multiple BOM sheets: only the first provides configs, others just add names.

    Важно: оба листа распознаются как BOM-кандидаты (有config columns + part_no + name).
    焊装BOM — первый → даёт конфиги и global_names.
    涂装BOM — второй → только global_names.
    """

    @pytest.fixture
    def multi_xlsx(self) -> str:
        """Both sheets need >= 3 data rows for is_sheet_bom_candidate."""
        data = {
            "焊装BOM": [
                ["序号", "零部件件号", "零件名称", "舒享版", "奢享版"],
                ["1", "W001", "Welded Part 1", "2", "2"],
                ["2", "W002", "Welded Part 2", "1", "1"],
                ["3", "W003", "Welded Part 3", "1", "0"],
            ],
            "涂装BOM": [
                ["序号", "零部件件号", "零件名称", "舒享版", "奢享版"],
                ["1", "P001", "Painted Part 1", "1", "2"],
                ["2", "P002", "Painted Part 2", "1", "0"],
                ["3", "P003", "Painted Part 3", "0", "1"],
            ],
        }
        return _create_xlsx(data)

    def test_first_sheet_provides_configs(self, multi_xlsx: str):
        bom = parse_bom(multi_xlsx)
        assert len(bom.config_names) == 2, f"Expected 2 configs, got {bom.config_names}"
        assert "舒享版" in bom.config_names
        assert "奢享版" in bom.config_names

    def test_both_sheets_parts_collected(self, multi_xlsx: str):
        bom = parse_bom(multi_xlsx)
        # First sheet parts go to bom.parts (焊装BOM is primary)
        assert "W001" in bom.parts, "First sheet part missing"
        # Second sheet parts go to global_names only (涂装BOM processed after primary_bom_found)
        assert "P001" in bom.global_names, "Second sheet part missing from global_names"
        assert "P001" not in bom.parts, "Second sheet part should NOT be in bom.parts"
        # Total in bom.parts: only 3 parts from 焊装BOM
        assert len(bom.parts) == 3, f"Expected 3 parts (first sheet only), got {len(bom.parts)}"

    def test_second_sheet_adds_global_names(self, multi_xlsx: str):
        bom = parse_bom(multi_xlsx)
        assert "P001" in bom.global_names, "Second sheet part missing from global names"
        cn, _ = bom.global_names["P001"]
        assert "Painted" in cn, f"Expected 'Painted Part 1', got '{cn}'"

    def test_second_sheet_does_not_add_to_configs(self, multi_xlsx: str):
        bom = parse_bom(multi_xlsx)
        # P001 from 涂装BOM should NOT be in config quantities (second sheet only adds names)
        for cn in bom.config_names:
            assert "P001" not in bom.config_quantities[cn], \
                f"P001 should not be in {cn} config (second sheet)"


# ═══════════════════════════════════════════════════════════════════════
#  7. parse_bom — Service sheets are skipped
# ═══════════════════════════════════════════════════════════════════════

class TestParseBomServiceSheets:
    """BOM with service sheets that should be ignored."""

    @pytest.fixture
    def service_xlsx(self) -> str:
        """Main BOM sheet needs >= 3 data rows for is_sheet_bom_candidate."""
        data = {
            "总装BOM": [
                ["序号", "零部件件号", "零件名称", "Config1", "Config2"],
                ["1", "P001", "Part1", "1", "2"],
                ["2", "P002", "Part2", "1", "1"],
                ["3", "P003", "Part3", "2", "0"],
            ],
            "EBOM": [
                ["Just engineering data", None],
            ],
            "变更记录": [
                ["Change", "Log"],
            ],
            "封面": [
                ["Title", None],
            ],
        }
        return _create_xlsx(data)

    def test_service_sheets_skipped(self, service_xlsx: str):
        bom = parse_bom(service_xlsx)
        # Only 总装BOM is a BOM candidate — EBOM/变更记录/封面 are service sheets
        assert len(bom.parts) == 3, f"Expected 3 parts (from 总装BOM), got {len(bom.parts)}"
        assert "P001" in bom.parts
        assert len(bom.config_names) == 2

    def test_service_sheets_dont_add_configs(self, service_xlsx: str):
        bom = parse_bom(service_xlsx)
        # EBOM/变更记录/封面 should not become config names
        for svc in ["EBOM", "变更记录", "封面"]:
            assert svc not in bom.config_names, f"'{svc}' should not be a config name"


# ═══════════════════════════════════════════════════════════════════════
#  8. parse_bom — Empty sheet handling
# ═══════════════════════════════════════════════════════════════════════

class TestParseBomEmptySheet:
    @pytest.fixture
    def empty_sheet_xlsx(self) -> str:
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "Main"
        ws1.cell(row=1, column=1, value="序号")
        ws1.cell(row=1, column=2, value="零部件件号")
        ws1.cell(row=1, column=3, value="名称")
        ws1.cell(row=1, column=4, value="Config")

        ws2 = wb.create_sheet(title="EmptySheet")
        # No data at all
        path = os.path.join(tempfile.mkdtemp(dir="/home/rpaup/projects/burlak"), "empty_test.xlsx")
        os.makedirs(os.path.dirname(path), exist_ok=True)
        wb.save(path)
        return path

    def test_empty_sheet_no_crash(self, empty_sheet_xlsx: str):
        """Empty sheets should be skipped gracefully without crash."""
        bom = parse_bom(empty_sheet_xlsx)
        assert bom is not None
        # Main sheet has headers but no data rows → not a BOM candidate → 0 parts
        assert len(bom.parts) == 0, "Should have 0 parts from empty sheets"
        assert len(bom.config_names) == 0, "Should have 0 configs from empty sheets"


# ═══════════════════════════════════════════════════════════════════════
#  9. parse_bom — Config name deduplication
# ═══════════════════════════════════════════════════════════════════════

class TestParseBomConfigDedup:
    @pytest.fixture
    def dedup_xlsx(self) -> str:
        """Config columns with duplicate names. Needs >= 3 data rows."""
        data = [
            ["序号", "零部件件号", "零件名称",
             "Config A", "Config A", "Config B"],
            ["1", "P001", "Part1", "1", "2", "1"],
            ["2", "P002", "Part2", "1", "0", "2"],
            ["3", "P003", "Part3", "2", "1", "0"],
        ]
        return _create_xlsx({"总装BOM": data})

    def test_config_dedup(self, dedup_xlsx: str):
        bom = parse_bom(dedup_xlsx)
        # After dedup, "Config A" should appear only once (first occurrence kept)
        config_a_count = sum(1 for c in bom.config_names if c == "Config A")
        assert config_a_count == 1, f"Config A should be deduplicated, got {bom.config_names}"
        assert "Config B" in bom.config_names
        assert len(bom.config_names) == 2, f"Expected 2 configs after dedup, got {bom.config_names}"

    def test_dedup_first_column_kept(self, dedup_xlsx: str):
        """Dedup removes duplicate columns but keeps only FIRST column's quantities."""
        bom = parse_bom(dedup_xlsx)
        # P001 in Config A: first "Config A" col has qty=1, second has qty=2
        # After dedup, only first column's data is kept (qty=1)
        assert bom.config_quantities["Config A"]["P001"] == 1.0, \
            f"Expected 1.0 (first column only), got {bom.config_quantities['Config A'].get('P001')}"


# ═══════════════════════════════════════════════════════════════════════
#  10. parse_bom — global_names fallback to parts
# ═══════════════════════════════════════════════════════════════════════

class TestParseBomGlobalNames:
    @pytest.fixture
    def global_names_xlsx(self) -> str:
        """Both sheets need >= 3 data rows for is_sheet_bom_candidate."""
        data = {
            "焊装BOM": [
                ["序号", "零部件件号", "零件名称", "Config1", "Config2"],
                ["1", "P001", "Part One", "1", "1"],
                ["2", "W002", "Weld Part 2", "1", "1"],
                ["3", "W003", "Weld Part 3", "1", "0"],
            ],
            "涂装BOM": [
                ["序号", "零部件件号", "零件名称（英文）", "Config1", "Config2"],
                ["1", "P001", "Part One En", "1", "1"],
                ["2", "P002", "Part Two", "1", "0"],
                ["3", "P003", "Part Three", "0", "1"],
            ],
        }
        return _create_xlsx(data)

    def test_global_names_merged(self, global_names_xlsx: str):
        bom = parse_bom(global_names_xlsx)
        # P001 should have cn from 焊装BOM (first occurrence) and en from 涂装BOM (en column)
        assert "P001" in bom.global_names
        cn, en = bom.global_names["P001"]
        assert "Part One" in cn, f"Expected 'Part One' in cn, got '{cn}'"
        # en should be from 涂装BOM's 零件名称（英文）column
        assert "Part One En" in en, f"Expected 'Part One En' in en, got '{en}'"


# ═══════════════════════════════════════════════════════════════════════
#  11. get_config_quantities
# ═══════════════════════════════════════════════════════════════════════

class TestGetConfigQuantities:
    @pytest.fixture
    def bom_data(self) -> BOMData:
        parts = {
            "P001": PartInfo(part_number="P001", name_cn="Part1"),
            "P002": PartInfo(part_number="P002", name_cn="Part2"),
            "P003": PartInfo(part_number="P003", name_cn="Part3"),
        }
        config_qty = {
            "Config A": {"P001": 2.0, "P002": 1.0},
            "Config B": {"P001": 1.0, "P003": 3.0},
        }
        return BOMData(
            parts=parts,
            config_names=["Config A", "Config B"],
            config_quantities=config_qty,
            global_names={"P001": ("Part1", ""), "P002": ("Part2", ""), "P003": ("Part3", "")},
        )

    def test_get_single_config(self, bom_data: BOMData):
        result = get_config_quantities(bom_data, "Config A")
        assert len(result) == 2
        assert result["P001"].quantity == 2.0
        assert result["P002"].quantity == 1.0
        assert result["P001"].name_cn == "Part1"

    def test_get_config_not_found(self, bom_data: BOMData):
        with pytest.raises(ValueError, match="не найдена"):
            get_config_quantities(bom_data, "NonExistent")

    def test_get_all_configs(self, bom_data: BOMData):
        result = get_all_config_quantities(bom_data)
        assert len(result) == 2
        assert "Config A" in result
        assert "Config B" in result
        assert result["Config A"]["P001"].quantity == 2.0
        assert result["Config B"]["P003"].quantity == 3.0


# ═══════════════════════════════════════════════════════════════════════
#  12. lookup_part_name
# ═══════════════════════════════════════════════════════════════════════

class TestLookupPartName:
    @pytest.fixture
    def bom_data(self) -> BOMData:
        parts = {
            "P001": PartInfo(part_number="P001", name_cn="Part1 CN", name_en="Part1 EN"),
        }
        return BOMData(
            parts=parts,
            config_names=["Config1"],
            config_quantities={"Config1": {"P001": 1.0}},
            global_names={"P001": ("P1 CN", "P1 EN"), "P002": ("P2 CN", "P2 EN")},
        )

    def test_finds_in_parts(self, bom_data: BOMData):
        cn, en = lookup_part_name(bom_data, "P001")
        assert cn == "Part1 CN"
        assert en == "Part1 EN"

    def test_falls_back_to_global(self, bom_data: BOMData):
        cn, en = lookup_part_name(bom_data, "P002")
        assert cn == "P2 CN"
        assert en == "P2 EN"

    def test_not_found(self, bom_data: BOMData):
        cn, en = lookup_part_name(bom_data, "P999")
        assert cn == ""
        assert en == ""

    def test_part_without_name_falls_back(self, bom_data: BOMData):
        """Part exists in parts but without name → fallback to global_names."""
        # Create a part without name
        bom_data.parts["P001"] = PartInfo(part_number="P001")  # no name
        # Remove from global_names to avoid cross-contamination
        bom_data.global_names.pop("P001", None)

        # Now add P001 to global_names
        bom_data.global_names["P001"] = ("Global CN", "Global EN")
        cn, en = lookup_part_name(bom_data, "P001")
        # Should get from global_names since part has no name
        assert cn == "Global CN"
        assert en == "Global EN"


# ═══════════════════════════════════════════════════════════════════════
#  13. BOMService
# ═══════════════════════════════════════════════════════════════════════

class TestBOMService:
    def test_initial_state(self):
        svc = BOMService()
        assert not svc.is_loaded
        assert svc.bom is None

    def test_load(self):
        # Create a simple BOM file with >= 3 data rows
        data = {
            "Test": [
                ["序号", "零部件件号", "零件名称", "用量\\n舒享版", "用量\\n奢享版"],
                ["1", "P001", "Part1", "1", "2"],
                ["2", "P002", "Part2", "1", "1"],
                ["3", "P003", "Part3", "2", "0"],
                ["4", "P004", "Part4", "1", "1"],
            ],
        }
        path = _create_xlsx(data)
        svc = BOMService()
        bom = svc.load(path)
        assert svc.is_loaded
        assert bom is not None
        # 4 data rows + enough configs = BOM candidate
        assert len(bom.parts) == 4

    def test_get_config_names(self):
        data = {
            "Test": [
                ["序号", "零部件件号", "名称", "C1", "C2"],
                ["1", "P001", "Part1", "1", "2"],
                ["2", "P002", "Part2", "1", "1"],
                ["3", "P003", "Part3", "2", "0"],
            ],
        }
        path = _create_xlsx(data)
        svc = BOMService()
        svc.load(path)
        names = svc.get_config_names()
        assert len(names) == 2, f"Expected 2 configs, got {len(names)}"

    def test_get_config_count(self):
        data = {
            "Test": [
                ["序号", "零部件件号", "名称", "C1", "C2", "C3"],
                ["1", "P001", "Part1", "1", "2", "1"],
                ["2", "P002", "Part2", "1", "0", "1"],
                ["3", "P003", "Part3", "2", "2", "0"],
            ],
        }
        path = _create_xlsx(data)
        svc = BOMService()
        svc.load(path)
        assert svc.get_config_count() == 3, f"Expected 3 configs, got {svc.get_config_count()}"

    def test_get_all_part_numbers(self):
        data = {
            "Test": [
                ["序号", "零部件件号", "名称", "C1", "C2"],
                ["1", "P001", "Part1", "1", "2"],
                ["2", "P002", "Part2", "1", "1"],
                ["3", "P003", "Part3", "2", "0"],
            ],
        }
        path = _create_xlsx(data)
        svc = BOMService()
        svc.load(path)
        pns = svc.get_all_part_numbers()
        assert len(pns) == 3, f"Expected 3 parts, got {len(pns)}"
        assert "P001" in pns

    def test_lookup_name(self):
        data = {
            "Test": [
                ["序号", "零部件件号", "零件名称", "C1", "C2"],
                ["1", "P001", "TestPart", "1", "2"],
                ["2", "P002", "Part2", "1", "1"],
                ["3", "P003", "Part3", "2", "0"],
            ],
        }
        path = _create_xlsx(data)
        svc = BOMService()
        svc.load(path)
        cn, en = svc.lookup_name("P001")
        assert "TestPart" in cn, f"Expected 'TestPart', got '{cn}'"

    def test_not_loaded_raises(self):
        svc = BOMService()
        with pytest.raises(RuntimeError):
            svc.get_config_names()
        with pytest.raises(RuntimeError):
            svc.get_parts_for_config("Any")
        with pytest.raises(RuntimeError):
            svc.get_all_configs()


# ═══════════════════════════════════════════════════════════════════════
#  14. Edge cases — parse_bom
# ═══════════════════════════════════════════════════════════════════════

class TestParseBomEdgeCases:
    def test_empty_workbook(self):
        """Workbook with no data sheets."""
        fd, path = tempfile.mkstemp(suffix=".xlsx", prefix="bom_empty_", dir="/home/rpaup/projects/burlak")
        os.close(fd)
        wb = Workbook()
        ws = wb.active
        ws.title = "Empty"
        wb.save(path)
        bom = parse_bom(path)
        assert len(bom.parts) == 0
        assert len(bom.config_names) == 0

    def test_no_part_no_column(self):
        """Sheet without part_no keywords should be skipped."""
        data = [
            ["Title", "Author", "Date"],
            ["BOM", "Test", "2024"],
        ]
        path = _create_xlsx({"Info": data})
        bom = parse_bom(path)
        assert len(bom.parts) == 0, "Should not parse sheet without part_no column"

    def test_part_number_with_dashes_is_cleaned(self):
        """Part numbers with dashes should be normalized."""
        data = [
            ["序号", "零部件件号", "名称", "Config1", "Config2"],
            ["1", "ABC-001-DEF", "Test Part", "1", "2"],
            ["2", "GHI-002", "Another", "2", "0"],
        ]
        # Need 2+ config columns and 3+ data rows for is_sheet_bom_candidate
        # Add a 3rd data row
        data.append(["3", "JKL-003", "Yet Another", "1", "1"])
        path = _create_xlsx({"BOM": data})
        bom = parse_bom(path)
        assert "ABC001DEF" in bom.parts, f"Dashes should be removed. Keys: {list(bom.parts.keys())}"
        assert "GHI002" in bom.parts
        assert "ABC-001-DEF" not in bom.parts

    def test_qty_from_string(self):
        """Quantity should be parsed from string values."""
        data = [
            ["序号", "零部件件号", "名称", "Config1", "Config2"],
            ["1", "P001", "Part", "2.5", "1.5"],
            ["2", "P002", "Part2", "1", "2"],
            ["3", "P003", "Part3", "0", "1"],
        ]
        path = _create_xlsx({"BOM": data})
        bom = parse_bom(path)
        cn = bom.config_names[0]
        assert bom.config_quantities[cn]["P001"] == 2.5, \
            f"Expected 2.5 for P001 in {cn}, got {bom.config_quantities[cn].get('P001')}"

    def test_zero_qty_not_added(self):
        """Parts with zero quantity should not be added to config."""
        data = [
            ["序号", "零部件件号", "名称", "Config1", "Config2"],
            ["1", "P001", "Part1", "0", "0"],
            ["2", "P002", "Part2", "1", "0"],
            ["3", "P003", "Part3", "0", "2"],
        ]
        path = _create_xlsx({"BOM": data})
        bom = parse_bom(path)
        # Need 2+ configs and 3+ data rows for is_sheet_bom_candidate
        cn = bom.config_names[0]
        # P001 has qty=0 in all configs → should NOT be in config quantities
        for cn_name in bom.config_names:
            assert "P001" not in bom.config_quantities[cn_name], \
                f"Zero qty part should not be added to {cn_name}"
            assert "P002" in bom.config_quantities[cn_name] or "P003" in bom.config_quantities[cn_name]

    def test_non_header_rows_before_header(self):
        """Rows before the actual header should be ignored."""
        data = [
            ["T1L WE BOM", None, None, None, None, None],
            [None, None, None, None, None, None],
            [None, None, None, None, None, None],
            ["序号", "零部件件号", "名称", "Config1", "Config2", "Config3"],
            ["1", "P001", "Part1", "1", "2", "1"],
            ["2", "P002", "Part2", "1", "1", "0"],
            ["3", "P003", "Part3", "2", "0", "0"],
        ]
        # Need 2+ config columns and 3+ data rows
        path = _create_xlsx({"BOM": data})
        bom = parse_bom(path)
        assert len(bom.parts) == 3, f"Expected 3 parts, got {len(bom.parts)}: {list(bom.parts.keys())}"
        assert bom.config_quantities[bom.config_names[0]].get("P001", 0) == 1.0

    def test_global_name_applied_to_part_without_name(self):
        """Part from secondary sheet gets name from global_names.

        涂装BOM has no config columns → not a BOM candidate → skipped.
        P001 only gets name from 焊装BOM (first sheet, preserved).
        """
        data = {
            "焊装BOM": [
                ["序号", "零部件件号", "零件名称", "Config1", "Config2"],
                ["1", "P001", "Welded Part", "1", "2"],
                ["2", "W002", "Weld Part 2", "1", "1"],
                ["3", "W003", "Weld Part 3", "2", "0"],
            ],
            "涂装BOM": [
                ["序号", "零部件件号", "零件名称"],  # no config columns → not BOM candidate
                ["1", "P001", "Painted Part"],
            ],
        }
        path = _create_xlsx(data)
        bom = parse_bom(path)
        # 涂装BOM skipped (not a BOM candidate) -> P001 named from 焊装BOM only
        assert "P001" in bom.parts, "P001 should be in parts"
        cn, _ = bom.global_names.get("P001", ("", ""))
        assert "Welded" in cn, f"Expected 'Welded Part' (first sheet), got '{cn}'"
        assert "Painted" not in cn, "Painted Part should not appear (skipped sheet)"
