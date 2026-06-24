"""Unit-тесты для comparator.py.

Покрытие:
  - Data structures (Discrepancy, DiscrepancyType, ConfigComparisonResult, etc.)
  - compare_single_config: 4 типа расхождений (ONLY_IN_BOM, ONLY_IN_CARDS,
    QUANTITY_MISMATCH, FUZZY_MATCH) + perfect match
  - compare_single_config_cached: cached version с global_names
  - compare_all_configs: multi-config comparison
  - format_discrepancy_report: single + multi format
  - MatchingEngine: service wrapper
  - compare(): legacy API
  - _get_card_numbers: helper
  - Edge cases: empty BOM, empty cards, perfect match
"""

from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
from unittest.mock import patch

import pytest

from burlak_parser.bom_parser import BOMData, PartInfo
from burlak_parser.card_parser import (
    CardParseResult,
    CardPart,
    CardsData,
)
from burlak_parser.comparator import (
    ConfigComparisonResult,
    Discrepancy,
    DiscrepancyType,
    MatchingEngine,
    MultiConfigComparisonResult,
    _compare_config_worker,
    _get_card_numbers,
    compare_all_configs,
    compare_single_config,
    compare_single_config_cached,
    format_discrepancy_report,
    verify_integrity,
)
from burlak_parser.fuzzy_matcher import FuzzyMatcher

# ═══════════════════════════════════════════════════════════════════════
#  ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ═══════════════════════════════════════════════════════════════════════


def _make_minimal_bom(
    parts: dict[str, tuple[str, float]],
    config_names: list[str],
    config_quantities: dict[str, dict[str, float]],
    global_names: dict[str, tuple[str, str]] | None = None,
) -> BOMData:
    """Create BOMData from simplified data."""
    part_objects = {}
    for pn, (name_cn, qty) in parts.items():
        part_objects[pn] = PartInfo(
            part_number=pn,
            name_cn=name_cn,
            quantity=qty,
        )
    if global_names is None:
        global_names = {}
        for pn, (name_cn, _) in parts.items():
            global_names[pn] = (name_cn, "")
    return BOMData(
        parts=part_objects,
        config_names=config_names,
        config_quantities=config_quantities,
        global_names=global_names,
    )


def _make_minimal_cards(
    all_parts: dict[str, float],
    part_sources: dict[str, list[tuple[str, str, float]]] | None = None,
) -> CardsData:
    """Create CardsData from simplified data."""
    if part_sources is None:
        part_sources = {}
        for pn, qty in all_parts.items():
            part_sources[pn] = [("Card001", "file.xlsx", qty)]
    return CardsData(
        all_parts=all_parts,
        part_sources=part_sources,
        card_results=[],
        total_cards_processed=1,
    )


def _make_card_result(
    card_number: str = "C001",
    parts: list[tuple[str, float]] | None = None,
) -> CardParseResult:
    """Create a simple CardParseResult."""
    if parts is None:
        parts = []
    card_parts = [
        CardPart(
            part_number=pn, quantity=qty, source_card=card_number, source_sheet="S1"
        )
        for pn, qty in parts
    ]
    aggregated = {pn: qty for pn, qty in parts}
    return CardParseResult(
        card_number=card_number,
        file_path=f"{card_number}.xlsx",
        sheets=[],
        parts=card_parts,
        aggregated_parts=aggregated,
    )


# ═══════════════════════════════════════════════════════════════════════
#  1. Data structures
# ═══════════════════════════════════════════════════════════════════════


class TestDiscrepancyType:
    def test_constants(self):
        assert DiscrepancyType.ONLY_IN_BOM == "Есть в BOM, нет в операционных картах"
        assert DiscrepancyType.ONLY_IN_CARDS == "Есть в операционных картах, нет в BOM"
        assert DiscrepancyType.QUANTITY_MISMATCH == "Разное количество"
        assert DiscrepancyType.FUZZY_MATCH == "Разный формат номера"


class TestDiscrepancy:
    def test_default_creation(self):
        d = Discrepancy(
            part_number="P001",
            name_cn="Деталь",
            name_en="",
            qty_bom=2.0,
            qty_cards=1.0,
            card_numbers=["C001"],
            discrepancy_type=DiscrepancyType.QUANTITY_MISMATCH,
        )
        assert d.part_number == "P001"
        assert d.qty_bom == 2.0
        assert d.qty_cards == 1.0
        assert d.config_name == ""
        assert d.fuzzy_matched_to == ""

    def test_string_format_qty_mismatch(self):
        d = Discrepancy(
            part_number="P001",
            name_cn="",
            name_en="",
            qty_bom=2.0,
            qty_cards=1.0,
            card_numbers=["C001"],
            discrepancy_type=DiscrepancyType.QUANTITY_MISMATCH,
            config_name="舒享版",
        )
        s = str(d)
        assert "Разное количество" in s
        assert "P001" in s
        assert "BOM=2.0" in s
        assert "Карты=1.0" in s
        assert "舒享版" in s

    def test_string_format_only_in_bom(self):
        d = Discrepancy(
            part_number="P002",
            name_cn="",
            name_en="",
            qty_bom=1.0,
            qty_cards=0.0,
            card_numbers=[],
            discrepancy_type=DiscrepancyType.ONLY_IN_BOM,
        )
        s = str(d)
        assert "Есть в BOM" in s

    def test_string_format_only_in_cards(self):
        d = Discrepancy(
            part_number="P003",
            name_cn="",
            name_en="",
            qty_bom=0.0,
            qty_cards=3.0,
            card_numbers=["C002"],
            discrepancy_type=DiscrepancyType.ONLY_IN_CARDS,
        )
        s = str(d)
        assert "Есть в операционных картах" in s

    def test_string_format_fuzzy_match(self):
        d = Discrepancy(
            part_number="5306200-ED001",
            name_cn="",
            name_en="",
            qty_bom=2.0,
            qty_cards=2.0,
            card_numbers=["C003"],
            discrepancy_type=DiscrepancyType.FUZZY_MATCH,
            fuzzy_matched_to="5306200ED001",
        )
        s = str(d)
        assert "Разный формат" in s
        assert "5306200-ED001 -> 5306200ED001" in s


class TestConfigComparisonResult:
    def test_default_creation(self):
        r = ConfigComparisonResult(
            config_name="舒享版",
            discrepancies=[],
        )
        assert r.config_name == "舒享版"
        assert r.discrepancies == []
        assert r.total_bom_parts == 0
        assert r.total_cards_parts == 0
        assert r.matched_parts == 0
        assert r.fuzzy_matched == 0

    def test_with_data(self):
        disc = [
            Discrepancy(
                "P001", "", "", 2.0, 1.0, ["C1"], DiscrepancyType.QUANTITY_MISMATCH
            ),
        ]
        r = ConfigComparisonResult(
            config_name="Test",
            discrepancies=disc,
            total_bom_parts=5,
            total_cards_parts=4,
            matched_parts=3,
            fuzzy_matched=1,
        )
        assert len(r.discrepancies) == 1
        assert r.total_bom_parts == 5
        assert r.matched_parts == 3


class TestMultiConfigComparisonResult:
    def test_default_creation(self):
        mc = MultiConfigComparisonResult(
            config_results=[],
            all_discrepancies=[],
        )
        assert mc.total_configs == 0
        assert mc.total_bom_unique_parts == 0
        assert mc.total_cards_unique_parts == 0

    def test_with_data(self):
        cr = ConfigComparisonResult(config_name="C1", discrepancies=[])
        disc = [Discrepancy("P1", "", "", 1.0, 0.0, [], DiscrepancyType.ONLY_IN_BOM)]
        mc = MultiConfigComparisonResult(
            config_results=[cr],
            all_discrepancies=disc,
            total_configs=1,
            total_bom_unique_parts=10,
            total_cards_unique_parts=8,
        )
        assert len(mc.config_results) == 1
        assert len(mc.all_discrepancies) == 1
        assert mc.total_bom_unique_parts == 10


class TestComparisonResultLegacy:
    def test_default_creation(self):
        r = ConfigComparisonResult(config_name="", discrepancies=[])
        assert r.discrepancies == []
        assert r.total_bom_parts == 0
        assert r.total_cards_parts == 0
        assert r.matched_parts == 0
        assert r.config_name == ""

    def test_with_data(self):
        d = Discrepancy("P1", "", "", 1.0, 0.0, [], DiscrepancyType.ONLY_IN_BOM)
        r = ConfigComparisonResult(
            config_name="Config1",
            discrepancies=[d],
            total_bom_parts=5,
            total_cards_parts=3,
            matched_parts=2,
        )
        assert len(r.discrepancies) == 1
        assert r.config_name == "Config1"
        assert r.matched_parts == 2


# ═══════════════════════════════════════════════════════════════════════
#  2. compare_single_config — Perfect match
# ═══════════════════════════════════════════════════════════════════════


class TestCompareSingleConfigPerfectMatch:
    """BOM and cards have identical parts with matching quantities."""

    @pytest.fixture
    def bom_parts(self) -> dict[str, PartInfo]:
        return {
            "P001": PartInfo("P001", name_cn="Part1", quantity=2.0),
            "P002": PartInfo("P002", name_cn="Part2", quantity=1.0),
        }

    @pytest.fixture
    def cards(self) -> CardsData:
        return _make_minimal_cards(
            {
                "P001": 2.0,
                "P002": 1.0,
            }
        )

    def test_perfect_match_no_discrepancies(self, bom_parts, cards):
        result = compare_single_config(bom_parts, cards, config_name="Config1")
        assert len(result.discrepancies) == 0
        assert result.matched_parts == 2
        assert result.total_bom_parts == 2
        assert result.total_cards_parts == 2

    def test_perfect_match_counts(self, bom_parts, cards):
        result = compare_single_config(bom_parts, cards, config_name="Config1")
        assert result.total_bom_parts == 2
        assert result.total_cards_parts == 2
        assert result.matched_parts == 2

    def test_empty_bom_no_parts(self):
        cards = _make_minimal_cards({"P001": 1.0})
        result = compare_single_config({}, cards, config_name="Empty")
        assert len(result.discrepancies) == 1
        assert result.discrepancies[0].discrepancy_type == DiscrepancyType.ONLY_IN_CARDS
        assert result.total_bom_parts == 0
        assert result.matched_parts == 0

    def test_empty_cards_no_parts(self):
        bom_parts = {"P001": PartInfo("P001", name_cn="Part1", quantity=1.0)}
        cards = _make_minimal_cards({})
        result = compare_single_config(bom_parts, cards, config_name="Config1")
        assert len(result.discrepancies) == 1
        assert result.discrepancies[0].discrepancy_type == DiscrepancyType.ONLY_IN_BOM
        assert result.total_cards_parts == 0
        assert result.matched_parts == 0

    def test_both_empty_no_discrepancies(self):
        result = compare_single_config({}, _make_minimal_cards({}))
        assert len(result.discrepancies) == 0
        assert result.matched_parts == 0


# ═══════════════════════════════════════════════════════════════════════
#  3. compare_single_config — ONLY_IN_BOM
# ═══════════════════════════════════════════════════════════════════════


class TestCompareSingleConfigOnlyInBom:
    """Parts present in BOM but missing from cards."""

    @pytest.fixture
    def bom_parts(self) -> dict[str, PartInfo]:
        return {
            "P001": PartInfo("P001", name_cn="Part1", quantity=2.0),
            "P002": PartInfo("P002", name_cn="Part2", quantity=1.0),
            "P003": PartInfo("P003", name_cn="Part3", quantity=3.0),
        }

    @pytest.fixture
    def cards(self) -> CardsData:
        return _make_minimal_cards(
            {
                "P001": 2.0,  # only P001 in cards
            }
        )

    def test_only_in_bom_found(self, bom_parts, cards):
        result = compare_single_config(bom_parts, cards, config_name="Config1")
        only_in_bom = [
            d
            for d in result.discrepancies
            if d.discrepancy_type == DiscrepancyType.ONLY_IN_BOM
        ]
        assert len(only_in_bom) == 2, f"Expected 2 only-in-bom, got {len(only_in_bom)}"
        part_nos = {d.part_number for d in only_in_bom}
        assert "P002" in part_nos
        assert "P003" in part_nos

    def test_only_in_bom_qty_zero(self, bom_parts, cards):
        result = compare_single_config(bom_parts, cards, config_name="Config1")
        for d in result.discrepancies:
            if d.discrepancy_type == DiscrepancyType.ONLY_IN_BOM:
                assert d.qty_cards == 0.0
                assert d.card_numbers == []
                assert d.name_cn != "", "Name should be preserved from BOM"

    def test_only_in_bom_with_fuzzy_no_false_positive(self, bom_parts):
        """P003 is ONLY_IN_BOM, not fuzzy-matched to something different."""
        cards_with_fuzzy = _make_minimal_cards(
            {
                "P001": 2.0,
                "P003-X": 1.0,  # different even after normalization
            }
        )
        result = compare_single_config(
            bom_parts,
            cards_with_fuzzy,
            config_name="Config1",
        )
        only_in_bom = [
            d
            for d in result.discrepancies
            if d.discrepancy_type == DiscrepancyType.ONLY_IN_BOM
        ]
        # P003 is in BOM but not in cards (P003-X is different)
        assert any(d.part_number == "P003" for d in only_in_bom)


# ═══════════════════════════════════════════════════════════════════════
#  4. compare_single_config — ONLY_IN_CARDS
# ═══════════════════════════════════════════════════════════════════════


class TestCompareSingleConfigOnlyInCards:
    """Parts present in cards but missing from BOM config."""

    @pytest.fixture
    def bom_parts(self) -> dict[str, PartInfo]:
        return {
            "P001": PartInfo("P001", name_cn="Part1", quantity=1.0),
        }

    @pytest.fixture
    def cards(self) -> CardsData:
        return _make_minimal_cards(
            {
                "P001": 1.0,
                "P002": 3.0,  # only in cards
                "P003": 2.0,  # only in cards
                "P004": 1.0,  # only in cards
            }
        )

    def test_only_in_cards_found(self, bom_parts, cards):
        result = compare_single_config(bom_parts, cards, config_name="Config1")
        only_in_cards = [
            d
            for d in result.discrepancies
            if d.discrepancy_type == DiscrepancyType.ONLY_IN_CARDS
        ]
        assert len(only_in_cards) == 3, (
            f"Expected 3 only-in-cards, got {len(only_in_cards)}"
        )
        part_nos = {d.part_number for d in only_in_cards}
        assert part_nos == {"P002", "P003", "P004"}

    def test_only_in_cards_qty_preserved(self, bom_parts, cards):
        result = compare_single_config(bom_parts, cards, config_name="Config1")
        for d in result.discrepancies:
            if d.discrepancy_type == DiscrepancyType.ONLY_IN_CARDS:
                assert d.qty_bom == 0.0
                assert d.qty_cards > 0.0
                # Card numbers should be extracted
                assert len(d.card_numbers) > 0

    def test_only_in_cards_names_empty(self, bom_parts, cards):
        result = compare_single_config(bom_parts, cards, config_name="Config1")
        for d in result.discrepancies:
            if d.discrepancy_type == DiscrepancyType.ONLY_IN_CARDS:
                assert d.name_cn == "", "Parts in cards only have no BOM name"


# ═══════════════════════════════════════════════════════════════════════
#  5. compare_single_config — QUANTITY_MISMATCH
# ═══════════════════════════════════════════════════════════════════════


class TestCompareSingleConfigQtyMismatch:
    """Parts in both BOM and cards but with different quantities."""

    @pytest.fixture
    def bom_parts(self) -> dict[str, PartInfo]:
        return {
            "P001": PartInfo("P001", name_cn="Part1", quantity=2.0),
            "P002": PartInfo("P002", name_cn="Part2", quantity=1.0),
            "P003": PartInfo("P003", name_cn="Part3", quantity=3.0),
        }

    @pytest.fixture
    def cards(self) -> CardsData:
        return _make_minimal_cards(
            {
                "P001": 1.0,  # mismatch: BOM=2, cards=1
                "P002": 1.0,  # match: BOM=1, cards=1
                "P003": 5.0,  # mismatch: BOM=3, cards=5
            }
        )

    def test_qty_mismatch_found(self, bom_parts, cards):
        result = compare_single_config(bom_parts, cards, config_name="Config1")
        mismatches = [
            d
            for d in result.discrepancies
            if d.discrepancy_type == DiscrepancyType.QUANTITY_MISMATCH
        ]
        assert len(mismatches) == 2, f"Expected 2 qty mismatches, got {len(mismatches)}"
        part_nos = {d.part_number for d in mismatches}
        assert part_nos == {"P001", "P003"}

    def test_perfect_match_excluded(self, bom_parts, cards):
        result = compare_single_config(bom_parts, cards, config_name="Config1")
        mismatches = [
            d
            for d in result.discrepancies
            if d.discrepancy_type == DiscrepancyType.QUANTITY_MISMATCH
        ]
        # P002 has matching qty → should NOT appear
        assert all(d.part_number != "P002" for d in mismatches)
        assert result.matched_parts >= 1  # at least P002 matched

    def test_qty_mismatch_counts(self, bom_parts, cards):
        result = compare_single_config(bom_parts, cards, config_name="Config1")
        for d in result.discrepancies:
            if d.discrepancy_type == DiscrepancyType.QUANTITY_MISMATCH:
                assert abs(d.qty_bom - d.qty_cards) > 0.001
                assert d.name_cn != ""

    def test_small_qty_diff_not_mismatch(self):
        """Differences <= 0.001 should not trigger mismatch."""
        bom = {"P001": PartInfo("P001", quantity=1.0)}
        cards = _make_minimal_cards({"P001": 1.0001})  # diff within threshold
        result = compare_single_config(bom, cards, config_name="C1")
        mismatches = [
            d
            for d in result.discrepancies
            if d.discrepancy_type == DiscrepancyType.QUANTITY_MISMATCH
        ]
        assert len(mismatches) == 0, "Small diff should not trigger mismatch"
        assert result.matched_parts == 1


# ═══════════════════════════════════════════════════════════════════════
#  6. compare_single_config — FUZZY_MATCH
# ═══════════════════════════════════════════════════════════════════════


class TestCompareSingleConfigFuzzyMatch:
    """Parts matched via fuzzy (non-exact) comparison."""

    @pytest.fixture
    def bom_parts(self) -> dict[str, PartInfo]:
        return {
            "5306200ED001": PartInfo("5306200ED001", name_cn="Балка", quantity=2.0),
            "P001": PartInfo("P001", name_cn="Part1", quantity=1.0),
        }

    @pytest.fixture
    def cards(self) -> CardsData:
        return _make_minimal_cards(
            {
                "5306200-ED001": 2.0,  # fuzzy-match: same normalized
                "P001": 1.0,  # exact match
            }
        )

    @pytest.fixture
    def fuzzy_matcher(self, bom_parts) -> FuzzyMatcher:
        return FuzzyMatcher(set(bom_parts.keys()))

    def test_fuzzy_match_detected(self, bom_parts, cards, fuzzy_matcher):
        result = compare_single_config(
            bom_parts,
            cards,
            config_name="Config1",
            fuzzy_matcher=fuzzy_matcher,
        )
        fuzzy = [
            d
            for d in result.discrepancies
            if d.discrepancy_type == DiscrepancyType.FUZZY_MATCH
        ]
        assert len(fuzzy) == 1, f"Expected 1 fuzzy match, got {len(fuzzy)}"
        # part_number теперь показывает BOM-оригинал, fuzzy_matched_to — номер из карт
        assert fuzzy[0].part_number == "5306200ED001"
        assert fuzzy[0].fuzzy_matched_to == "5306200-ED001"

    def test_fuzzy_match_qty_preserved(self, bom_parts, cards, fuzzy_matcher):
        result = compare_single_config(
            bom_parts,
            cards,
            config_name="Config1",
            fuzzy_matcher=fuzzy_matcher,
        )
        for d in result.discrepancies:
            if d.discrepancy_type == DiscrepancyType.FUZZY_MATCH:
                assert d.qty_bom == 2.0
                assert d.qty_cards == 2.0
                assert d.name_cn == "Балка"

    def test_fuzzy_match_not_in_only_in_cards(self, bom_parts, cards, fuzzy_matcher):
        """Fuzzy-matched parts should NOT appear in ONLY_IN_CARDS."""
        result = compare_single_config(
            bom_parts,
            cards,
            config_name="Config1",
            fuzzy_matcher=fuzzy_matcher,
        )
        only_in_cards = [
            d
            for d in result.discrepancies
            if d.discrepancy_type == DiscrepancyType.ONLY_IN_CARDS
        ]
        # После исправления: part_number = BOM-оригинал, fuzzy_matched_to = номер из карт
        # ONLY_IN_CARDS показывают оригинальный номер из карт через original_part_numbers
        # Поскольку CardsData создаётся через _make_minimal_cards (без original_part_numbers),
        # номер остаётся как в all_parts ключе — "5306200-ED001" не будет, т.к. он в fuzzy
        for d in only_in_cards:
            assert d.part_number not in {"5306200-ED001", "5306200ED001"}, (
                "Fuzzy-matched should not be in ONLY_IN_CARDS"
            )

    def test_fuzzy_match_not_in_only_in_bom(self, bom_parts, cards, fuzzy_matcher):
        """Fuzzy-matched parts should NOT cause FALSE ONLY_IN_BOM for BOM orig."""
        result = compare_single_config(
            bom_parts,
            cards,
            config_name="Config1",
            fuzzy_matcher=fuzzy_matcher,
        )
        only_in_bom = [
            d
            for d in result.discrepancies
            if d.discrepancy_type == DiscrepancyType.ONLY_IN_BOM
        ]
        # 5306200ED001 is fuzzy-matched → should NOT be in ONLY_IN_BOM
        # part_number теперь BOM-оригинал (5306200ED001)
        assert not any(d.part_number == "5306200ED001" for d in only_in_bom), (
            "Fuzzy-matched BOM part should not appear in ONLY_IN_BOM"
        )

    def test_no_fuzzy_matcher(self, bom_parts, cards):
        """Without fuzzy_matcher, '5306200-ED001' is ONLY_IN_CARDS and '5306200ED001' is ONLY_IN_BOM."""
        result = compare_single_config(
            bom_parts,
            cards,
            config_name="Config1",
            fuzzy_matcher=None,
        )
        only_in_bom = [
            d
            for d in result.discrepancies
            if d.discrepancy_type == DiscrepancyType.ONLY_IN_BOM
        ]
        only_in_cards = [
            d
            for d in result.discrepancies
            if d.discrepancy_type == DiscrepancyType.ONLY_IN_CARDS
        ]
        # ONLY_IN_BOM: part_number = part.part_number = "5306200ED001" (оригинал из BOM)
        # ONLY_IN_CARDS: part_number из cards_data (без original_part_numbers) = "5306200-ED001"
        assert any(d.part_number == "5306200ED001" for d in only_in_bom)
        assert any(d.part_number == "5306200-ED001" for d in only_in_cards)

    def test_fuzzy_match_no_collision(self):
        """Different parts that look similar should NOT fuzzy-match."""
        bom = {"ABCD123": PartInfo("ABCD123", quantity=1.0)}
        cards = _make_minimal_cards({"ABCD124": 1.0})  # different digit
        fm = FuzzyMatcher(set(bom.keys()))
        result = compare_single_config(bom, cards, config_name="C1", fuzzy_matcher=fm)
        fuzzy = [
            d
            for d in result.discrepancies
            if d.discrepancy_type == DiscrepancyType.FUZZY_MATCH
        ]
        assert len(fuzzy) == 0, "ABCD124 != ABCD123 even after normalization"


# ═══════════════════════════════════════════════════════════════════════
#  7. compare_single_config — Mixed types
# ═══════════════════════════════════════════════════════════════════════


class TestCompareSingleConfigMixed:
    """All 4 discrepancy types in one comparison."""

    @pytest.fixture
    def bom_parts(self) -> dict[str, PartInfo]:
        return {
            "P001": PartInfo("P001", name_cn="Part1", quantity=2.0),
            "P002": PartInfo("P002", name_cn="Part2", quantity=1.0),
            "P003": PartInfo("P003", name_cn="Part3", quantity=3.0),
            "5306200ED001": PartInfo("5306200ED001", name_cn="Балка", quantity=2.0),
        }

    @pytest.fixture
    def cards(self) -> CardsData:
        return _make_minimal_cards(
            {
                "P001": 1.0,  # QTY mismatch
                "P002": 1.0,  # perfect match
                "P004": 2.0,  # ONLY_IN_CARDS
                "5306200-ED001": 2.0,  # fuzzy match
            }
        )

    @pytest.fixture
    def fuzzy_matcher(self, bom_parts) -> FuzzyMatcher:
        return FuzzyMatcher(set(bom_parts.keys()))

    def test_all_types_present(self, bom_parts, cards, fuzzy_matcher):
        result = compare_single_config(
            bom_parts,
            cards,
            config_name="Config1",
            fuzzy_matcher=fuzzy_matcher,
        )
        types = {d.discrepancy_type for d in result.discrepancies}
        expected = {
            DiscrepancyType.QUANTITY_MISMATCH,
            DiscrepancyType.ONLY_IN_BOM,
            DiscrepancyType.ONLY_IN_CARDS,
            DiscrepancyType.FUZZY_MATCH,
        }
        assert types == expected, f"Expected all 4 types, got {types}"

    def test_sort_order(self, bom_parts, cards, fuzzy_matcher):
        """Discrepancies should be sorted: QTY → ONLY_BOM → ONLY_CARDS → FUZZY."""
        result = compare_single_config(
            bom_parts,
            cards,
            config_name="Config1",
            fuzzy_matcher=fuzzy_matcher,
        )
        type_order = [d.discrepancy_type for d in result.discrepancies]
        # Check that types appear in expected order
        seen_qty = False
        seen_bom = False
        seen_cards = False
        seen_fuzzy = False
        for t in type_order:
            if t == DiscrepancyType.QUANTITY_MISMATCH:
                seen_qty = True
            elif t == DiscrepancyType.ONLY_IN_BOM:
                assert not seen_cards and not seen_fuzzy, (
                    "ONLY_BOM should come before ONLY_CARDS"
                )
                seen_bom = True
            elif t == DiscrepancyType.ONLY_IN_CARDS:
                assert not seen_fuzzy, "ONLY_CARDS should come before FUZZY"
                seen_cards = True
            elif t == DiscrepancyType.FUZZY_MATCH:
                seen_fuzzy = True
        assert seen_qty and seen_bom and seen_cards and seen_fuzzy


# ═══════════════════════════════════════════════════════════════════════
#  8. compare_single_config_cached
# ═══════════════════════════════════════════════════════════════════════


class TestCompareSingleConfigCached:
    """Cached variant with pre-computed norm set and fuzzy pairs."""

    @pytest.fixture
    def bom_parts(self) -> dict[str, PartInfo]:
        return {
            "P001": PartInfo("P001", name_cn="Part1", quantity=2.0),
            "P002": PartInfo("P002", name_cn="Part2", quantity=1.0),
        }

    @pytest.fixture
    def cards(self) -> CardsData:
        return _make_minimal_cards(
            {
                "P001": 1.0,  # qty mismatch
                "P003": 3.0,  # only in cards
            }
        )

    def test_basic_cached(self, bom_parts, cards):
        result = compare_single_config_cached(
            bom_parts,
            cards,
            config_name="C1",
        )
        # P001: qty mismatch (BOM=2, cards=1)
        # P002: ONLY_IN_BOM (not in cards at all)
        # P003: ONLY_IN_CARDS (not in BOM)
        assert len(result.discrepancies) == 3, (
            f"Expected 3 discrepancies, got {len(result.discrepancies)}"
        )
        assert result.total_bom_parts == 2

    def test_cached_with_global_names(self, bom_parts, cards):
        """ONLY_IN_CARDS parts get names from global_names."""
        global_names = {"P003": ("Global Part3", "")}
        result = compare_single_config_cached(
            bom_parts,
            cards,
            config_name="C1",
            global_names=global_names,
        )
        only_cards = [
            d
            for d in result.discrepancies
            if d.discrepancy_type == DiscrepancyType.ONLY_IN_CARDS
        ]
        assert len(only_cards) == 1
        assert only_cards[0].name_cn == "Global Part3", (
            "Name should come from global_names"
        )

    def test_cached_with_norm_set(self, bom_parts, cards):
        """cards_norm_set excludes normalized-equivalent BOM parts from ONLY_IN_BOM."""
        cards_norm = {"P003"}  # only P003 is truly in cards
        result = compare_single_config_cached(
            bom_parts,
            cards,
            config_name="C1",
            cards_norm_set=cards_norm,
        )
        only_bom = [
            d
            for d in result.discrepancies
            if d.discrepancy_type == DiscrepancyType.ONLY_IN_BOM
        ]
        # P002 is in BOM but not in cards_norm → ONLY_IN_BOM
        assert any(d.part_number == "P002" for d in only_bom)

    def test_cached_with_fuzzy_pairs(self):
        """Pre-computed fuzzy pairs exclude parts from ONLY_IN_CARDS."""
        bom_parts = {
            "5306200ED001": PartInfo("5306200ED001", name_cn="Балка", quantity=2.0),
            "P001": PartInfo("P001", name_cn="Part1", quantity=1.0),
        }
        fuzzy_pairs = {"5306200-ED001": "5306200ED001"}
        cards = _make_minimal_cards(
            {
                "P001": 1.0,
                "5306200-ED001": 2.0,
            }
        )
        result = compare_single_config_cached(
            bom_parts,
            cards,
            config_name="C1",
            fuzzy_matched_pairs=fuzzy_pairs,
        )
        only_cards = [
            d
            for d in result.discrepancies
            if d.discrepancy_type == DiscrepancyType.ONLY_IN_CARDS
        ]
        # 5306200-ED001 is fuzzy-matched → should NOT appear in ONLY_IN_CARDS
        assert not any(d.part_number == "5306200-ED001" for d in only_cards), (
            "Fuzzy pair should exclude from ONLY_IN_CARDS"
        )
        fuzzy = [
            d
            for d in result.discrepancies
            if d.discrepancy_type == DiscrepancyType.FUZZY_MATCH
        ]
        # part_number теперь BOM-оригинал (5306200ED001), fuzzy_matched_to — номер из карт
        assert any(d.part_number == "5306200ED001" for d in fuzzy)

    def test_cached_empty(self):
        result = compare_single_config_cached({}, _make_minimal_cards({}))
        assert len(result.discrepancies) == 0
        assert result.matched_parts == 0


# ═══════════════════════════════════════════════════════════════════════
#  9. compare_all_configs — Multi-config
# ═══════════════════════════════════════════════════════════════════════


class TestCompareAllConfigs:
    """compare_all_configs with multiple configurations."""

    @pytest.fixture
    def bom(self) -> BOMData:
        parts = {
            "P001": PartInfo("P001", name_cn="Part1"),
            "P002": PartInfo("P002", name_cn="Part2"),
            "P003": PartInfo("P003", name_cn="Part3"),
        }
        config_qty = {
            "舒享版": {"P001": 2.0, "P002": 1.0},
            "奢享版": {"P001": 1.0, "P003": 3.0},
        }
        return BOMData(
            parts=parts,
            config_names=["舒享版", "奢享版"],
            config_quantities=config_qty,
            global_names={
                "P001": ("Part1", ""),
                "P002": ("Part2", ""),
                "P003": ("Part3", ""),
            },
        )

    @pytest.fixture
    def cards(self) -> CardsData:
        return _make_minimal_cards(
            {
                "P001": 1.0,  # qty mismatch in both configs
                "P002": 1.0,  # perfect match in 舒享版, absent in 奢享版
                "P004": 3.0,  # only in cards
            }
        )

    def test_all_configs_compared(self, bom, cards):
        result = compare_all_configs(bom, cards, use_fuzzy=True)
        assert result.total_configs == 2, (
            f"Expected 2 configs, got {result.total_configs}"
        )
        assert len(result.config_results) == 2

    def test_discrepancies_per_config(self, bom, cards):
        result = compare_all_configs(bom, cards, use_fuzzy=True)
        # 舒享版: P001 qty mismatch, P004 only in cards, P003 only in BOM (absent)
        # Actually P003 is ONLY in BOM for 舒享版 since it's not in cards
        # Wait: cards has P001, P002, P004. BOM 舒享版 has P001, P002.
        # - P001 qty mismatch: BOM=2, cards=1
        # - P002 perfect match: BOM=1, cards=1
        # - P004 only in cards
        # 奢享版: P001 qty mismatch, P003 only in BOM, P004 only in cards
        for cr in result.config_results:
            assert len(cr.discrepancies) >= 2, (
                f"{cr.config_name}: expected >= 2 discrepancies, got {len(cr.discrepancies)}"
            )

    def test_total_discrepancies_summed(self, bom, cards):
        result = compare_all_configs(bom, cards, use_fuzzy=True)
        total = sum(len(cr.discrepancies) for cr in result.config_results)
        assert len(result.all_discrepancies) == total, (
            "all_discrepancies should be sum of all config discrepancies"
        )

    def test_unique_part_counts(self, bom, cards):
        result = compare_all_configs(bom, cards, use_fuzzy=True)
        assert result.total_bom_unique_parts == 3, "BOM has 3 unique parts"
        assert result.total_cards_unique_parts == 3, "Cards have 3 unique parts"

    def test_single_config_via_compare_all(self):
        """compare_all_configs with single-config BOM."""
        parts = {"P001": PartInfo("P001", quantity=1.0)}
        bom = BOMData(
            parts=parts,
            config_names=["Single"],
            config_quantities={"Single": {"P001": 1.0}},
            global_names={"P001": ("Part1", "")},
        )
        cards = _make_minimal_cards({"P001": 2.0})
        result = compare_all_configs(bom, cards, use_fuzzy=False)
        assert result.total_configs == 1
        assert len(result.all_discrepancies) == 1  # qty mismatch
        assert (
            result.all_discrepancies[0].discrepancy_type
            == DiscrepancyType.QUANTITY_MISMATCH
        )

    def test_no_fuzzy_multi_config(self, bom, cards):
        """Without fuzzy matching, dash-different parts are separate."""
        result = compare_all_configs(bom, cards, use_fuzzy=False)
        assert result.total_configs == 2

    def test_part_only_in_bom_isolated(self):
        """Part in BOM but completely absent from cards → ONLY_IN_BOM."""
        parts = {"P999": PartInfo("P999", name_cn="Missing", quantity=1.0)}
        bom = BOMData(
            parts=parts,
            config_names=["C1"],
            config_quantities={"C1": {"P999": 1.0}},
            global_names={"P999": ("Missing", "")},
        )
        cards = _make_minimal_cards({})  # no cards at all
        result = compare_all_configs(bom, cards, use_fuzzy=False)
        assert len(result.all_discrepancies) == 1, (
            f"Expected 1 discrepancy, got {len(result.all_discrepancies)}"
        )
        assert (
            result.all_discrepancies[0].discrepancy_type == DiscrepancyType.ONLY_IN_BOM
        )
        assert result.all_discrepancies[0].part_number == "P999"


# ═══════════════════════════════════════════════════════════════════════
#  10. _get_card_numbers
# ═══════════════════════════════════════════════════════════════════════


class TestGetCardNumbers:
    def test_single_source(self):
        cards = _make_minimal_cards(
            {"P001": 1.0},
            part_sources={"P001": [("C001", "f.xlsx", 1.0)]},
        )
        nums = _get_card_numbers("P001", cards)
        assert nums == ["C001"]

    def test_multiple_sources_dedup(self):
        """Same card number appearing multiple times should be deduplicated."""
        cards = _make_minimal_cards(
            {"P001": 3.0},
            part_sources={
                "P001": [
                    ("C001", "f1.xlsx", 1.0),
                    ("C002", "f2.xlsx", 1.0),
                    ("C001", "f3.xlsx", 1.0),  # duplicate
                ],
            },
        )
        nums = _get_card_numbers("P001", cards)
        assert len(nums) == 2, f"Expected 2 unique cards, got {nums}"
        assert "C001" in nums
        assert "C002" in nums

    def test_part_not_in_sources(self):
        cards = _make_minimal_cards({})
        nums = _get_card_numbers("NONEXISTENT", cards)
        assert nums == []

    def test_empty_sources(self):
        cards = _make_minimal_cards({}, part_sources={})
        nums = _get_card_numbers("P001", cards)
        assert nums == []


# ═══════════════════════════════════════════════════════════════════════
#  11. format_discrepancy_report
# ═══════════════════════════════════════════════════════════════════════


class TestFormatDiscrepancyReport:
    def test_single_config_no_discrepancies(self):
        r = ConfigComparisonResult(
            config_name="TestConfig",
            discrepancies=[],
            total_bom_parts=5,
            total_cards_parts=5,
            matched_parts=5,
        )
        report = format_discrepancy_report(r)
        assert "TestConfig" in report
        assert "Несоответствий не найдено" in report

    def test_single_config_with_discrepancies(self):
        d = Discrepancy(
            "P001",
            "",
            "",
            2.0,
            1.0,
            ["C1"],
            DiscrepancyType.QUANTITY_MISMATCH,
            config_name="C1",
        )
        r = ConfigComparisonResult(
            config_name="Test",
            discrepancies=[d],
            total_bom_parts=3,
            total_cards_parts=2,
            matched_parts=1,
        )
        report = format_discrepancy_report(r)
        assert "Разное количество" in report
        assert "P001" in report
        assert "2.0" in report
        assert "1.0" in report

    def test_multi_config_no_discrepancies(self):
        mc = MultiConfigComparisonResult(
            config_results=[],
            all_discrepancies=[],
            total_configs=2,
            total_bom_unique_parts=5,
            total_cards_unique_parts=4,
        )
        report = format_discrepancy_report(mc)
        assert "Расхождений не найдено" in report

    def test_multi_config_with_discrepancies(self):
        d = Discrepancy(
            "P001",
            "",
            "",
            2.0,
            1.0,
            ["C1"],
            DiscrepancyType.QUANTITY_MISMATCH,
            config_name="Config1",
        )
        cr = ConfigComparisonResult(
            config_name="Config1",
            discrepancies=[d],
            total_bom_parts=3,
            total_cards_parts=2,
            matched_parts=1,
        )
        mc = MultiConfigComparisonResult(
            config_results=[cr],
            all_discrepancies=[d],
            total_configs=1,
            total_bom_unique_parts=3,
            total_cards_unique_parts=2,
        )
        report = format_discrepancy_report(mc)
        assert "ОТЧЁТ ПРОВЕРКИ КОМПЛЕКТАЦИЙ" in report
        assert "КРАТКАЯ СВОДКА" in report
        assert "Config1" in report
        assert "1.0" in report
        assert "2.0" in report

    def test_multi_config_all_types(self):
        disc = [
            Discrepancy(
                "P001",
                "",
                "",
                2.0,
                1.0,
                [],
                DiscrepancyType.QUANTITY_MISMATCH,
                config_name="C1",
            ),
            Discrepancy(
                "P002",
                "",
                "",
                1.0,
                0.0,
                [],
                DiscrepancyType.ONLY_IN_BOM,
                config_name="C1",
            ),
            Discrepancy(
                "P003",
                "",
                "",
                0.0,
                3.0,
                ["C2"],
                DiscrepancyType.ONLY_IN_CARDS,
                config_name="C1",
            ),
        ]
        cr = ConfigComparisonResult(config_name="C1", discrepancies=disc)
        mc = MultiConfigComparisonResult(
            config_results=[cr],
            all_discrepancies=disc,
            total_configs=1,
            total_bom_unique_parts=5,
            total_cards_unique_parts=4,
        )
        report = format_discrepancy_report(mc)
        assert "Разное количество" in report
        assert "Есть в BOM, нет в операционных картах" in report
        assert "Есть в операционных картах, нет в BOM" in report
        assert "P001" in report
        assert "P002" in report
        assert "P003" in report

    def test_multi_config_format_russian(self):
        """Report should be in Russian as per factory worker expectations."""
        d = Discrepancy(
            "P001",
            "",
            "",
            2.0,
            1.0,
            [],
            DiscrepancyType.QUANTITY_MISMATCH,
            config_name="C1",
        )
        cr = ConfigComparisonResult(config_name="C1", discrepancies=[d])
        mc = MultiConfigComparisonResult(
            config_results=[cr],
            all_discrepancies=[d],
            total_configs=1,
            total_bom_unique_parts=3,
            total_cards_unique_parts=2,
        )
        report = format_discrepancy_report(mc)
        assert "Всего проверено комплектаций" in report
        assert "Найдено несоответствий" in report


# ═══════════════════════════════════════════════════════════════════════
#  13. MatchingEngine
# ═══════════════════════════════════════════════════════════════════════


class TestMatchingEngine:
    def test_default_creation(self):
        engine = MatchingEngine()
        assert engine.use_fuzzy is True

    def test_compare_all_no_fuzzy(self):
        bom = BOMData(
            parts={"P001": PartInfo("P001", quantity=1.0)},
            config_names=["C1"],
            config_quantities={"C1": {"P001": 1.0}},
            global_names={},
        )
        cards = _make_minimal_cards({"P001": 2.0})
        engine = MatchingEngine(use_fuzzy=False)
        result = engine.compare(bom, cards)
        assert isinstance(result, MultiConfigComparisonResult)
        assert result.total_configs == 1
        assert len(result.all_discrepancies) == 1  # qty mismatch

    def test_compare_all_with_fuzzy(self):
        bom = BOMData(
            parts={
                "5306200ED001": PartInfo("5306200ED001", name_cn="Балка", quantity=2.0)
            },
            config_names=["C1"],
            config_quantities={"C1": {"5306200ED001": 2.0}},
            global_names={"5306200ED001": ("Балка", "")},
        )
        cards = _make_minimal_cards({"5306200-ED001": 2.0})
        engine = MatchingEngine(use_fuzzy=True)
        result = engine.compare(bom, cards)
        fuzzy = [
            d
            for d in result.all_discrepancies
            if d.discrepancy_type == DiscrepancyType.FUZZY_MATCH
        ]
        assert len(fuzzy) == 1, "Fuzzy match should be detected by engine"

    def test_single_config_via_engine(self):
        bom = BOMData(
            parts={"P001": PartInfo("P001", quantity=2.0)},
            config_names=["C1", "C2"],
            config_quantities={"C1": {"P001": 2.0}, "C2": {"P001": 1.0}},
            global_names={},
        )
        cards = _make_minimal_cards({"P001": 1.0})
        engine = MatchingEngine(use_fuzzy=False)
        result = engine.compare(bom, cards, single_config="C1")
        assert result.total_configs == 1, "Should compare only C1"
        assert len(result.config_results) == 1
        assert result.config_results[0].config_name == "C1"

    def test_single_config_nonexistent(self):
        """Non-existent single_config should result in empty comparison."""
        bom = BOMData(
            parts={},
            config_names=["C1"],
            config_quantities={"C1": {}},
            global_names={},
        )
        cards = _make_minimal_cards({"P001": 1.0})
        engine = MatchingEngine()
        result = engine.compare(bom, cards, single_config="NonExistent")
        # Should have 1 config result with no BOM parts
        assert len(result.config_results) == 1
        assert result.config_results[0].config_name == "NonExistent"

    def test_fuzzy_engine_perfect_match(self):
        """Engine with fuzzy enabled should still detect exact matches."""
        bom = BOMData(
            parts={"P001": PartInfo("P001", quantity=1.0)},
            config_names=["C1"],
            config_quantities={"C1": {"P001": 1.0}},
            global_names={},
        )
        cards = _make_minimal_cards({"P001": 1.0})
        engine = MatchingEngine(use_fuzzy=True)
        result = engine.compare(bom, cards)
        assert len(result.all_discrepancies) == 0, "Exact match should pass"


# ═══════════════════════════════════════════════════════════════════════
#  14. Edge cases
# ═══════════════════════════════════════════════════════════════════════


class TestCompareEdgeCases:
    def test_all_parts_match_exactly(self):
        """Every part matches perfectly → zero discrepancies."""
        bom = {
            "P001": PartInfo("P001", quantity=1.0),
            "P002": PartInfo("P002", quantity=2.0),
            "P003": PartInfo("P003", quantity=3.0),
        }
        cards = _make_minimal_cards({"P001": 1.0, "P002": 2.0, "P003": 3.0})
        result = compare_single_config(bom, cards)
        assert len(result.discrepancies) == 0
        assert result.matched_parts == 3

    def test_part_with_qty_variations(self):
        """Different qty in different configs generates config-specific discrepancies."""
        parts = {"P001": PartInfo("P001", name_cn="Same")}
        bom = BOMData(
            parts=parts,
            config_names=["C1", "C2"],
            config_quantities={"C1": {"P001": 1.0}, "C2": {"P001": 2.0}},
            global_names={"P001": ("Same", "")},
        )
        cards = _make_minimal_cards({"P001": 1.5})
        result = compare_all_configs(bom, cards, use_fuzzy=False)
        # C1: bom=1.0, cards=1.5 → mismatch
        # C2: bom=2.0, cards=1.5 → mismatch
        assert len(result.all_discrepancies) == 2
        for d in result.all_discrepancies:
            assert d.discrepancy_type == DiscrepancyType.QUANTITY_MISMATCH
            assert d.part_number == "P001"

    def test_same_part_different_qty_in_configs(self):
        """Part with different qty across configs."""
        parts = {"P001": PartInfo("P001", quantity=0.0)}
        bom = BOMData(
            parts=parts,
            config_names=["C_A", "C_B"],
            config_quantities={"C_A": {"P001": 1.0}, "C_B": {"P001": 3.0}},
            global_names={"P001": ("Part1", "")},
        )
        cards = _make_minimal_cards({"P001": 2.0})
        result = compare_all_configs(bom, cards, use_fuzzy=False)
        # Both configs should show mismatch (1:2, 3:2)
        assert len(result.all_discrepancies) == 2

    def test_large_bom_and_cards(self):
        """Performance validation with many parts (no crash)."""
        bom_parts = {}
        cards_parts = {}
        for i in range(100):
            pn = f"P{i:03d}"
            bom_parts[pn] = PartInfo(pn, name_cn=f"Part{i}", quantity=float(i % 5 + 1))
            cards_parts[pn] = float(i % 3 + 1)
        cards = _make_minimal_cards(cards_parts)
        result = compare_single_config(bom_parts, cards, config_name="Large")
        # Some will match, some won't
        assert result.total_bom_parts == 100
        assert len(result.discrepancies) > 0

    def test_part_number_with_spaces_normalized(self):
        """Spaces in part numbers should not affect comparison via fuzzy matcher."""
        bom_parts = {"ABC 123": PartInfo("ABC 123", quantity=1.0)}
        cards = _make_minimal_cards({"ABC-123": 1.0})
        fm = FuzzyMatcher(set(bom_parts.keys()))
        result = compare_single_config(
            bom_parts, cards, config_name="C1", fuzzy_matcher=fm
        )
        fuzzy = [
            d
            for d in result.discrepancies
            if d.discrepancy_type == DiscrepancyType.FUZZY_MATCH
        ]
        assert len(fuzzy) == 1, "Spaces and dashes should fuzzy-match"
        # part_number теперь BOM-оригинал ("ABC 123"), fuzzy_matched_to — номер из карт
        assert fuzzy[0].part_number == "ABC 123"
        assert fuzzy[0].fuzzy_matched_to == "ABC-123"

    def test_no_cards_at_all(self):
        """Empty cards data → all BOM parts are ONLY_IN_BOM."""
        bom_parts = {
            "P001": PartInfo("P001", quantity=1.0),
            "P002": PartInfo("P002", quantity=2.0),
        }
        cards = _make_minimal_cards({})
        result = compare_single_config(bom_parts, cards)
        only_bom = len(
            [
                d
                for d in result.discrepancies
                if d.discrepancy_type == DiscrepancyType.ONLY_IN_BOM
            ]
        )
        assert only_bom == 2, "All BOM parts should be ONLY_IN_BOM"
        assert result.matched_parts == 0

    def test_no_bom_at_all(self):
        """Empty BOM → all card parts are ONLY_IN_CARDS."""
        cards = _make_minimal_cards({"P001": 1.0, "P002": 2.0})
        result = compare_single_config({}, cards)
        only_cards = len(
            [
                d
                for d in result.discrepancies
                if d.discrepancy_type == DiscrepancyType.ONLY_IN_CARDS
            ]
        )
        assert only_cards == 2, "All card parts should be ONLY_IN_CARDS"
        assert result.matched_parts == 0


# ═══════════════════════════════════════════════════════════════════════
#  15. compare_single_config — Invalid part numbers (edge cases)
# ═══════════════════════════════════════════════════════════════════════


class TestCompareSingleConfigInvalidParts:
    """Edge cases: invalid part numbers are skipped from discrepancies."""

    def test_only_in_bom_skip_invalid_part_number(self):
        """Invalid BOM part number skipped from ONLY_IN_BOM (line 140)."""
        bom_parts = {
            "AB": PartInfo("AB", name_cn="Invalid", quantity=1.0),  # 2 chars → invalid
            "P001": PartInfo("P001", name_cn="Part1", quantity=2.0),
        }
        cards = _make_minimal_cards({"P001": 2.0})
        result = compare_single_config(bom_parts, cards, config_name="C1")
        only_bom = [
            d
            for d in result.discrepancies
            if d.discrepancy_type == DiscrepancyType.ONLY_IN_BOM
        ]
        assert len(only_bom) == 0, "AB should be skipped as invalid"
        assert result.matched_parts == 1

    def test_only_in_cards_skip_invalid_part_number(self):
        """Invalid card part number skipped from ONLY_IN_CARDS (line 161)."""
        bom_parts = {
            "P001": PartInfo("P001", name_cn="Part1", quantity=1.0),
        }
        cards = _make_minimal_cards(
            {
                "P001": 1.0,
                "XZ": 3.0,  # 2 chars → invalid
            }
        )
        result = compare_single_config(bom_parts, cards, config_name="C1")
        only_cards = [
            d
            for d in result.discrepancies
            if d.discrepancy_type == DiscrepancyType.ONLY_IN_CARDS
        ]
        assert len(only_cards) == 0, "XZ should be skipped as invalid"
        assert result.total_cards_parts == 2
        assert result.matched_parts == 1

    def test_fuzzy_skip_invalid_bom_part(self):
        """Fuzzy pair with invalid BOM part number is skipped (line 180)."""
        bom_parts = {
            "AB": PartInfo("AB", name_cn="Invalid", quantity=1.0),  # 2 chars → invalid
            "P001": PartInfo("P001", name_cn="Part1", quantity=1.0),
        }
        cards = _make_minimal_cards(
            {
                "A-B": 1.0,  # normalizes to "AB"
                "P001": 1.0,
            }
        )
        fm = FuzzyMatcher(set(bom_parts.keys()))
        result = compare_single_config(
            bom_parts, cards, config_name="C1", fuzzy_matcher=fm
        )
        fuzzy = [
            d
            for d in result.discrepancies
            if d.discrepancy_type == DiscrepancyType.FUZZY_MATCH
        ]
        assert len(fuzzy) == 0, "Fuzzy pair with invalid BOM part should be skipped"

    def test_fuzzy_part_not_in_bom_parts(self):
        """Fuzzy BOM part not in bom_parts dict is skipped (line 183)."""
        all_bom_parts = {"P001", "P002"}
        bom_parts = {
            "P001": PartInfo("P001", name_cn="Part1", quantity=1.0),
        }
        cards = _make_minimal_cards(
            {
                "P001": 1.0,
                "P0-02": 1.0,  # normalizes to "P002"
            }
        )
        fm = FuzzyMatcher(all_bom_parts)
        result = compare_single_config(
            bom_parts, cards, config_name="C1", fuzzy_matcher=fm
        )
        fuzzy = [
            d
            for d in result.discrepancies
            if d.discrepancy_type == DiscrepancyType.FUZZY_MATCH
        ]
        assert len(fuzzy) == 0, "Fuzzy pair with missing BOM part should be skipped"

    def test_common_part_skip_invalid_part_number(self):
        """Invalid common part skipped from qty comparison (line 204)."""
        bom_parts = {
            "XY": PartInfo("XY", name_cn="Invalid", quantity=2.0),
            "P001": PartInfo("P001", name_cn="Part1", quantity=1.0),
        }
        cards = _make_minimal_cards(
            {
                "XY": 1.0,
                "P001": 1.0,
            }
        )
        result = compare_single_config(bom_parts, cards, config_name="C1")
        mismatches = [
            d
            for d in result.discrepancies
            if d.discrepancy_type == DiscrepancyType.QUANTITY_MISMATCH
        ]
        assert len(mismatches) == 0, "XY should be skipped, P001 matches"
        assert result.matched_parts == 1, "Only P001 counted as matched"


# ═══════════════════════════════════════════════════════════════════════
#  16. compare_single_config_cached — Invalid part numbers
# ═══════════════════════════════════════════════════════════════════════


class TestCompareSingleConfigCachedInvalidParts:
    """Edge cases for cached variant: invalid part numbers are skipped."""

    def test_cached_only_in_bom_skip_invalid(self):
        """Cached: invalid BOM part skipped from ONLY_IN_BOM (line 279)."""
        bom_parts = {
            "AB": PartInfo("AB", name_cn="Invalid", quantity=1.0),
            "P001": PartInfo("P001", name_cn="Part1", quantity=2.0),
        }
        cards = _make_minimal_cards({"P001": 2.0})
        result = compare_single_config_cached(bom_parts, cards, config_name="C1")
        only_bom = [
            d
            for d in result.discrepancies
            if d.discrepancy_type == DiscrepancyType.ONLY_IN_BOM
        ]
        assert len(only_bom) == 0

    def test_cached_only_in_cards_skip_invalid(self):
        """Cached: invalid card part skipped from ONLY_IN_CARDS (line 292)."""
        bom_parts = {"P001": PartInfo("P001", name_cn="Part1", quantity=1.0)}
        cards = _make_minimal_cards({"P001": 1.0, "XZ": 3.0})
        result = compare_single_config_cached(bom_parts, cards, config_name="C1")
        only_cards = [
            d
            for d in result.discrepancies
            if d.discrepancy_type == DiscrepancyType.ONLY_IN_CARDS
        ]
        assert len(only_cards) == 0

    def test_cached_fuzzy_skip_invalid_part(self):
        """Cached: fuzzy pair with invalid part is skipped (line 306)."""
        bom_parts = {
            "AB": PartInfo("AB", name_cn="Invalid", quantity=1.0),
            "P001": PartInfo("P001", name_cn="Part1", quantity=1.0),
        }
        cards = _make_minimal_cards({"P001": 1.0, "A-B": 1.0})
        fuzzy_pairs = {"A-B": "AB"}
        result = compare_single_config_cached(
            bom_parts,
            cards,
            config_name="C1",
            fuzzy_matched_pairs=fuzzy_pairs,
        )
        fuzzy = [
            d
            for d in result.discrepancies
            if d.discrepancy_type == DiscrepancyType.FUZZY_MATCH
        ]
        assert len(fuzzy) == 0

    def test_cached_fuzzy_part_not_in_bom_parts(self):
        """Cached: fuzzy BOM part not in bom_parts is skipped (line 309)."""
        bom_parts = {"P001": PartInfo("P001", name_cn="Part1", quantity=1.0)}
        cards = _make_minimal_cards({"P001": 1.0, "P0-02": 1.0})
        fuzzy_pairs = {"P0-02": "P002"}
        result = compare_single_config_cached(
            bom_parts,
            cards,
            config_name="C1",
            fuzzy_matched_pairs=fuzzy_pairs,
        )
        fuzzy = [
            d
            for d in result.discrepancies
            if d.discrepancy_type == DiscrepancyType.FUZZY_MATCH
        ]
        assert len(fuzzy) == 0

    def test_cached_common_part_skip_invalid(self):
        """Cached: invalid common part skipped from qty check (line 326)."""
        bom_parts = {
            "XY": PartInfo("XY", name_cn="Invalid", quantity=2.0),
            "P001": PartInfo("P001", name_cn="Part1", quantity=1.0),
        }
        cards = _make_minimal_cards({"XY": 1.0, "P001": 1.0})
        result = compare_single_config_cached(bom_parts, cards, config_name="C1")
        mismatches = [
            d
            for d in result.discrepancies
            if d.discrepancy_type == DiscrepancyType.QUANTITY_MISMATCH
        ]
        assert len(mismatches) == 0
        assert result.matched_parts == 1


# ═══════════════════════════════════════════════════════════════════════
#  17. _compare_config_worker — Direct call
# ═══════════════════════════════════════════════════════════════════════


class TestCompareConfigWorker:
    """Direct test of _compare_config_worker (covers lines 371-387)."""

    def test_worker_direct_call(self):
        """Call _compare_config_worker directly with proper data."""
        bom_parts_dict = {
            "P001": ("Part1", "Part1_EN", 2.0, "P001"),
        }
        cards_all_parts = {"P001": 1.0, "P002": 3.0}
        cards_part_sources = {
            "P001": [("C001", "f.xlsx", 1.0)],
            "P002": [("C002", "f2.xlsx", 3.0)],
        }
        cards_original = {"P001": "P001", "P002": "P002"}
        result = _compare_config_worker(
            config_name="C1",
            bom_parts_dict=bom_parts_dict,
            cards_all_parts=cards_all_parts,
            cards_part_sources=cards_part_sources,
            cards_original_part_numbers=cards_original,
            fuzzy_matched_pairs={},
            cards_norm_set=set(),
            global_names_dict={"P002": ("Global Part2", "")},
        )
        assert isinstance(result, ConfigComparisonResult)
        assert result.config_name == "C1"
        assert len(result.discrepancies) >= 1
        only_cards = [
            d
            for d in result.discrepancies
            if d.discrepancy_type == DiscrepancyType.ONLY_IN_CARDS
        ]
        assert any(d.part_number == "P002" for d in only_cards)
        for d in only_cards:
            if d.part_number == "P002":
                assert d.name_cn == "Global Part2"


# ═══════════════════════════════════════════════════════════════════════
#  18. compare_all_configs — Parallel error handling
# ═══════════════════════════════════════════════════════════════════════


class TestCompareAllConfigsParallelErrors:
    """Error handling in parallel processing (covers lines 503-504)."""

    def test_parallel_worker_error_handled(self):
        """Worker exception in parallel is caught and logged."""
        parts = {
            "P001": PartInfo("P001", name_cn="Part1"),
            "P002": PartInfo("P002", name_cn="Part2"),
        }
        bom = BOMData(
            parts=parts,
            config_names=["C1", "C2"],
            config_quantities={
                "C1": {"P001": 1.0},
                "C2": {"P002": 2.0},
            },
            global_names={"P001": ("Part1", ""), "P002": ("Part2", "")},
        )
        cards = _make_minimal_cards({"P001": 1.0})
        with patch("burlak_parser.comparator.ProcessPoolExecutor", ThreadPoolExecutor):
            with patch(
                "burlak_parser.comparator._compare_config_worker",
                side_effect=ValueError("Worker crashed"),
            ):
                result = compare_all_configs(bom, cards, use_fuzzy=False)
        assert len(result.config_results) == 0
        assert len(result.all_discrepancies) == 0
        assert result.total_configs == 2


# ═══════════════════════════════════════════════════════════════════════
#  19. format_discrepancy_report — >30 parts per type
# ═══════════════════════════════════════════════════════════════════════


class TestFormatReportManyParts:
    """Report formatting with >30 parts of one type (covers line 732)."""

    def test_multi_config_report_more_than_30_parts(self):
        """Report with >30 ONLY_IN_BOM parts shows '... и ещё N деталей'."""
        disc = []
        for i in range(35):
            pn = f"P{i:03d}"
            disc.append(
                Discrepancy(
                    pn,
                    "",
                    "",
                    1.0,
                    0.0,
                    [],
                    DiscrepancyType.ONLY_IN_BOM,
                    config_name="C1",
                )
            )
        cr = ConfigComparisonResult(config_name="C1", discrepancies=disc[:5])
        mc = MultiConfigComparisonResult(
            config_results=[cr],
            all_discrepancies=disc,
            total_configs=1,
            total_bom_unique_parts=35,
            total_cards_unique_parts=0,
        )
        report = format_discrepancy_report(mc)
        assert "... и ещё 5 деталей" in report

    def test_multi_config_report_exactly_30_parts(self):
        """Report with exactly 30 parts does NOT show '... и ещё'."""
        disc = []
        for i in range(30):
            pn = f"P{i:03d}"
            disc.append(
                Discrepancy(
                    pn,
                    "",
                    "",
                    1.0,
                    0.0,
                    [],
                    DiscrepancyType.ONLY_IN_BOM,
                    config_name="C1",
                )
            )
        cr = ConfigComparisonResult(config_name="C1", discrepancies=disc)
        mc = MultiConfigComparisonResult(
            config_results=[cr],
            all_discrepancies=disc,
            total_configs=1,
            total_bom_unique_parts=30,
            total_cards_unique_parts=0,
        )
        report = format_discrepancy_report(mc)
        assert "... и ещё" not in report

    def test_multi_config_report_mixed_types_with_many(self):
        """Multiple types each with >30 parts show '... и ещё' per type."""
        disc = []
        for i in range(35):
            pn = f"P{i:03d}"
            disc.append(
                Discrepancy(
                    pn,
                    "",
                    "",
                    1.0,
                    0.0,
                    [],
                    DiscrepancyType.ONLY_IN_BOM,
                    config_name="C1",
                )
            )
        qty_disc = []
        for i in range(32):
            pn = f"Q{i:03d}"
            qty_disc.append(
                Discrepancy(
                    pn,
                    "",
                    "",
                    2.0,
                    1.0,
                    [],
                    DiscrepancyType.QUANTITY_MISMATCH,
                    config_name="C1",
                )
            )
        all_disc = disc + qty_disc
        cr = ConfigComparisonResult(config_name="C1", discrepancies=all_disc[:5])
        mc = MultiConfigComparisonResult(
            config_results=[cr],
            all_discrepancies=all_disc,
            total_configs=1,
            total_bom_unique_parts=67,
            total_cards_unique_parts=0,
        )
        report = format_discrepancy_report(mc)
        assert "... и ещё 5 деталей" in report
        assert "... и ещё 2 деталей" in report


# ═══════════════════════════════════════════════════════════════════════
#  20. verify_integrity — верификация целостности
# ═══════════════════════════════════════════════════════════════════════


class TestVerifyIntegrity:
    """verify_integrity — проверка целостности результатов сверки.

    Покрывает:
      - Все комплектации OK
      - Нарушение в одной комплектации
      - Нарушение глобальной суммы
      - Пустой результат
    """

    def test_all_configs_ok(self):
        result = MultiConfigComparisonResult(
            config_results=[
                ConfigComparisonResult(
                    config_name="Config1",
                    discrepancies=[],
                    total_bom_parts=3,
                    total_cards_parts=3,
                    matched_parts=3,
                ),
            ],
            all_discrepancies=[],
            total_configs=1,
        )
        integrity = verify_integrity(result)
        assert integrity.is_ok
        assert integrity.configs_ok == 1
        assert integrity.total_configs == 1
        assert len(integrity.config_issues) == 0
        assert integrity.global_issue == ""

    def test_config_mismatch_detected(self):
        disc = Discrepancy(
            part_number="P001",
            name_cn="",
            name_en="",
            qty_bom=1.0,
            qty_cards=0.0,
            card_numbers=[],
            discrepancy_type=DiscrepancyType.ONLY_IN_BOM,
            config_name="Config1",
        )
        result = MultiConfigComparisonResult(
            config_results=[
                ConfigComparisonResult(
                    config_name="Config1",
                    discrepancies=[disc],
                    total_bom_parts=5,
                    total_cards_parts=4,
                    matched_parts=3,
                ),
            ],
            all_discrepancies=[disc],
            total_configs=1,
        )
        integrity = verify_integrity(result)
        assert not integrity.is_ok
        assert integrity.configs_ok == 0
        assert len(integrity.config_issues) == 1
        assert "Config1" in integrity.config_issues[0]
        assert "учтено 4" in integrity.config_issues[0]
        assert "ожидалось 5" in integrity.config_issues[0]

    def test_global_sum_ok(self):
        disc = Discrepancy(
            part_number="P001",
            name_cn="",
            name_en="",
            qty_bom=1.0,
            qty_cards=0.0,
            card_numbers=[],
            discrepancy_type=DiscrepancyType.ONLY_IN_BOM,
            config_name="Config1",
        )
        result = MultiConfigComparisonResult(
            config_results=[
                ConfigComparisonResult(
                    config_name="Config1",
                    discrepancies=[disc],
                    total_bom_parts=1,
                    total_cards_parts=0,
                    matched_parts=0,
                ),
            ],
            all_discrepancies=[disc],
            total_configs=1,
        )
        integrity = verify_integrity(result)
        assert integrity.is_ok

    def test_empty_result(self):
        result = MultiConfigComparisonResult(
            config_results=[],
            all_discrepancies=[],
            total_configs=0,
        )
        integrity = verify_integrity(result)
        assert integrity.is_ok
        assert integrity.total_configs == 0
        assert integrity.configs_ok == 0

    def test_multiple_configs_mixed_results(self):
        disc = Discrepancy(
            part_number="P001",
            name_cn="",
            name_en="",
            qty_bom=1.0,
            qty_cards=0.0,
            card_numbers=[],
            discrepancy_type=DiscrepancyType.ONLY_IN_BOM,
            config_name="Config1",
        )
        result = MultiConfigComparisonResult(
            config_results=[
                ConfigComparisonResult(
                    config_name="ConfigOK",
                    discrepancies=[],
                    total_bom_parts=2,
                    total_cards_parts=2,
                    matched_parts=2,
                ),
                ConfigComparisonResult(
                    config_name="ConfigBad",
                    discrepancies=[disc],
                    total_bom_parts=10,
                    total_cards_parts=5,
                    matched_parts=4,
                ),
            ],
            all_discrepancies=[disc],
            total_configs=2,
        )
        integrity = verify_integrity(result)
        assert not integrity.is_ok
        assert integrity.configs_ok == 1
        assert integrity.total_configs == 2
        assert len(integrity.config_issues) == 1
        assert "ConfigBad" in integrity.config_issues[0]

    def test_config_has_correct_details(self):
        """details_by_config содержит детали для конфигурации с нарушением."""
        disc = Discrepancy(
            part_number="P001",
            name_cn="",
            name_en="",
            qty_bom=1.0,
            qty_cards=0.0,
            card_numbers=[],
            discrepancy_type=DiscrepancyType.ONLY_IN_BOM,
            config_name="Config1",
        )
        result = MultiConfigComparisonResult(
            config_results=[
                ConfigComparisonResult(
                    config_name="Config1",
                    discrepancies=[disc],
                    total_bom_parts=1,
                    total_cards_parts=0,
                    matched_parts=0,
                ),
            ],
            all_discrepancies=[disc],
            total_configs=1,
        )
        integrity = verify_integrity(result)
        assert integrity.is_ok, "matched(0) + only_in_bom(1) = 1 == total_bom_parts(1)"
        assert integrity.details_by_config[0]["is_ok"] is True
        assert integrity.details_by_config[0]["diff"] == 0

    def test_details_by_config_structure(self):
        result = MultiConfigComparisonResult(
            config_results=[
                ConfigComparisonResult(
                    config_name="TestConfig",
                    discrepancies=[],
                    total_bom_parts=5,
                    total_cards_parts=5,
                    matched_parts=5,
                ),
            ],
            all_discrepancies=[],
            total_configs=1,
        )
        integrity = verify_integrity(result)
        assert len(integrity.details_by_config) == 1
        detail = integrity.details_by_config[0]
        assert detail["config_name"] == "TestConfig"
        assert detail["is_ok"] is True
        assert detail["total_bom_parts"] == 5
        assert detail["accounted"] == 5
        assert detail["matched"] == 5
        assert detail["diff"] == 0

    def test_fuzzy_match_ok(self):
        disc = Discrepancy(
            part_number="5306200ED001",
            name_cn="",
            name_en="",
            qty_bom=2.0,
            qty_cards=2.0,
            card_numbers=["C001"],
            discrepancy_type=DiscrepancyType.FUZZY_MATCH,
            config_name="Config1",
            fuzzy_matched_to="5306200-ED001",
        )
        result = MultiConfigComparisonResult(
            config_results=[
                ConfigComparisonResult(
                    config_name="Config1",
                    discrepancies=[disc],
                    total_bom_parts=1,
                    total_cards_parts=1,
                    matched_parts=0,
                    fuzzy_matched=1,
                ),
            ],
            all_discrepancies=[disc],
            total_configs=1,
        )
        integrity = verify_integrity(result)
        assert integrity.is_ok, "Fuzzy match should be accounted in integrity check"
        assert integrity.configs_ok == 1
