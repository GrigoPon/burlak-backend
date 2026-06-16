"""Unit-тесты для file_classifier.py.

Покрытие:
  - FileClassification dataclass
  - classify_file: AS-паттерн, CARD_START, DIGIT_START, service keywords,
    CP7/CP8, heuristic (extract_card_number), alt heuristic (garbled names),
    unknown format
  - _find_card_number_in_name: поиск в любой части имени
  - _extract_operation_number: все 3 regex-паттерна
  - filter_operational_cards, get_parseable_files, get_splittable_files
"""

from __future__ import annotations

import os
import pytest

from burlak_parser.file_classifier import (
    FileClassification,
    classify_file,
    filter_operational_cards,
    get_parseable_files,
    get_splittable_files,
    _contains_service_keywords,
    _extract_operation_number,
    _find_card_number_in_name,
)


# ═══════════════════════════════════════════════════════════════════════
#  1. FileClassification dataclass
# ═══════════════════════════════════════════════════════════════════════

class TestFileClassification:
    def test_default_operation_number(self):
        """operation_number defaults to empty string."""
        fc = FileClassification(
            file_path="test.xlsx",
            file_name="test",
            parent_folder=".",
            is_operational_card=True,
            is_service_file=False,
            should_split=True,
            should_parse_parts=True,
        )
        assert fc.operation_number == ""

    def test_all_fields_populated(self):
        """All fields can be set explicitly."""
        fc = FileClassification(
            file_path="/path/to/op.xlsx",
            file_name="op",
            parent_folder="to",
            is_operational_card=True,
            is_service_file=False,
            should_split=True,
            should_parse_parts=True,
            operation_number="A001",
        )
        assert fc.file_path == "/path/to/op.xlsx"
        assert fc.file_name == "op"
        assert fc.parent_folder == "to"
        assert fc.is_operational_card is True
        assert fc.is_service_file is False
        assert fc.should_split is True
        assert fc.should_parse_parts is True
        assert fc.operation_number == "A001"

    def test_service_file_fields(self):
        """Service file fields."""
        fc = FileClassification(
            file_path="cover.xlsx",
            file_name="cover",
            parent_folder=".",
            is_operational_card=False,
            is_service_file=True,
            should_split=False,
            should_parse_parts=False,
        )
        assert fc.is_operational_card is False
        assert fc.is_service_file is True
        assert fc.should_split is False
        assert fc.should_parse_parts is False


# ═══════════════════════════════════════════════════════════════════════
#  2. classify_file — operational cards by operation number
# ═══════════════════════════════════════════════════════════════════════

class TestClassifyOperational:
    def test_as_pattern_sqrt(self):
        """SQRT1L-A-AS-04001 matches PREFIX_AS_RE."""
        fc = classify_file("SQRT1L-A-AS-04001-20点扫描.xlsx")
        assert fc.is_operational_card is True
        assert fc.should_parse_parts is True
        assert fc.should_split is True
        assert fc.operation_number == "SQRT1L-A-AS-04001"
        assert fc.is_service_file is False

    def test_as_pattern_g01(self):
        """G01-A-AS-05001 matches PREFIX_AS_RE."""
        fc = classify_file("G01-A-AS-05001-Установка.xlsx")
        assert fc.is_operational_card is True
        assert fc.operation_number == "G01-A-AS-05001"

    def test_card_start_a001(self):
        """A001 matches CARD_START_RE (letter + 2+ digits)."""
        fc = classify_file("A001-Контроль.xlsx")
        assert fc.is_operational_card is True
        assert fc.operation_number == "A001"

    def test_card_start_tp01(self):
        """TP01-операция matches CARD_START_RE (2 letters + 2 digits)."""
        fc = classify_file("TP01-операция.xlsx")
        assert fc.is_operational_card is True
        assert fc.operation_number == "TP01"

    def test_card_start_lg01(self):
        """LG01 matches CARD_START_RE (4 letters + digits)."""
        fc = classify_file("LG01-операция.xlsx")
        assert fc.is_operational_card is True
        assert fc.operation_number == "LG01"

    def test_digit_start_038(self):
        """038 matches DIGIT_START_RE (2+ digits at start)."""
        fc = classify_file("038-Установка двери.xlsx")
        assert fc.is_operational_card is True
        assert fc.operation_number == "038"

    def test_digit_start_1234(self):
        """1234 matches DIGIT_START_RE."""
        fc = classify_file("1234-операция.xlsx")
        assert fc.is_operational_card is True
        assert fc.operation_number == "1234"

    def test_digit_start_overrides_service_keyword(self):
        """Service keyword takes priority over operation number.

        '封面' in filename → service file, even though '038' is an operation number.
        This ensures files like 'CP7作业指导书封面及目录.xlsx' are caught as service.
        """
        fc = classify_file("038-封面.xlsx")
        assert fc.is_operational_card is False
        assert fc.is_service_file is True  # service keyword wins
        assert fc.should_parse_parts is False

    def test_as_overrides_service_keyword(self):
        """Service keyword takes priority over AS pattern.

        '封面' in filename → service file, even with AS pattern.
        """
        fc = classify_file("G01-A-AS-05001-封面.xlsx")
        assert fc.is_operational_card is False
        assert fc.is_service_file is True  # service keyword wins
        assert fc.should_parse_parts is False

    def test_file_with_spaces(self):
        """File with spaces and mixed content (starts with card pattern)."""
        fc = classify_file("LG01 作业指导书.xlsx")
        assert fc.is_operational_card is True
        assert fc.operation_number == "LG01"


# ═══════════════════════════════════════════════════════════════════════
#  3. classify_file — service files (keywords)
# ═══════════════════════════════════════════════════════════════════════

class TestClassifyServiceFile:
    def test_chinese_cover(self):
        """封面 -> service file, no parse, no split."""
        fc = classify_file("封面.xlsx")
        assert fc.is_operational_card is False
        assert fc.is_service_file is True
        assert fc.should_parse_parts is False
        assert fc.should_split is False

    def test_chinese_catalog(self):
        """目录 -> service file."""
        fc = classify_file("目录.xlsx")
        assert fc.is_service_file is True
        assert fc.should_parse_parts is False

    def test_chinese_record_table(self):
        """记录表 -> service file."""
        fc = classify_file("记录表.xlsx")
        assert fc.is_service_file is True

    def test_chinese_empty_template(self):
        """空表 -> service file, template — не разделяется."""
        fc = classify_file("空表.xlsx")
        assert fc.is_service_file is True
        assert fc.should_parse_parts is False
        assert fc.should_split is False

    def test_chinese_fill_template(self):
        """填写范本 -> service file, template."""
        fc = classify_file("填写范本.xlsx")
        assert fc.is_service_file is True

    def test_chinese_fill_instructions(self):
        """填写说明 -> service file, template."""
        fc = classify_file("填写说明.xlsx")
        assert fc.is_service_file is True

    def test_chinese_hours_summary(self):
        """工艺现场工时汇总清单 -> service file."""
        fc = classify_file("工艺现场工时汇总清单.xlsx")
        assert fc.is_service_file is True

    def test_chinese_hours_short(self):
        """工时汇总 -> service file."""
        fc = classify_file("工时汇总.xlsx")
        assert fc.is_service_file is True

    def test_russian_cover(self):
        """обложка -> service file."""
        fc = classify_file("обложка.xlsx")
        assert fc.is_service_file is True

    def test_russian_content(self):
        """содержание -> service file."""
        fc = classify_file("содержание.xlsx")
        assert fc.is_service_file is True

    def test_english_cover(self):
        """Cover -> service file (case-insensitive after lower())."""
        fc = classify_file("Cover.xlsx")
        assert fc.is_service_file is True

    def test_english_toc(self):
        """toc -> service file."""
        fc = classify_file("toc.xlsx")
        assert fc.is_service_file is True

    def test_english_template(self):
        """template -> service file."""
        fc = classify_file("template.xlsx")
        assert fc.is_service_file is True

    def test_cover_in_folder(self):
        """cover in a subfolder."""
        fc = classify_file("/path/to/docs/cover.xlsx")
        assert fc.is_service_file is True
        assert fc.parent_folder == "docs"


# ═══════════════════════════════════════════════════════════════════════
#  4. classify_file — heuristic (extract_card_number_from_filepath)
# ═══════════════════════════════════════════════════════════════════════

class TestClassifyHeuristic:
    def test_sqrt_extracted(self):
        """SQRT1L-17-AS-04001 in filename -> heuristic match."""
        fc = classify_file("SQRT1L-17-AS-04001 20点扫描.xlsx")
        assert fc.is_operational_card is True
        assert fc.should_parse_parts is True
        assert fc.operation_number == "SQRT1L-17-AS-04001"

    def test_g01_as_extracted(self):
        """G01-AS-05001 in filename -> heuristic match via CARD_START_RE (returns G01)."""
        fc = classify_file("G01-AS-05001-Установка.xlsx")
        assert fc.is_operational_card is True
        # _extract_operation_number matches CARD_START_RE first (G01)
        assert fc.operation_number == "G01"

    def test_tp_number(self):
        """TP-0123[копия] -> TP matched by CARD_NUMBER_RE (stops at non-\w char)."""
        # extract_card_number_from_filepath returns "TP-0123" which != "TP-0123[копия]"
        fc = classify_file("TP-0123[копия].xlsx")
        assert fc.is_operational_card is True
        assert fc.operation_number == "TP-0123"

    def test_a123_number(self):
        """A123 -> heuristic match."""
        fc = classify_file("A123.xlsx")
        assert fc.is_operational_card is True

    def test_card_number_similar_to_name(self):
        """Card number equals filename -> falls through to alt heuristic."""
        # extract_card_number_from_filepath returns full filename when no pattern matched
        # Since card_no == file_name, it falls to alt heuristic
        fc = classify_file("процесс.xlsx")
        # "процесс" has no letter+digits pattern → unknown
        assert fc.is_operational_card is False
        assert fc.should_parse_parts is False

    def test_sqrt_in_subfolder(self):
        """SQRT card in subfolder."""
        fc = classify_file("/cards/SQRT1L-17-AS-04001.xlsx")
        assert fc.is_operational_card is True
        assert fc.parent_folder == "cards"


# ═══════════════════════════════════════════════════════════════════════
#  6. classify_file — alt heuristic (garbled filenames)
# ═══════════════════════════════════════════════════════════════════════

class TestClassifyAltHeuristic:
    def test_swm_garbled_filename(self):
        """SWM card with garbled encoding -> alt heuristic finds G01P."""
        fc = classify_file("5. G01Pш╜жщЧич║┐х╖ешЙ║хНб.xlsx")
        assert fc.is_operational_card is True
        assert fc.should_parse_parts is True
        assert fc.operation_number == "G01P"

    def test_chinese_with_card_number(self):
        """Chinese text with card number in middle."""
        fc = classify_file("测试ABC123文件.xlsx")
        assert fc.is_operational_card is True
        assert fc.operation_number == "ABC123"

    def test_number_then_stuff(self):
        """Number-prefix with card pattern in middle -> alt heuristic."""
        fc = classify_file("5. LG01总装卡.xlsx")
        assert fc.is_operational_card is True
        assert fc.operation_number == "LG01"


# ═══════════════════════════════════════════════════════════════════════
#  7. classify_file — unknown format
# ═══════════════════════════════════════════════════════════════════════

class TestClassifyUnknown:
    def test_random_text(self):
        """Random text without any pattern -> unknown."""
        fc = classify_file("просто файл.xlsx")
        assert fc.is_operational_card is False
        assert fc.is_service_file is False
        assert fc.should_parse_parts is False
        assert fc.should_split is False

    def test_service_keyword_with_number(self):
        """Service keyword with number but not matching patterns -> unknown."""
        # '记录表' is a service keyword, and there's no operation number pattern
        fc = classify_file("记录表 2024.xlsx")
        assert fc.is_service_file is True  # service keyword detected
        assert fc.is_operational_card is False
        assert fc.should_parse_parts is False

    def test_single_letter_prefix(self):
        """Single letter + 1 digit is not enough for CARD_START."""
        fc = classify_file("A1-операция.xlsx")
        assert fc.is_operational_card is False

    def test_garbled_only(self):
        """Only garbled chars without recognizable pattern."""
        fc = classify_file("ш╜жщЧич║┐х╖ешЙ║хНб.xlsx")
        assert fc.is_operational_card is False
        assert fc.should_parse_parts is False

    def test_no_extension_in_path(self):
        """Path without recognizable .xlsx extension still works (classifier checks basename)."""
        fc = classify_file("/path/to/some_file")
        assert fc.is_operational_card is False
        assert fc.should_parse_parts is False


# ═══════════════════════════════════════════════════════════════════════
#  8. _find_card_number_in_name
# ═══════════════════════════════════════════════════════════════════════

class TestFindCardNumberInName:
    def test_g01p_found(self):
        """G01P found in garbled filename."""
        assert _find_card_number_in_name("5. G01Pш╜жщЧич║┐х╖ешЙ║хНб") == "G01P"

    def test_lg01_found(self):
        """LG01 found in Chinese filename."""
        assert _find_card_number_in_name("5. LG01总装卡") == "LG01"

    def test_abc123_found(self):
        """ABC123 found in Chinese text."""
        assert _find_card_number_in_name("测试ABC123文件") == "ABC123"

    def test_multiple_candidates(self):
        """First match is returned."""
        assert _find_card_number_in_name("G01P测试ABC123") == "G01P"

    def test_no_match_empty(self):
        """No match returns empty string."""
        assert _find_card_number_in_name("просто текст") == ""

    def test_no_match_service_keyword(self):
        """Service keywords (CJK) don't match (no latin letters)."""
        assert _find_card_number_in_name("封面") == ""

    def test_no_match_cyrillic(self):
        """Cyrillic-only text doesn't match."""
        assert _find_card_number_in_name("обложка") == ""

    def test_no_match_english_word(self):
        """English word without digits doesn't match."""
        assert _find_card_number_in_name("cover") == ""

    def test_single_char_prefix(self):
        """Single letter + 2+ digits matches (e.g. G01)."""
        assert _find_card_number_in_name("G01P") == "G01P"

    def test_four_letter_prefix(self):
        """4 letters + digits matches."""
        assert _find_card_number_in_name("LG01-процесс") == "LG01"

    def test_empty_string(self):
        """Empty string returns empty."""
        assert _find_card_number_in_name("") == ""

    def test_special_chars_before_pattern(self):
        """Special chars before the pattern still finds it."""
        assert _find_card_number_in_name("...---ABC123...") == "ABC123"


# ═══════════════════════════════════════════════════════════════════════
#  9. _extract_operation_number
# ═══════════════════════════════════════════════════════════════════════

class TestExtractOperationNumber:
    def test_prefix_as_sqrt(self):
        """SQRT1L-A-AS-04001 -> AS pattern."""
        assert _extract_operation_number("SQRT1L-A-AS-04001-20点扫描") == "SQRT1L-A-AS-04001"

    def test_prefix_as_g01(self):
        """G01-A-AS-05001 -> AS pattern."""
        assert _extract_operation_number("G01-A-AS-05001-Установка") == "G01-A-AS-05001"

    def test_prefix_as_swm(self):
        """SWM-A-AS-001 -> AS pattern."""
        assert _extract_operation_number("SWM-A-AS-001") == "SWM-A-AS-001"

    def test_card_start_a001(self):
        """A001 -> CARD_START_RE."""
        assert _extract_operation_number("A001-Контроль") == "A001"

    def test_card_start_t1l_not_matched(self):
        """T1L has only 1 digit -> CARD_START_RE does NOT match (needs 2+ digits)."""
        assert _extract_operation_number("T1L作业指导书") == ""

    def test_card_start_lg01(self):
        """LG01 -> CARD_START_RE (4 letters + 2 digits)."""
        assert _extract_operation_number("LG01-операция") == "LG01"

    def test_card_start_tp01(self):
        """TP01 -> CARD_START_RE (2 letters + 2 digits)."""
        assert _extract_operation_number("TP01-операция") == "TP01"

    def test_digit_start_038(self):
        """038 -> DIGIT_START_RE."""
        assert _extract_operation_number("038-Установка двери") == "038"

    def test_digit_start_1234(self):
        """1234 -> DIGIT_START_RE."""
        assert _extract_operation_number("1234-операция") == "1234"

    def test_no_match_random(self):
        """Random text -> no match."""
        assert _extract_operation_number("просто текст") == ""

    def test_no_match_single_digit(self):
        """Single digit at start -> not enough (needs 2+)."""
        assert _extract_operation_number("5-операция") == ""

    def test_no_match_cover(self):
        """cover -> no numbers."""
        assert _extract_operation_number("cover") == ""

    def test_empty_string(self):
        """Empty string -> no match."""
        assert _extract_operation_number("") == ""


# ═══════════════════════════════════════════════════════════════════════
#  10. filter_operational_cards, get_parseable_files, get_splittable_files
# ═══════════════════════════════════════════════════════════════════════

class TestFilterFunctions:
    def test_filter_operational_cards_mixed(self):
        """Filter classifies a mix of operational/service files."""
        paths = [
            "038-операция.xlsx",
            "封面.xlsx",
            "SQRT1L-A-AS-04001.xlsx",
        ]
        results = filter_operational_cards(paths)
        assert len(results) == 3
        assert results[0].is_operational_card is True
        assert results[1].is_operational_card is False
        assert results[2].is_operational_card is True

    def test_filter_operational_cards_empty(self):
        """Empty list returns empty list."""
        assert filter_operational_cards([]) == []

    def test_get_parseable_files(self):
        """Filter classifications for parseable files."""
        classifications = [
            FileClassification("op1.xlsx", "op1", ".", is_operational_card=True, is_service_file=False, should_split=True, should_parse_parts=True),
            FileClassification("service.xlsx", "service", ".", is_operational_card=False, is_service_file=True, should_split=False, should_parse_parts=False),
            FileClassification("op2.xlsx", "op2", ".", is_operational_card=True, is_service_file=False, should_split=True, should_parse_parts=True),
        ]
        parseable = get_parseable_files(classifications)
        assert parseable == ["op1.xlsx", "op2.xlsx"]

    def test_get_parseable_files_empty(self):
        """Empty list returns empty list."""
        assert get_parseable_files([]) == []

    def test_get_splittable_files(self):
        """Filter classifications for splittable files."""
        classifications = [
            FileClassification("op.xlsx", "op", ".", is_operational_card=True, is_service_file=False, should_split=True, should_parse_parts=True),
            FileClassification("service.xlsx", "service", ".", is_operational_card=False, is_service_file=True, should_split=False, should_parse_parts=False),
            FileClassification("unknown.xlsx", "unknown", ".", is_operational_card=False, is_service_file=False, should_split=True, should_parse_parts=False),
        ]
        splittable = get_splittable_files(classifications)
        assert len(splittable) == 2
        assert splittable[0].file_path == "op.xlsx"
        assert splittable[1].file_path == "unknown.xlsx"

    def test_get_splittable_files_returns_classifications(self):
        """get_splittable_files returns FileClassification objects, not paths."""
        classifications = [
            FileClassification("op.xlsx", "op", ".", is_operational_card=True, is_service_file=False, should_split=True, should_parse_parts=True),
        ]
        result = get_splittable_files(classifications)
        assert isinstance(result[0], FileClassification)
        assert result[0].file_name == "op"

    def test_get_splittable_files_empty(self):
        """Empty list returns empty list."""
        assert get_splittable_files([]) == []


# ═══════════════════════════════════════════════════════════════════════
#  11. _contains_service_keywords (direct tests)
# ═══════════════════════════════════════════════════════════════════════

class TestContainsServiceKeywords:
    def test_chinese_cover(self):
        assert _contains_service_keywords("封面") is True

    def test_chinese_catalog(self):
        assert _contains_service_keywords("目录") is True

    def test_english_cover(self):
        assert _contains_service_keywords("cover") is True

    def test_english_template(self):
        assert _contains_service_keywords("template") is True

    def test_english_toc(self):
        assert _contains_service_keywords("toc") is True

    def test_operational_name(self):
        """Operational card names are not service keywords."""
        assert _contains_service_keywords("038-установка") is False

    def test_empty_string(self):
        assert _contains_service_keywords("") is False

    def test_card_number_not_service(self):
        assert _contains_service_keywords("5306200-ED001") is False

    def test_russian_oblozhka(self):
        assert _contains_service_keywords("обложка") is True

    def test_russian_soderzhanie(self):
        assert _contains_service_keywords("содержание") is True

    def test_mixed_with_service_keyword(self):
        """Filename containing a keyword should be detected."""
        assert _contains_service_keywords("document_cover_page") is True
