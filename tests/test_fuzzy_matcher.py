"""Unit-тесты для fuzzy_matcher.py.

Покрытие:
  - normalize_part_number: дефисы, пробелы, спецсимволы, uppercase, точки
  - is_fuzzy_match: точное совпадение, fuzzy, разные номера, пустые строки
  - is_valid_part_number: валидные PN, мусор, короткие, только буквы
  - FuzzyMatcher: индекс, find_fuzzy_match, get_normalized, пустой набор
"""

from __future__ import annotations

import pytest

from burlak_parser.fuzzy_matcher import (
    FuzzyMatcher,
    is_fuzzy_match,
    is_valid_part_number,
    normalize_part_number,
)


# ═══════════════════════════════════════════════════════════════════════
#  1. normalize_part_number
# ═══════════════════════════════════════════════════════════════════════

class TestNormalizePartNumber:
    def test_removes_dashes(self):
        """Regular dash should be removed."""
        assert normalize_part_number("5306200-ED001") == "5306200ED001"

    def test_removes_spaces(self):
        """Spaces should be removed."""
        assert normalize_part_number("ABC 123") == "ABC123"

    def test_removes_dots(self):
        """Dots should be removed."""
        assert normalize_part_number("ABC.1.2.3") == "ABC123"

    def test_removes_em_dash(self):
        """Em-dash should be removed."""
        assert normalize_part_number("ABC—123") == "ABC123"

    def test_removes_slashes(self):
        """Slashes should be removed."""
        assert normalize_part_number("ABC/123/DEF") == "ABC123DEF"

    def test_removes_underscores(self):
        """Underscores should be removed."""
        assert normalize_part_number("ABC_123_DEF") == "ABC123DEF"

    def test_uppercases(self):
        """Result should be uppercase."""
        assert normalize_part_number("abc-123") == "ABC123"

    def test_multiple_dashes(self):
        """Multiple dashes should all be removed."""
        assert normalize_part_number("A-B-C---123") == "ABC123"

    def test_already_normalized(self):
        """Already normalized string stays unchanged."""
        assert normalize_part_number("ABC123DEF") == "ABC123DEF"

    def test_with_parentheses(self):
        """Parentheses should be removed."""
        assert normalize_part_number("ABC(123)DEF") == "ABC123DEF"

    def test_mixed_special_chars(self):
        """Mix of all special chars should be cleaned."""
        assert normalize_part_number("A-B C.D/E_F") == "ABCDEF"

    def test_leading_trailing_whitespace(self):
        """Leading/trailing whitespace should be trimmed."""
        assert normalize_part_number("  ABC-123  ") == "ABC123"

    def test_with_comma_and_semicolon(self):
        """Commas and semicolons should be removed."""
        assert normalize_part_number("ABC,123;DEF") == "ABC123DEF"

    def test_empty_string(self):
        """Empty string returns empty string."""
        assert normalize_part_number("") == ""

    def test_only_special_chars(self):
        """Only special chars returns empty string."""
        assert normalize_part_number("--..  ..--") == ""


# ═══════════════════════════════════════════════════════════════════════
#  2. is_fuzzy_match
# ═══════════════════════════════════════════════════════════════════════

class TestIsFuzzyMatch:
    def test_exact_match(self):
        """Exact same string is a match."""
        assert is_fuzzy_match("5306200ED001", "5306200ED001") is True

    def test_dash_vs_no_dash(self):
        """Same number with/without dashes is a fuzzy match."""
        assert is_fuzzy_match("5306200-ED001", "5306200ED001") is True

    def test_space_vs_dash(self):
        """Space vs dash is a fuzzy match."""
        assert is_fuzzy_match("ABC 123", "ABC-123") is True

    def test_different_case(self):
        """Different case should still match (uppercased)."""
        assert is_fuzzy_match("abc-123", "ABC-123") is True

    def test_different_digit_fails(self):
        """Different digit should NOT match, even if similar."""
        assert is_fuzzy_match("ABCD123", "ABCD124") is False

    def test_different_length_fails(self):
        """Different length should NOT match after normalization."""
        assert is_fuzzy_match("ABCD-123", "ABCD-1234") is False

    def test_extra_letter_fails(self):
        """Different letter should NOT match."""
        assert is_fuzzy_match("ABCD-123", "ABCE-123") is False

    def test_em_dash_vs_regular_dash(self):
        """Em-dash vs regular dash is a match (both removed)."""
        assert is_fuzzy_match("ABC—123", "ABC-123") is True

    def test_dot_vs_dash(self):
        """Dot vs dash is a match (both removed)."""
        assert is_fuzzy_match("ABC.123", "ABC-123") is True

    def test_empty_strings(self):
        """Two empty strings match."""
        assert is_fuzzy_match("", "") is True

    def test_one_empty_string(self):
        """One empty string does NOT match a non-empty one after normalization."""
        assert is_fuzzy_match("", "ABC") is False

    def test_special_chars_vs_none(self):
        """All-special vs all-special matches (both become empty)."""
        assert is_fuzzy_match("--..", "..--") is True


# ═══════════════════════════════════════════════════════════════════════
#  3. is_valid_part_number
# ═══════════════════════════════════════════════════════════════════════

class TestIsValidPartNumber:
    def test_letters_and_digits(self):
        """Standard part number with letters and digits is valid."""
        assert is_valid_part_number("ABC123DEF") is True

    def test_with_dashes(self):
        """Part number with dashes is valid (dashes kept for pattern match)."""
        assert is_valid_part_number("5306200-ED001") is True

    def test_numeric_only(self):
        """Numeric-only string >= 3 chars is valid."""
        assert is_valid_part_number("12345") is True

    def test_short_string_invalid(self):
        """String shorter than 3 chars is invalid."""
        assert is_valid_part_number("AB") is False

    def test_empty_string_invalid(self):
        """Empty string is invalid."""
        assert is_valid_part_number("") is False

    def test_none_invalid(self):
        """None should be treated as invalid."""
        assert is_valid_part_number(None) is False  # type: ignore

    def test_na_garbage(self):
        """'N/A' is garbage and should be invalid."""
        assert is_valid_part_number("N/A") is False

    def test_none_word_garbage(self):
        """'None' is garbage."""
        assert is_valid_part_number("None") is False

    def test_dash_garbage(self):
        """Just a dash is garbage."""
        assert is_valid_part_number("—") is False

    def test_chinese_garbage(self):
        """Chinese text '无' (nothing) is garbage."""
        assert is_valid_part_number("无") is False

    def test_numeric_3_digits(self):
        """3-digit number is valid (minimum length)."""
        assert is_valid_part_number("123") is True

    def test_letters_only_valid_if_long(self):
        """Letters only is valid in fuzzy_matcher (requires at least 3 chars + letter)."""
        assert is_valid_part_number("ABCDEF") is True

    def test_mixed_case(self):
        """Mixed case is still valid."""
        assert is_valid_part_number("Abc-123-Def") is True

    def test_with_special_chars(self):
        """Special chars like dots and slashes are acceptable in PN."""
        assert is_valid_part_number("ABC/123.DEF") is True

    def test_too_short_numeric(self):
        """2-digit number is too short."""
        assert is_valid_part_number("12") is False

    def test_null_string(self):
        """'null' is garbage."""
        assert is_valid_part_number("null") is False

    def test_whitespace_only(self):
        """Whitespace-only string should be invalid after strip."""
        assert is_valid_part_number("   ") is False


# ═══════════════════════════════════════════════════════════════════════
#  4. FuzzyMatcher — basic
# ═══════════════════════════════════════════════════════════════════════

class TestFuzzyMatcherBasic:
    def test_single_part(self):
        """Single part in BOM set."""
        fm = FuzzyMatcher({"5306200-ED001"})
        assert fm.find_fuzzy_match("5306200ED001") == "5306200-ED001"

    def test_exact_match(self):
        """Exact match returns the same part."""
        fm = FuzzyMatcher({"ABC123"})
        assert fm.find_fuzzy_match("ABC123") == "ABC123"

    def test_no_match(self):
        """No match returns None."""
        fm = FuzzyMatcher({"ABC123"})
        assert fm.find_fuzzy_match("XYZ999") is None

    def test_normalize_retrieval(self):
        """get_normalized returns normalized form."""
        fm = FuzzyMatcher({"ABC-123"})
        assert fm.get_normalized("ABC-123") == "ABC123"

    def test_multiple_parts(self):
        """Multiple parts in BOM set."""
        fm = FuzzyMatcher({"P001", "5306200-ED001", "ABC 123"})
        assert fm.find_fuzzy_match("5306200ED001") == "5306200-ED001"
        assert fm.find_fuzzy_match("ABC-123") == "ABC 123"
        assert fm.find_fuzzy_match("P001") == "P001"

    def test_empty_set(self):
        """Empty BOM set returns None for all."""
        fm = FuzzyMatcher(set())
        assert fm.find_fuzzy_match("ABC123") is None


# ═══════════════════════════════════════════════════════════════════════
#  5. FuzzyMatcher — edge cases
# ═══════════════════════════════════════════════════════════════════════

class TestFuzzyMatcherEdgeCases:
    def test_duplicate_normalized_forms(self):
        """Two different BOM parts with same normalized form: returns first one."""
        fm = FuzzyMatcher({"5306200-ED001", "5306200ED001"})
        # Both normalize to "5306200ED001"
        result = fm.find_fuzzy_match("5306200-ED001")
        assert result is not None
        # Should return one of the two (order depends on set iteration)
        assert result in ("5306200-ED001", "5306200ED001")

    def test_case_insensitive_match(self):
        """Match should be case-insensitive."""
        fm = FuzzyMatcher({"ABC-123"})
        assert fm.find_fuzzy_match("abc-123") == "ABC-123"

    def test_spaces_in_bom_part(self):
        """BOM parts with spaces matched by card parts with dashes."""
        fm = FuzzyMatcher({"ABC 123"})
        assert fm.find_fuzzy_match("ABC-123") == "ABC 123"

    def test_dots_in_bom_part(self):
        """BOM parts with dots matched by card parts with dashes."""
        fm = FuzzyMatcher({"ABC.123"})
        assert fm.find_fuzzy_match("ABC-123") == "ABC.123"

    def test_get_normalized_for_unknown(self):
        """get_normalized works even for parts not in index."""
        fm = FuzzyMatcher({"ABC123"})
        assert fm.get_normalized("DEF-456") == "DEF456"

    def test_long_part_numbers(self):
        """Long part numbers are handled correctly."""
        long_pn = "A" * 50 + "-123"
        fm = FuzzyMatcher({long_pn})
        assert fm.find_fuzzy_match(long_pn.replace("-", "")) == long_pn

    def test_unicode_chars_in_part(self):
        """Part numbers with unicode dashes should match regular dashes."""
        fm = FuzzyMatcher({"ABC—123"})  # em-dash
        assert fm.find_fuzzy_match("ABC-123") == "ABC—123"  # regular dash


# ═══════════════════════════════════════════════════════════════════════
#  6. FuzzyMatcher — integration with real-world scenarios
# ═══════════════════════════════════════════════════════════════════════

class TestFuzzyMatcherRealWorld:
    def test_t1l_part_number(self):
        """T1L-style: part with dash in BOM matches clean version from cards."""
        fm = FuzzyMatcher({"5306200-ED001", "551002664AA", "Q146Z0825F36"})
        assert fm.find_fuzzy_match("5306200ED001") == "5306200-ED001"
        assert fm.find_fuzzy_match("551002664AA") == "551002664AA"
        assert fm.find_fuzzy_match("Q146Z0825F36") == "Q146Z0825F36"

    def test_swm_part_number(self):
        """SWM-style: long part numbers with dashes."""
        fm = FuzzyMatcher({"4007100-ED002-AA00000"})
        assert fm.find_fuzzy_match("4007100ED002AA00000") == "4007100-ED002-AA00000"
        assert fm.find_fuzzy_match("4007100-ED002-AA00000") == "4007100-ED002-AA00000"

    def test_multiple_dashes_in_bom(self):
        """Part with multiple dashes in BOM."""
        fm = FuzzyMatcher({"A-B-C-D-123"})
        assert fm.find_fuzzy_match("ABCD123") == "A-B-C-D-123"

    def test_no_false_match_on_similar_parts(self):
        """Similar but different parts should NOT match."""
        fm = FuzzyMatcher({"ABCD123", "ABCD124"})
        assert fm.find_fuzzy_match("ABCD125") is None
        assert fm.find_fuzzy_match("ABCD124") == "ABCD124"
        assert fm.find_fuzzy_match("ABCD123") == "ABCD123"
