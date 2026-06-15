# Burlak Parser

BOM parsing and reconciliation with assembly operation cards for automotive manufacturing.

Original part-number format (with dashes) preserved in comparison results. Detailed split statistics. Integrity verification. Expanded test suite with conftest.py and automatic temp file cleanup.

---

## Install

```
python3 -m venv venv && source venv/bin/activate
pip install -r requirements.txt
pip install xlrd
```

Dependencies: openpyxl, xlsxwriter, tqdm.

---

## Testing

```
pip install pytest pytest-mock
pytest tests/
```

---

## Usage

```
python -m burlak_parser.main --bom "BOM.xlsx" --cards "./cards/" [OPTIONS]
```

---

## Original Part-Number Format

`CardsData.original_part_numbers` tracks the original (dashed) format alongside the normalized version. All discrepancy types use the original BOM part number — `1234-56-78` stays as-is in reports.

## Split Statistics

After splitting, console output shows:
- Operational card count, format breakdown (xlsx/xls)
- Average sheets per file
- Service files skipped

## Integrity Verification

Post-comparison check: every BOM part has a result entry, quantities are consistent, no duplicate records.

## Path Normalization

`_collect_related_files()` normalizes absolute and relative paths in ZIP archives for correct collection of related files (sharedStrings, styles, drawings).

---

## Tests

```
tests/
├── __init__.py
├── conftest.py           # Auto cleanup fixture
├── test_bom_parser.py
├── test_card_parser.py
├── test_comparator.py
├── test_fuzzy_matcher.py
├── test_heuristic_analyzer.py
├── test_file_classifier.py
├── test_main.py
├── test_report_generator.py
└── test_splitter.py      # 12 test classes
```

---

## Architecture

```
burlak_parser/
├── __init__.py
├── main.py
├── bom_parser.py
├── card_parser.py         # Original part numbers
├── file_classifier.py
├── fuzzy_matcher.py
├── splitter.py            # Split statistics
├── comparator.py          # Original numbers in discrepancies
├── report_generator.py
└── heuristic_analyzer.py
```
