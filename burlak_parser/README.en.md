# Burlak Parser

BOM parsing and reconciliation with assembly operation cards for automotive manufacturing.

Multi-table parsing: section boundary detection within a single sheet — each table parsed separately. Garbled filename handling for encoding-damaged files. Full module test suite (7 files, all modules covered).

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

## Multi-Table Parsing

`_collect_all_tables()` detects section boundaries within a sheet. If a sheet contains multiple tables (e.g., multiple operations), each is parsed separately. Boundary detected by:
- New column header appearance (keywords 件号, part no, etc.)
- Three or more consecutive empty rows

## Garbled Filenames

Some files have encoding-damaged names (e.g., `5. G01Pш╜ж▓W5.xlsx`). `CARD_NUMBER_ANYWHERE_RE` finds the card number in any part of the filename, ignoring unreadable characters.

## Heuristic Analyzer Improvements

- Added "серийный", "serial" to anti-keywords
- Confidence threshold lowered to 0.25
- Scan width limited to 40 columns
- `find_part_table()` supports start_row, scans above and below header

---

## Tests

```
tests/
├── __init__.py
├── test_bom_parser.py
├── test_card_parser.py
├── test_comparator.py
├── test_fuzzy_matcher.py
├── test_heuristic_analyzer.py
├── test_file_classifier.py
├── test_main.py
├── test_report_generator.py
└── test_splitter.py
```

---

## Architecture

```
burlak_parser/
├── __init__.py
├── main.py
├── bom_parser.py
├── card_parser.py         # Multi-table parsing
├── file_classifier.py     # Garbled filenames
├── fuzzy_matcher.py
├── splitter.py
├── comparator.py
├── report_generator.py
└── heuristic_analyzer.py  # Improved heuristics
```
