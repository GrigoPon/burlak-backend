# Burlak Parser

BOM parsing and reconciliation with assembly operation cards for automotive manufacturing.

Nested ZIP archive support with recursive extraction. File deduplication by signature (size + content). Professional dashboard in Excel report with metric cards.

---

## Install

```
python3 -m venv venv && source venv/bin/activate
pip install -r requirements.txt
pip install xlrd
```

Dependencies: openpyxl, xlsxwriter, tqdm.

---

## Usage

```
python -m burlak_parser.main --bom "BOM.xlsx" --cards "./cards/" [OPTIONS]
```

### File Discovery

`_find_excel_files()`:
- Recursively extracts ZIP archives, including nested ones
- Deduplicates files by signature (file size + first 4096 bytes)
- Filters temp files (~$) and non-Excel formats
- Extracted files go to a temp directory, cleaned up on completion

### Report Dashboard

The Summary sheet is redesigned as a professional dashboard:
- Header with metadata (configs, parts in BOM and cards)
- Metric cards: total discrepancies, quantity mismatch, only in BOM, only in cards
- Processed/corrupted file count
- Config table with alternating row colors

---

## Architecture

```
burlak_parser/
├── __init__.py
├── main.py
├── bom_parser.py
├── card_parser.py       # Nested ZIPs, signature dedup
├── file_classifier.py
├── fuzzy_matcher.py
├── splitter.py
├── comparator.py
└── report_generator.py  # Dashboard
```
