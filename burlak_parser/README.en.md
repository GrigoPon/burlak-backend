# Burlak Parser

BOM parsing and reconciliation with assembly operation cards for automotive manufacturing.

Filename-based deduplication. Nested ZIP archives are processed after main files to prevent duplicate insertion before originals.

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

Directory and ZIP traversal algorithm:
1. Collect all main files (from root directory and top-level ZIP) without dedup
2. Process nested ZIP archives — files are checked for duplicate names against already-collected main files
3. Temp files (~$) and non-Excel formats are discarded
4. Temp directories are cleaned up on completion

---

## Architecture

```
burlak_parser/
├── __init__.py
├── main.py
├── bom_parser.py
├── card_parser.py       # Filename-based dedup
├── file_classifier.py
├── fuzzy_matcher.py
├── splitter.py
├── comparator.py
└── report_generator.py
```
