# Burlak Parser

BOM parsing and reconciliation with assembly operation cards.

BOM files may contain both real configurations (numeric quantities) and VIN columns (S / - values). The parser automatically detects the boundary and excludes VIN breakdown, keeping only actual configurations.

---

## Install

```
python3 -m venv venv && source venv/bin/activate
pip install -r requirements.txt
pip install xlrd
```

Dependencies: openpyxl, pandas, xlsxwriter, tqdm.

---

## Usage

```
python -m burlak_parser.main --bom "BOM.xlsx" --cards "./cards/" [--config <name>]
```

### BOM Parsing

The parser finds all columns to the right of the part number and checks their content for numeric values. Columns containing only `S` (Same) and `-` (not applicable) — VIN breakdown — are automatically excluded. Only real configurations with part quantities remain.

---

## Structure

```
burlak_parser/
├── __init__.py
├── main.py
├── bom_parser.py        # Auto-exclude VIN columns
├── card_parser.py       # ZIP splitting, improved empty-sheet detection
├── comparator.py
└── report_generator.py
```
