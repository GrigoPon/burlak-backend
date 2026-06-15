# Burlak Parser

Bill-of-Materials (BOM) parsing and reconciliation with assembly operation cards.

Cross-references catalog part numbers from a BOM file against part numbers in operation cards (xlsx/xls). Identifies discrepancies: parts missing from cards, extra parts in cards, quantity mismatches.

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

### Options

| Parameter | Description |
|-----------|-------------|
| `--bom` | Path to BOM file (.xlsx) |
| `--cards` | Path to cards folder or ZIP archive |
| `--config` | Configuration name (interactive if omitted) |

### Output

Artifacts in `./output/`:
1. `report.txt` — text discrepancy report
2. `discrepancy_report.xlsx` — Excel report (3 sheets: summary, discrepancies, all BOM parts)
3. `split_cards/` — split single-sheet files
4. `split_cards.zip` — archive of split files

---

## How it works

### Step 1: BOM Parsing

`bom_parser.py` reads the xlsx specification, finds the header row (零件号, PartNo), identifies columns: part number, name (CN), name (EN), configuration columns.

### Step 2: Card Parsing

`card_parser.py` scans files in the folder or ZIP archive. For each file: identifies card number, finds the parts table by keywords (料号, 零件号, 用量, qty), extracts part number, quantity, name. Merges multi-line part numbers. Aggregates duplicates.

### Step 3: Sheet Splitting

Each multi-sheet xlsx is converted into individual files (one sheet = one file). Empty sheets are skipped. xls files are not split.

### Step 4: Comparison

Three discrepancy types: only in BOM, only in cards, quantity mismatch.

### Step 5: Report

Text report + Excel report with color-coded discrepancies + ZIP archive.

---

## Structure

```
burlak_parser/
├── __init__.py
├── main.py
├── bom_parser.py
├── card_parser.py
├── comparator.py
└── report_generator.py
```
