# Burlak Parser

BOM parsing and reconciliation with assembly operation cards.

Multi-sheet xlsx card files are split into individual sheets with full formatting preservation — images, styles, merged cells, page setup — via direct ZIP structure manipulation.

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
| `--cards` | Cards folder or ZIP archive |
| `--config` | Configuration name |
| `--output` | Output directory (default: `./output`) |
| `--no-split` | Skip multi-sheet splitting |
| `--verbose` | Debug output |

### Output

1. `report.txt`
2. `discrepancy_report.xlsx` (3 sheets)
3. `split_cards/` — split single-sheet files (formatting preserved)
4. `split_cards.zip`

---

## Sheet Splitting

Two methods:

**Primary (ZIP manipulation).** xlsx is a ZIP archive of XML files. The method copies the source file byte-by-byte, then removes all sheets except the target from the ZIP structure. Preserves 100% of original formatting: fonts, colors, borders, fills, alignment, merged cells, column widths, row heights, freeze panes, images, page setup.

**Fallback (openpyxl).** If the ZIP method fails, creates a new Workbook and deep-copies cell styles.

---

## Structure

```
burlak_parser/
├── __init__.py
├── main.py
├── bom_parser.py
├── card_parser.py       # ZIP extraction, deep style copy
├── comparator.py
└── report_generator.py
```
