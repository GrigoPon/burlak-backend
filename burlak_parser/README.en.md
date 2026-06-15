# Burlak Parser

BOM parsing and reconciliation with assembly operation cards.

Fuzzy part-number matching via difflib.SequenceMatcher. Template sheet filtering (cover, TOC, record sheets, blank forms). Multi-sheet file splitting via ProcessPoolExecutor.

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
python -m burlak_parser.main --bom "BOM.xlsx" --cards "./cards/" [OPTIONS]
```

### Options

| Parameter | Description | Default |
|-----------|-------------|---------|
| `--bom` | Path to BOM file (.xlsx) | required |
| `--cards` | Cards folder or ZIP | required |
| `--config` | Configuration name | interactive |
| `--output` | Output directory | `./output` |
| `--skip-templates` | Filter template sheets | True |
| `--no-skip-templates` | Include templates | — |
| `--no-fuzzy` | Disable fuzzy matching | — |
| `--no-split` | Skip splitting | — |
| `--verbose` | Debug output | — |

---

## Fuzzy Matching

Uses difflib.SequenceMatcher:
- **Normalization:** remove hyphens, spaces, dots, slashes, special chars
- **Comparison:** ratio() of two normalized strings
- **Types:** FUZZY_IN_BOM, FUZZY_IN_CARDS
- Controlled by `--no-fuzzy` flag

## Template Filtering

`_is_template_sheet()` detects service sheets by keywords: 空表, 封面, 目录, 记录表, Sheet1. Controlled by `--skip-templates`.

## Parallel Splitting

ProcessPoolExecutor for multi-sheet file splitting. `_split_single_sheet()` is a picklable top-level function.

---

## Structure

```
burlak_parser/
├── __init__.py
├── main.py
├── bom_parser.py
├── card_parser.py       # Template filtering, ProcessPoolExecutor
├── comparator.py        # Fuzzy matching
└── report_generator.py  # Fuzzy column
```
