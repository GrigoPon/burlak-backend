# Burlak Parser

BOM parsing and reconciliation with assembly operation cards for automotive manufacturing.

All configurations processed in a single run. File classification (operation cards / service files), safe fuzzy part-number matching, ZIP-based sheet splitting with full formatting preservation, parallel execution.

Modular service-class architecture, ready for both CLI and programmatic use.

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

### Options

| Parameter | Short | Description |
|-----------|-------|-------------|
| `--bom` | `-b` | Path to BOM file (.xlsx) |
| `--cards` | `-c` | Cards folder or ZIP |
| `--output` | `-o` | Output directory |
| `--single-config` | `-s` | Single config mode |
| `--config` | `-k` | Config name |
| `--no-split` | | Don't split cards |
| `--no-fuzzy` | | Disable fuzzy matching |
| `--workers` | `-w` | Process count |
| `--verbose` | `-v` | Debug log |

### Modes

- **Default** — all configurations, full discrepancy matrix
- `--single-config` — single configuration

### Output

1. `report.txt`
2. `discrepancies.xlsx` (4 sheets)
3. `split_cards/`
4. `split_cards.zip`

---

## Architecture

```
burlak_parser/
├── __init__.py
├── main.py              # CLI, pipeline
├── bom_parser.py        # BOMService
├── card_parser.py       # CardService
├── file_classifier.py   # FileClassifier
├── fuzzy_matcher.py     # FuzzyMatcher
├── splitter.py          # CardSplitter
├── comparator.py        # MatchingEngine
└── report_generator.py  # Reporter
```

### Services

| Service | Purpose |
|---------|---------|
| BOMService | Load and parse BOM files |
| CardService | Load and parse operation cards |
| FileClassifier | Classify: operation card or service file |
| FuzzyMatcher | Safe fuzzy part-number comparison |
| CardSplitter | ZIP-based multi-sheet splitting |
| MatchingEngine | BOM vs card reconciliation |
| Reporter | Report generation (XLSX, TXT, ZIP) |

---

## How It Works

### File Classification

Each file is classified before parsing:
- **Operation cards** — name starts with operation number (038-Установка, A123-Контроль, SQRT1L-A-AS-038)
- **Service files** — name contains 封面, 目录, 记录表, 空表

### BOM Parsing

Automatic column detection and VIN filtering. openpyxl with read_only=True.

### Card Parsing

Universal loader (xlsx / xls). Part table search, multi-line number merging, aggregation. Parallel via ProcessPoolExecutor.

### ZIP Splitting

"Remove unwanted" method: copy source, remove all sheets except target from ZIP structure. 100% formatting preservation. Named range cleanup to prevent Excel "Removed Feature" error.

### Fuzzy Matching

Two numbers match if identical after removing hyphens, spaces, dots, special chars. Difference in any digit or letter = no match.

### Multi-Config Reconciliation

Global fuzzy index, per-config discrepancy matrix: only in BOM, only in cards, quantity mismatch, fuzzy match.

### Report

4 sheets: summary, per-config, discrepancies, fuzzy matches.
