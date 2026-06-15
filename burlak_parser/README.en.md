# Burlak Parser

BOM parsing and reconciliation with assembly operation cards for automotive manufacturing.

In-memory data loading (load_from_bytes / load_async) for HTTP server integration. Detailed split statistics (SplitStatistics) with skip reasons. verify_integrity() function for result validation. Context manager support for services.

All configurations processed in a single run, heuristic column detection in three languages, safe fuzzy matching, ZIP-based splitting with full formatting preservation.

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
| `--split-stats` | `-S` | Detailed split statistics |
| `--verbose` | `-v` | Debug log |

### Output

1. `report.txt`
2. `discrepancies.xlsx` (4 sheets)
3. `split_cards/`
4. `split_cards.zip`

---

## In-Memory API

Services support loading from byte streams without disk writes:

```python
from burlak_parser.bom_parser import BOMService
from burlak_parser.card_parser import CardService

# Synchronous
service = BOMService()
data = service.load_from_bytes(bom_bytes)

# Async
data = await service.load_async(bom_bytes)

# Context manager
async with CardService() as service:
    data = await service.load_async(cards_bytes)
```

## SplitStatistics

`CardSplitter.split_many_parallel()` returns `(files, errors)`. Each file has `FileSplitStats`: sheet count, split count, skip reasons (template, empty, service), file size, format.

`SplitStatistics` aggregates: top skip reasons, top files by skips.

## verify_integrity()

Checks integrity of comparison results: every BOM part has a result, quantities are consistent, no duplicates. Returns `IntegrityCheck` with passed/failed status.

## Notes

- Template sheets with data are processed (not skipped)
- Service files are not split
- Original part-number format (with dashes) is preserved in reports

---

## Architecture

```
burlak_parser/
├── __init__.py
├── main.py                 # CLI, --split-stats
├── bom_parser.py           # BOMService, load_from_bytes, load_async
├── card_parser.py          # CardService, load_from_bytes, load_async
├── file_classifier.py      # FileClassifier
├── fuzzy_matcher.py        # FuzzyMatcher
├── splitter.py             # CardSplitter, SplitStatistics
├── comparator.py           # MatchingEngine
├── report_generator.py     # Reporter
└── heuristic_analyzer.py   # HeuristicAnalyzer
tests/
├── __init__.py
├── conftest.py
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
