# Burlak Parser

BOM parsing and reconciliation with assembly operation cards for automotive manufacturing.

Heuristic document structure analyzer for automatic column detection (part number, quantity, name) in three languages: Russian, English, Chinese. Enables processing BOMs from different vendors without manual configuration.

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

---

## Heuristic Analyzer

`heuristic_analyzer.py` dynamically determines Excel document structure without hardcoded column indices.

### Column Detection

Scans headers (first rows) and matches against keywords:

| Category | RU | EN | CN |
|----------|-----|-----|------|
| Part number | номер, артикул, код, деталь, обозначение | part no, part number, p/n, code | 件号, 图号, 零件号, 代号 |
| Quantity | кол-во, количество | qty, quantity | 数量, 用量 |
| Name | наименование, название, описание | name, description, part name | 名称, 描述 |

### Algorithm

1. `find_header_rows` — find header row by keywords
2. `detect_column_types` — classify columns (part_no, name, qty, config, meta)
3. `find_data_region` — determine data table boundaries
4. Fallback to positional heuristics if headers not found
5. Return column configuration with confidence score

---

## Architecture

```
burlak_parser/
├── __init__.py
├── main.py
├── bom_parser.py
├── card_parser.py
├── file_classifier.py
├── fuzzy_matcher.py
├── splitter.py
├── comparator.py
├── report_generator.py
└── heuristic_analyzer.py   # Heuristic analyzer
```
