# Burlak Parser

BOM parsing and reconciliation with assembly operation cards for automotive manufacturing.

Heuristic document structure analyzer for automatic column detection in three languages.

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
└── heuristic_analyzer.py
```
