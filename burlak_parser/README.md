# Burlak Parser

Система разбора ведомостей материалов (BOM) и сверки с операционными картами сборки для автомобильного производства.

Эвристический анализатор структуры документа для автоматического определения колонок по заголовкам на трёх языках.

---

## Установка

```
python3 -m venv venv && source venv/bin/activate
pip install -r requirements.txt
pip install xlrd
```

Зависимости: openpyxl, xlsxwriter, tqdm.

---

## Использование

```
python -m burlak_parser.main --bom "BOM.xlsx" --cards "./cards/" [OPTIONS]
```

### Параметры

| Параметр | Описание |
|----------|----------|
| `--bom` | Путь к BOM-файлу (.xlsx) |
| `--cards` | Папка/ZIP с операционными картами |
| `--output` | Директория результатов |
| `--single-config` | Режим одной комплектации |
| `--config` | Название комплектации |
| `--no-split` | Не разделять карты |
| `--no-fuzzy` | Отключить нечёткое сравнение |
| `--workers` | Количество процессов |
| `--verbose` | Подробный лог |

---

## Архитектура

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
