# Burlak Parser

Система разбора ведомостей материалов (BOM) и сверки с операционными картами сборки для автомобильного производства.

Загрузка данных из байтового потока (load_from_bytes) и асинхронная загрузка (load_async) для интеграции с HTTP-серверами. Детальная статистика разделения (SplitStatistics) с причинами пропуска листов. Функция verify_integrity() для верификации целостности результатов. Context manager для сервисов.

Обработка всех комплектаций за один запуск, эвристическое определение колонок на трёх языках, безопасное нечёткое сравнение номеров, ZIP-разделение с полным сохранением форматирования.

---

## Установка

```
python3 -m venv venv && source venv/bin/activate
pip install -r requirements.txt
pip install xlrd
```

Зависимости: openpyxl, xlsxwriter, tqdm.

---

## Тестирование

```
pip install pytest pytest-mock
pytest tests/
```

---

## Использование

```
python -m burlak_parser.main --bom "BOM.xlsx" --cards "./cards/" [OPTIONS]
```

### Параметры

| Параметр | Кратко | Описание |
|----------|--------|----------|
| `--bom` | `-b` | Путь к BOM-файлу (.xlsx) |
| `--cards` | `-c` | Папка/ZIP с операционными картами |
| `--output` | `-o` | Директория результатов (default: `./output`) |
| `--single-config` | `-s` | Режим одной комплектации |
| `--config` | `-k` | Название комплектации (для --single-config) |
| `--no-split` | | Не разделять карты |
| `--no-fuzzy` | | Отключить нечёткое сравнение |
| `--workers` | `-w` | Количество процессов (default: auto) |
| `--split-stats` | `-S` | Детальная статистика разделения |
| `--verbose` | `-v` | Подробный лог (debug) |

### Режимы

- **По умолчанию** — обработка всех комплектаций одновременно
- `--single-config` — только одна комплектация

### Вывод

1. `report.txt` — текстовый отчёт
2. `discrepancies.xlsx` — 4 листа: сводка, по комплектациям, расхождения, fuzzy matches
3. `split_cards/` — разделённые однолистовые файлы
4. `split_cards.zip` — архив

---

## In-memory API

Сервисы поддерживают загрузку из байтового потока без сохранения на диск:

```python
from burlak_parser.bom_parser import BOMService
from burlak_parser.card_parser import CardService

# Синхронная загрузка
service = BOMService()
data = service.load_from_bytes(bom_bytes)

# Асинхронная загрузка
data = await service.load_async(bom_bytes)

# Context manager
async with CardService() as service:
    data = await service.load_async(cards_bytes)
    # cleanup() вызывается автоматически
```

## SplitStatistics

`CardSplitter.split_many_parallel()` возвращает кортеж `(files, errors)`. Каждый файл сопровождается структурой `FileSplitStats`:
- Количество листов, разделено, пропущено
- Причины пропуска (template, empty, service)
- Размер файла, формат

`SplitStatistics` агрегирует данные: топ причин пропусков, топ файлов по количеству пропусков.

## verify_integrity()

Проверяет целостность результатов сверки:
- Каждая деталь из BOM имеет запись в результатах
- Суммы количества сходятся
- Нет дублирующихся записей

Возвращает `IntegrityCheck` со статусом passed/failed.

## Особенности

- Шаблонные листы с данными обрабатываются (не пропускаются)
- Служебные файлы не разделяются
- Оригинальный формат номера детали (с дефисами) сохраняется в отчёте

---

## Архитектура

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
