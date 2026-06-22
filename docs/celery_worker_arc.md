# Архитектурное описание компонента: Celery Worker Container (C4 — Level 3)

## 1. Общее описание

Данный документ описывает внутреннюю структуру, зоны ответственности и логику взаимодействия компонентов внутри контейнера **Celery Worker**. Контейнер отвечает за фоновую обработку технологических карт (ETL-пайплайн), асинхронное взаимодействие с ML-сервисом перевода и финальную предикативную сверку с эталонным BOM.

**Соответствие стадиям пайплайна:**
- `unpacking` → Unpack Task
- `analyzing_mapping` → Analyze Mapping Task
- `processing_cards` → Oper Processor (Process Card Task)
- `aggregating` → Final Aggregator (Aggregate Task)
- `packaging` → Package Task (часть Final Aggregator)

---

## 2. Спецификация компонентов

### 2.1. Task Dispatcher (Диспетчер задач)

**Тип:** Точка входа Celery (Celery Tasks Entrypoint).

**Обязанности:**
- Прием сообщений из очередей брокера (Redis).
- Маршрутизация задач на основе типа:
  - `unpack` → `unpack.delay()`
  - `analyze_mapping` → `analyze_mapping.delay()`
  - `process_card` → `process_card.delay()`
  - `aggregate` → `aggregate.delay()`
- Первичное логирование и инициализация статуса задачи в SQLite.

---

### 2.2. Модуль: Unpack Task (Распаковка архива)

**Соответствие файлу:** `app/worker/tasks/unpack.py`

**Обязанности:**
- Потоковое чтение оглавления ZIP-архива через `zipfile.ZipFile.infolist()` (запрещена физическая распаковка на диск).
- Создание записей о картах в БД (таблица `cards`).
- Перевод задачи на стадию `analyzing_mapping`.

---

### 2.3. Модуль: Analyze Mapping Task (Определение структуры через AI)

**Соответствие файлу:** `app/worker/tasks/analyze_mapping.py`

**Обязанности:**
- Вызывает `excel_service.py` для конвертации BOM и нескольких примеров карт в JSON-слепок (первые 20 строк каждого листа).
- Отправляет JSON-слепок в ML-сервис для определения полей сопоставления (ключей матчинга).
- Сохраняет результат в `jobs.mapping_config`.
- Переводит задачу на стадию `processing_cards`.

**Внутренние подкомпоненты:**

#### Excel → JSON Converter (Конвертер структуры Excel)
- **Тип:** Детерминированная функция в `excel_service.py`.
- **Обязанности:**
  - Чтение `.xlsx` через `openpyxl` (data_only=True).
  - Взятие первых 20 строк каждого листа.
  - Преобразование в JSON-слепок (структура, не данные).
  - Не извлекает материалы — только готовит слепок для AI.

#### AI Structure Analyzer (Анализатор структуры)
- **Тип:** Клиент к ML-сервису в `translation_adapter.py` или отдельный метод.
- **Обязанности:**
  - Получение JSON-слепка от Excel → JSON Converter.
  - Отправка в ML-сервис с промптом на определение структуры.
  - Получение схемы (координаты колонок, строки начала данных, ключи сопоставления).
  - **Зона ответственности AI минимальна:** не извлекает данные, только описывает координаты.

---

### 2.4. Модуль: Oper Processor (Пайплайн обработки карт)

**Соответствие файлу:** `app/worker/tasks/process_card.py`

Инкапсулирует логику поштучной обработки входящих файлов. Работает параллельно на множестве воркеров.

**Внутренние подкомпоненты:**

#### Excel Reader And Extractor (Класс-комбо)
- **Тип:** Класс в `excel_service.py`.
- **Обязанности:**
  - **Чтение:** Открывает оригинальный файл `.xlsx` напрямую из ZIP-архива через `zipfile.open()` (потоковое чтение по смещению, без распаковки на диск).
  - **Парсинг по схеме:** Использует `mapping_config` (полученный от Analyze Mapping Task) для извлечения данных из нужных колонок.
  - **Экстракция:** С помощью регулярных выражений извлекает материальные позиции («наименование — количество»).
  - **Оптимизация:** Выполняет дедупликацию всего полезного текста, формируя **сжатый пул уникальных строк** для перевода.

#### Translation Client
- **Тип:** Класс в `translation_adapter.py`.
- **Обязанности:**
  - Принимает пул уникальных строк.
  - Осуществляет асинхронные (`asyncio` / `httpx`) пакетные запросы к внешнему ML-сервису.
  - **Важно:** Работает в неблокирующем режиме (`await`). Воркер не простаивает в ожидании ответа сети, а продолжает параллельно обрабатывать другие карты (пул `gevent` или `threads`).

#### Result Writer
- **Тип:** Класс в `excel_service.py`.
- **Обязанности:**
  - Принимает каркас структуры файла (от Excel Reader And Extractor) и словарь с готовыми переводами (от Translation Client).
  - **Style-preserving режим:** Открывает оригинал через `openpyxl` и заменяет значения `.value` в ячейках с сохранением всех стилей, формул и разметки.
  - Записывает в **Shared Storage** два файла:
    - `card_XX_materials.json` (извлечённые материалы).
    - `card_XX_translated.xlsx` (переведённый документ).

#### Атомарная синхронизация (вместо Celery Chord)
- **Тип:** Логика в конце `process_card.py`.
- **Обязанности:**
  - Вызывает `repository.increment_progress(job_id)` — атомарное увеличение счетчика в SQLite (в рамках транзакции WAL).
  - Проверяет: если `processed + failed == total`, воркер **самостоятельно** отправляет задачу `aggregate.delay()`.
  - **Отказ от Celery Chord:** Тяжелые аккорды не используются.

---

### 2.5. Модуль: Final Aggregator (Финальная сверка и сборка)

**Соответствие файлам:** `app/worker/tasks/aggregate.py` и `app/worker/tasks/package.py`

Запускается строго один раз, когда все задачи `process_card` завершены (триггер — атомарная проверка счетчика в SQLite).

**Внутренние подкомпоненты:**

#### BOM Loader
- **Тип:** Класс в `excel_service.py` или `comparison_service.py`.
- **Обязанности:**
  - Импортирует мастер-файл `BOM.xlsx` в Pandas DataFrame.
  - Производит нормализацию и стандартизацию названий материалов (используя `mapping_config`).

#### Material Summarizer
- **Тип:** Класс в `comparison_service.py`.
- **Обязанности:**
  - Сканирует директорию проекта.
  - Агрегирует и суммирует объемы из всех `card_*_materials.json`.
  - При обнаружении маркеров ошибок (`card_*_error.json`) логирует их для вывода в финальный отчет.
  - **Частичный отчет:** Агрегатор успешно собирает отчет даже при наличии ошибок в некоторых картах.

#### Comparator Engine
- **Тип:** Класс в `comparison_service.py`.
- **Обязанности:**
  - Выполняет математическое сопоставление (join/merge) агрегированных данных с эталоном BOM.
  - Вычисляет дефицит, профицит и пересортицу.

#### Report Packager
- **Тип:** Класс в `comparison_service.py` + логика в `package.py`.
- **Обязанности:**
  - Генерирует финальный документ `diff.xlsx` (включая список битых файлов со статусом «Ошибка обработки файла, данные не учтены»).
  - Архивирует все файлы `card_*_translated.xlsx` в результирующий `translated_cards.zip`.
  - Производит финальную запись статуса в SQLite (`status: done` или `status: error`, если `failed > 0`).

---

## 3. Политика обработки ошибок и отказоустойчивости (Fault Tolerance)

1. **Изоляция сбоев при обработке карт:**
   - При возникновении критической ошибки в `process_card`, задача перехватывается блоком `try-except`.
   - Задача завершается со статусом `SUCCESS` (чтобы не блокировать триггер запуска Агрегатора).
   - На диск пишется маркер неисправности `card_XX_error.json`.
   - Счетчик `failed` увеличивается атомарно.

2. **Частичный отчет:**
   - Final Aggregator спроектирован для успешной сборки отчета даже при наличии ошибок в некоторых картах.
   - Необработанные файлы вносятся в итоговый `diff.xlsx` со статусом «Ошибка обработки файла, данные не учтены».

3. **Критическая ошибка:**
   - Если `failed > 0` по итогам обработки всех карт, статус задачи переводится в `error`.

---
## 4. Диаграмма компонентов (Mermaid)
```mermaid
graph TD
    AI[AI]
    TS["Translation Service<br/>(External API)"]
    SS["Shared Storage<br/>(File System)"]
    DB[("SQLite<br/>(Task DB)")]

    subgraph CeleryWorker ["Celery Worker (Container)"]
        TD[Task Dispatcher]

        %% Новая задача: Analyze Mapping
        AM[Analyze Mapping Task]
        ASA[AI Structure Analyzer]

        subgraph OperProcessorSub ["Oper Processor"]
            OP[Oper Processor]
            ERE[Excel Reader And Extractor]
            TC[Translation Client]
            RW[Result Writer]

            OP --> ERE
            ERE --> TC
            TC --> RW
        end

        subgraph FinalAggregatorSub ["Final Aggregator"]
            FA[Final Aggregator]
            BL[BOM Loader]
            MS[Material Summarizer]
            CE[Comparator Engine]
            RP[Report Packager]

            FA --> BL
            FA --> MS
            BL --> CE
            MS --> CE
            CE --> RP
        end

        TD -- "unpack.done → analyze_mapping" --> AM
        AM --> ASA
        TD -- "process_card" --> OP
        TD -- "aggregation.run" --> FA
    end

    %% Связи Analyze Mapping
    ASA <--> AI
    ASA -. "Reads BOM + sample cards" .-> SS
    AM -->|saves mapping_config| DB

    %% Связи Oper Processor
    ERE -. "Reads card_XX.xlsx using mapping_config" .-> SS
    TC -. "REST" .-> TS
    RW -. "Writes materials JSON & translated xlsx" .-> SS

    %% Связи Final Aggregator
    BL -. "Reads BOM.xlsx" .-> SS
    MS -. "Reads all card_*_materials.json" .-> SS
    RP -. "Writes discrepancy report and zip" .-> SS
    RP -->|Updates task status| DB