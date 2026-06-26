# Чек-листы и Критерии Приемки (Definition of Done) по Воркстримам

Этот документ содержит конкретные требования к качеству и готовности кода для каждого разработчика по 5 выделенным воркстримам.

---

## 1. Разработчик 1: Инфраструктура Celery и DB Persistence (Workstream A)

### Чек-лист задач:
- [ ] Создан файл [app/worker/celery_app.py](file:///burlak-backend/app/worker/celery_app.py), инициализирующий приложение Celery с настройками `settings.redis_url` в качестве брокера сообщений и бэкенда результатов.
- [ ] Настроено автоматическое обнаружение задач в пакете `app.worker.tasks`.
- [ ] В [app/db/sync_repository.py](file:///burlak-backend/app/db/sync_repository.py) добавлены и покрыты типами следующие синхронные методы работы с БД SQLite:
  - `update_mapping_config(job_id: int, mapping_config: dict) -> None` — сохранение конфигурации маппинга от ML-сервиса в виде JSON.
  - `update_job_status(job_id: int, status: str, stage: str | None = None) -> None` — обновление статуса и шага выполнения задачи.
  - `get_mapping_config(job_id: int) -> dict` — получение JSON-конфигурации маппинга для задачи.
  - `get_job_files(job_id: int) -> tuple[str, str]` — возвращает пути к файлу BOM и архиву карт.
- [ ] Все транзакции в `sync_repository.py` используют `BEGIN IMMEDIATE` для предотвращения взаимоблокировок (deadlocks) в WAL-режиме SQLite.

### Критерии приемки (Definition of Done):
1. **Тесты:** Написаны юнит-тесты в `app/db/test_db.py`, которые запускаются в синхронном режиме и проверяют корректность записи/чтения данных во всех новых методах `sync_repository.py` на тестовой SQLite базе данных.
2. **Метрики:** При инициализации `celery_app` отсутствуют ошибки импорта модулей или подключения.

---

## 2. Разработчик 2: Парсинг Excel и ML-Адаптеры (Workstream B)

### Чек-лист задач:
- [ ] **Интеграция:** Перенесены модули нормализации ([normalizer.py](file:///burlak-backend/burlak_parser/normalizer.py)) и нечеткого сравнения ([fuzzy_matcher.py](file:///burlak-backend/burlak_parser/fuzzy_matcher.py)) из папки `burlak_parser` в `app/services/` с адаптацией локальных импортов.
- [ ] Создан [app/services/snapshot_service.py](file:///burlak-backend/app/services/snapshot_service.py):
  - Метод `extract_snapshot_from_bytes(data: bytes, filename: str, max_rows: int = 50) -> dict` считывает XLSX через `openpyxl.load_workbook(io.BytesIO(data), data_only=True)`.
  - **Безопасность:** количество колонок принудительно ограничивается: `max_col = min(ws.max_column or 0, 50)`.
  - Метод `group_by_format(file_paths: list[str]) -> dict[str, list[str]]` группирует файлы по их префиксным именам.
- [ ] Создан [app/services/structure_adapter.py](file:///burlak-backend/app/services/structure_adapter.py):
  - Реализован синхронный класс `StructureAdapter` на базе `httpx.Client`.
  - Метод `analyze_structure` отправляет POST-запрос на `/api/v1/analyze-structure` и возвращает `mapping_config`.
  - Метод `translate_batch` отправляет уникальные строки на `/api/v1/translate` (пакетный перевод с китайского на английский) и возвращает словарь соответствий.
- [ ] Создан [app/services/card_parser_service.py](file:///burlak-backend/app/services/card_parser_service.py):
  - Логика классификации файлов по регулярным выражениям и ключевым словам.
  - Логика извлечения позиций деталей (номер, название, количество) с учетом границ таблицы (`end_markers`).

### Критерии приемки (Definition of Done):
1. **Отсутствие дисковых операций:** Парсер и snapshot-сервис работают полностью в памяти с `io.BytesIO` и не вызывают методы записи на диск (`tempfile.NamedTemporaryFile` не используется).
2. **Юнит-тесты:** Написаны тесты, проверяющие корректное извлечение данных из тестовых XLSX-файлов с помощью моковых конфигов разметки, а также тесты группировки файлов.

---

## 3. Разработчик 3: Задачи Основного Конвейера Celery (Workstream C)

### Чек-лист задач:
- [ ] Метод `dispatch_processing` в `app/services/job_processing_service.py` изменен для вызова Celery-задачи `unpack.delay(job_id)`.
- [ ] Создана задача [app/worker/tasks/unpack.py](file:///burlak-backend/app/worker/tasks/unpack.py):
  - Потоковое чтение ZIP через `zipfile.ZipFile.infolist()` (без полной распаковки на диск!).
  - Запись путей файлов карт в таблицу `cards` со статусом `pending`.
  - Запуск `analyze_mapping.delay(job_id)`.
- [ ] Создана задача [app/worker/tasks/analyze_mapping.py](file:///burlak-backend/app/worker/tasks/analyze_mapping.py):
  - Чтение представительных карт, формирование и отправка слепков в ML.
  - Сохранение `mapping_config` через `sync_repository`.
  - Запуск задач `process_card.delay(job_id, card_path)` для каждой карты.
- [ ] Создана задача [app/worker/tasks/process_card.py](file:///burlak-backend/app/worker/tasks/process_card.py):
  - Чтение карты через `zipfile.open()`, классификация, парсинг данных.
  - Пакетный перевод уникальных текстов через `StructureAdapter.translate_batch`.
  - Запись переведенного файла XLSX во временную директорию `/data/{job_id}/translated_cards/`.
  - Запись JSON с деталями в `/data/{job_id}/card_XX_materials.json`.
  - Атомарный инкремент прогресса в БД. Если задача последняя — вызов `aggregate.delay(job_id)`.
  - **Отказоустойчивость:** вся логика обернута в `try/except`. При падении вызывается `increment_progress(success=False)` и пишется лог ошибки.

### Критерии приемки (Definition of Done):
1. **Изоляция ошибок:** Ошибка разбора одной карты не валит всю очередь Celery; битый файл просто помечается как `failed` в БД, и прогресс идет дальше.
2. **Интеграционные тесты:** Написан тест в режиме `CELERY_TASK_ALWAYS_EAGER=True`, эмулирующий загрузку ZIP-архива и прохождение шагов `unpack` -> `analyze_mapping` -> `process_card`.

---

## 4. Разработчик 4: Агрегация и Финализация Результатов (Workstream D)

### Чек-лист задач:
- [ ] **Интеграция:** Создан модуль [app/services/comparison_service.py](file:///burlak-backend/app/services/comparison_service.py) на базе логики сопоставления (`comparator.py`) и генерации Excel-отчетов (`report_generator.py`) из папки `burlak_parser`.
- [ ] Создана задача [app/worker/tasks/aggregate.py](file:///burlak-backend/app/worker/tasks/aggregate.py):
  - Загрузка `BOM.xlsx` в Pandas DataFrame.
  - Сборка данных из всех сгенерированных `card_*_materials.json`.
  - Сопоставление данных через Pandas. Выявление расхождений (дефицит/профицит).
  - Запись отчета в `/data/{job_id}/diff.xlsx` (включает лист со списком упавших карт и текстами их ошибок).
  - Запуск `package.delay(job_id)`.
- [ ] Создана задача [app/worker/tasks/package.py](file:///burlak-backend/app/worker/tasks/package.py):
  - Архивирование всех файлов из `/data/{job_id}/translated_cards/` в `/data/{job_id}/translated_cards.zip`.
  - **Очистка диска:** Рекурсивное удаление папки `/data/{job_id}/translated_cards/` после создания архива.
  - Финальное обновление статуса задачи: `done` (если `failed == 0`) или `error` (если `failed > 0`).

### Критерии приемки (Definition of Done):
1. **Очистка диска:** На диске после завершения задачи гарантированно удаляется папка со всеми 1000 несжатыми картами XLSX, остаются только `diff.xlsx` и `translated_cards.zip`.
2. **Юнит-тесты:** Написаны тесты, проверяющие правильность формирования математических расхождений в Pandas и сборку итогового zip-файла.

---

## 5. Разработчик 5: DevOps и Окружение (Workstream E)

### Чек-лист задач:
- [ ] Создан файл `docker-compose.yml` в корне проекта.
- [ ] Описаны сервисы:
  - `redis`: образ `redis:7-alpine`.
  - `web`: сборка текущего FastAPI приложения, запуск uvicorn.
  - `worker`: сборка текущего приложения Celery, запуск воркера.
  - `ml-mock`: легковесный контейнер на Python (FastAPI/Flask), эмулирующий ответы ML-сервиса (методы `/analyze-structure` и `/translate`) для локальной разработки.
- [ ] Настроены volume-монтирования для папки `/data` и файла базы данных `jobs.db`, чтобы API и Celery-воркер имели доступ к одним и тем же файлам.

### Критерии приемки (Definition of Done):
1. **Работоспособность окружения:** Команда `docker compose up --build` собирает и запускает все сервисы без ошибок.
2. **Health Check:** Воркер Celery успешно подключается к контейнеру Redis, а FastAPI эндпоинт `/api/v1/health` отвечает статусом `200 OK`.
