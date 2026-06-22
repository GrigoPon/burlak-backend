# ТЗ на разработку Бэкенда (BOM Verification System)

> **Связанные документы:**
> - [`fastapi_api_architecture.md`](fastapi_api_architecture.md) — архитектура FastAPI бэкенда (слои, компоненты, диаграммы)
> - [`api_reference.md`](api_reference.md) — полная спецификация эндпоинтов с примерами запросов/ответов
> - [`error_handling.md`](error_handling.md) — коды ошибок, иерархия исключений, примеры
> - [`data_flow.md`](data_flow.md) — сквозной поток данных через все компоненты системы
> - [`container_architecture.md`](container_architecture.md) — C4 Level 2 (контейнеры)
> - [`celery_worker_arc.md`](celery_worker_arc.md) — C4 Level 3 (воркеры)

---

## 1. Архитектурный стек и конфигурация

* **Базовый язык:** Python 3.12
* **Фреймворк:** FastAPI + Uvicorn
* **Очередь задач:** Celery (Брокер = Redis, Результаты = Redis)
* **База данных:** SQLite в режиме WAL (библиотека `aiosqlite` для асинхронного взаимодействия)
* **Анализ таблиц:** `openpyxl` + `pandas`
* **Интеграция с ML:** Асинхронный HTTP-клиент `httpx`

Все параметры приложения настраиваются через класс `BaseSettings` (`pydantic-settings`) в файле `app/core/config.py`, считывающий переменные из `.env`:

* `REDIS_URL` (по умолчанию `redis://redis:6379/0`)
* `ML_SERVICE_URL` (адрес сервиса машинного перевода)
* `STORAGE_PATH` (корневой каталог для сохранения файлов, по умолчанию `/data`)
* `CHUNK_SIZE_BYTES` (размер одного чанка, строго `20971520` байт / 20 МБ)

---

## 2. Жизненный цикл задачи (State Machine)

Управление конвейером осуществляется через изменение полей `status` и `stage` в таблице `jobs`. 

**Важное правило:** Любая критическая ошибка на любом этапе или наличие хотя бы одного поврежденного файла (`failed > 0`) переводит статус задачи в `error`.

```
[awaiting_upload] 
       │
   (Фронтенд загрузил BOM и Архив, вызвал /start)
       ▼
[processing / unpacking] ──> [processing / analyzing_mapping] ──> [processing / processing_cards]
                                                                             │
                                      ┌──────────────────────────────────────┴────────────────────────────────┐
                                      │                                                                       │
                          (Успешно обработано 100% карт)                                        (Есть хотя бы одна ошибка / битый файл)
                                      ▼                                                                           ▼
[processing / aggregating] ──> [processing / packaging] ──> [done]                                              [error]
```

### Описание стадий:
1. **`unpacking`**: Чтение оглавления ZIP-архива, создание записей о картах в БД.
2. **`analyzing_mapping`**: Отправка BOM и нескольких примеров карт в ML-сервис (в формате JSON) для определения полей, по которым будет происходить сопоставление. Результат сохраняется в `jobs.mapping_config`.
3. **`processing_cards`**: Параллельная обработка каждой карты: парсинг → перевод через ML → инкремент прогресса в БД.
4. **`aggregating`**: Сбор всех результатов, сверка с BOM по найденным ключам, генерация `diff.xlsx` (включая список битых файлов).
5. **`packaging`**: Сборка финального архива `translated_cards.zip`.

---

## 3. Критически важные правила реализации (Памятка разработчику)

1. **Потоковая работа с ZIP (`services/archive_service.py`):** Запрещено физически распаковывать архив весом до 1.3 ГБ на жесткий диск сервера. Таска `unpack.py` должна считать оглавление архива через `zipfile.ZipFile.infolist()`. Воркеры `process_card.py` должны читать бинарный поток конкретного файла напрямую из архива по смещению через `zipfile.open()`.
2. **Идемпотентность загрузки (`api/v1/files.py`):** Если из-за сбоя сети фронтенд повторно отправляет чанк `n`, бэкенд проверяет его наличие на диске через `core/storage.py`. Если размер файла совпадает — возвращаем `200 OK`. Если не совпадает — отдаем `422 CHUNK_CORRUPTED`.
3. **Атомарная синхронизация в SQLite вместо Celery Chord:** Отказаться от использования тяжелых Celery-аккордов (`chord`). Метод `repository.increment_progress(job_id)` должен атомарно (в рамках транзакции WAL) увеличивать счетчик обработанных карт. Воркер в конце выполнения таски `process_card` проверяет: если `processed + failed == total`, он самостоятельно отправляет задачу агрегации результатов `aggregate.delay()`.
4. **Стриминг ответов (`api/v1/results.py`):** Файлы `diff.xlsx` и `translated_cards.zip` должны отдаваться клиенту через `FastAPI.responses.StreamingResponse`. Запрещено читать файлы целиком в оперативную память бэкенд-процесса.
5. **Работа с ML и XLSX:**
    * Для отправки данных в ML-сервис использовать конвертацию `xlsx` → `json`.
    * **Варианты генерации итоговых карт:**
        * *Simplified*: Создание новых файлов с переведенными данными (без сохранения стиля).
        * *Style-preserving*: Открытие оригинала через `openpyxl` и замена значений `.value` в ячейках с сохранением всех стилей, формул и разметки.

---

## 4. Контракты API (Спецификация эндпоинтов)

Все ответы в случае ошибок должны возвращаться в едином формате:

```json
{ "error": { "code": "КОД_ОШИБКИ", "message": "Человекочитаемый текст", "detail": null } }
```

### 1. POST `/api/v1/jobs` — Создать задачу
* **Выход (201 Created):** Инициализирует задачу в состоянии `awaiting_upload`.

### 2. PUT `/api/v1/jobs/{job_id}/files/{role}/chunks/{n}` — Загрузить чанк
* **Path:** `role` (`bom` или `archive`), `n` (индекс чанка, 0-based).
* **Headers:** `Content-Type: application/octet-stream`, `X-Total-Chunks`.

### 3. POST `/api/v1/jobs/{job_id}/files/{role}/complete` — Подтвердить загрузку
* **Логика:** Вызывает `storage.assemble_chunks()`, склеивает файл, удаляет чанки.

### 4. POST `/api/v1/jobs/{job_id}/start` — Запустить обработку
* **Логика:** Проверяет готовность файлов. Переводит статус в `processing`, стадию в `unpacking`, триггерит `unpack.delay()`.

### 5. GET `/api/v1/jobs/{job_id}` — Получить статус задачи
* **Выход (200 OK):** Возвращает состояние прогресса.

### 6. GET `/api/v1/jobs/{job_id}/results/diff` — Скачать таблицу расхождений
* **Выход (200 OK):** Бинарный поток `diff.xlsx`.
* **Ошибки:** `RESULTS_NOT_READY` (409 — если статус задачи не `done` и не `error`).

### 7. GET `/api/v1/jobs/{job_id}/results/cards` — Скачать переведённые карты
* **Выход (200 OK):** Бинарный поток `translated_cards.zip`.

### 8. GET `/api/v1/health` — Проверка состояния системы

---

## 5. Точная файловая структура бэкенда

```
backend/
├── app/
│   ├── api/                           # HTTP слой: роутинг и валидация
│   │   └── v1/
│   │       ├── router.py
│   │       ├── jobs.py
│   │       ├── files.py
│   │       ├── results.py
│   │       └── health.py
│   │
│   ├── worker/                        # Оркестрация фонового конвейера Celery
│   │   ├── celery_app.py
│   │   └── tasks/
│   │       ├── unpack.py              # Чтение ZIP, создание записей
│   │       ├── analyze_mapping.py     # Определение ключей сопоставления через ML
│   │       ├── process_card.py        # Парсинг, перевод, инкремент SQLite
│   │       ├── aggregate.py           # Сборка diff.xlsx
│   │       └── package.py             # Сборка translated_cards.zip
│   │
│   ├── services/                      # Бизнес-логика
│   │   ├── archive_service.py         # Потоковая работа с zipfile
│   │   ├── excel_service.py           # XLSX <-> JSON, запись данных
│   │   ├── comparison_service.py      # Алгоритмы сопоставления
│   │   └── translation_adapter.py     # HTTP-клиент для ML
│   │
│   ├── db/                            # Слой персистентности (SQLite)
│   │   ├── database.py
│   │   ├── models.py                  # Таблица jobs (вкл. mapping_config)
│   │   └── repository.py
│   │
│   ├── schemas/                       # Валидация Pydantic
│   │   ├── job.py
│   │   └── file.py
│   │
│   ├── core/                          # Системное ядро
│   │   ├── config.py
│   │   ├── exceptions.py
│   │   └── storage.py
│   │
│   └── main.py
│
├── tests/                             # Тесты
├── pyproject.toml
├── uv.lock
├── .python-version
├── Dockerfile
└── .env.example
```
