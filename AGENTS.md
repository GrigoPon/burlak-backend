# Agent Guidelines for Backend

## Commands

- **Install deps:** `uv sync`
- **Test All:** `uv run pytest`
- **Lint/Format:** `uv run ruff check . --fix`
- **Type Check:** `uv run mypy .`

> mypy is configured with `strict = true` in `pyproject.toml`. Passing mypy with `Any` everywhere is not acceptable.

---

## Code Style & Guidelines

- **Python 3.12+**: Use strict type hinting for all function signatures.
- **Imports**: Standard Library → Third-party → Local modules.
- **Naming**: `snake_case` for functions/variables, `PascalCase` for classes/schemas.
- **Error Handling**: Use custom exceptions from `core/exceptions.py`. All API errors must follow the uniform format:
  `{"error": {"code": "CODE", "message": "text", "detail": null}}`
- **Architecture**: Follow the structure in `docs/technical_specification.md` strictly.

---

## Architecture Boundaries

This is the most critical rule. Violating it creates untestable, bloated tasks.

**`worker/tasks/` are orchestrators only.** A task is allowed to:
- Call service methods
- Update job status in the DB
- Trigger the next task via `.delay()`

**A task must never contain business logic inline.**

```python
# BAD — logic inside the task
@celery_app.task
def process_card(job_id: str, card_path: str) -> None:
    with zipfile.ZipFile(archive_path) as zf:
        with zf.open(card_path) as f:
            df = pd.read_excel(f)
            cells = [{"address": ..., "value": ...} for ...]
            response = httpx.post(ML_URL, json={"cells": cells})
            ...  # 50 more lines

# GOOD — task is an orchestrator
@celery_app.task(bind=True, max_retries=3, retry_backoff=True, retry_backoff_max=30)
def process_card(self, job_id: str, card_path: str) -> None:
    try:
        card_data = archive_service.read_card(job_id, card_path)
        parsed = excel_service.parse(card_data)
        translated = translation_adapter.translate(parsed)
        excel_service.write_translated(job_id, card_path, translated)
        result = sync_repository.increment_progress(job_id)
        if result.is_complete:
            aggregate.delay(job_id)
    except Exception as exc:
        raise self.retry(exc=exc)
```

---

## Critical Constraints

### ZIP Handling
- **Never** use `zipfile.extractall()`. The 1.3 GB archive must never be fully unpacked to disk.
- Read the archive table of contents via `zipfile.ZipFile.infolist()`.
- Stream individual files from the archive via `zipfile.open(card_path)`.

### Celery
- **No `chord` or `group`** for synchronization. They are fragile under load.
- Use atomic SQLite progress updates via `sync_repository.increment_progress()`.
- The last worker to finish calls `aggregate.delay(job_id)` when `processed + failed == total`.
- All tasks must be configured with retries:

```python
@celery_app.task(bind=True, max_retries=3, retry_backoff=True, retry_backoff_max=30)
def process_card(self, job_id: str, card_path: str) -> None: ...
```

### Database — two repositories, not one

FastAPI runs in an async context; Celery workers run in a sync context. They use different DB drivers over the same SQLite file.

| Context | Driver | Module |
|---|---|---|
| FastAPI (async) | `aiosqlite` | `db/async_repository.py` |
| Celery workers (sync) | `sqlite3` stdlib | `db/sync_repository.py` |

**Never call `aiosqlite` from a Celery task.** It forces `asyncio.run()` inside every task, creating a new event loop for each of the 1000 cards — this is an anti-pattern.

```python
# BAD — inside a Celery task
async def _work():
    async with aiosqlite.connect(DB_PATH) as db:
        await async_repository.increment_progress(db, job_id)
asyncio.run(_work())  # new event loop per task — never do this

# GOOD — inside a Celery task
sync_repository.increment_progress(job_id)  # plain sqlite3, WAL-safe
```

The atomic increment in `sync_repository.py` must use `BEGIN IMMEDIATE` to be safe under concurrent WAL writes.

### XLSX

There are two modes for output card generation. Mode is determined by job config, not hardcoded.

- **Style-preserving mode**: Open the original file via `openpyxl`. Modify **only** `cell.value`. Never touch styles, fonts, borders, formulas, or merged cells.
- **Simplified mode**: Create a new workbook with translated data. No style requirements.

### File Downloads
- Use `StreamingResponse` for all file downloads (`diff.xlsx`, `translated_cards.zip`).
- **Never** read result files into memory in the API process.

---

## Error Handling in Workers

Business requirement: this system is used in manufacturing. Partial results are more dangerous than no results.

- Any unrecoverable error in `process_card` (after all retries exhausted) must set `jobs.failed += 1`.
- After all cards are processed: if `failed > 0`, set job status to `error`, not `done`.
- A job in `error` state still generates `diff.xlsx` if aggregation ran — the error list is included in that file.
- System-level failures (Redis down, disk full, ML unreachable for all retries) set status to `error` immediately and stop the pipeline.

```python
# In sync_repository.py — atomic progress update returns pipeline decision
def increment_progress(job_id: str, *, success: bool) -> ProgressResult:
    """
    Atomically increments processed or failed counter.
    Returns ProgressResult with is_complete=True when processed + failed == total.
    Caller is responsible for triggering aggregate.delay() if is_complete.
    """
```
