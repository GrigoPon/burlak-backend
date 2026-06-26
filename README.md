# Burlak Backend

BOM parsing and comparison system for automotive manufacturing.# burlak-backend
=======
# BOM Verification System Backend (Burlak Backend)

This is the backend repository for the BOM Verification System. The system matches Bill of Materials (BOM) spreadsheets against assembly operational cards for automotive manufacturing, translates Chinese material names and card text into Russian, and generates discrepancy reports.

## Tech Stack

* Python 3.12+
* FastAPI (Uvicorn)
* Celery (Redis Broker and Backend)
* SQLite in WAL mode (aiosqlite for FastAPI, standard sqlite3 for Celery tasks)
* openpyxl & pandas (Excel parsing and comparison engine)
* httpx (ML service HTTP integration)

## Architecture Overview

The system processes large datasets asynchronously using a Celery pipeline:
1. **Unpacking:** The ZIP archive containing operational cards is parsed. Card filenames are registered in SQLite.
2. **Structure Analysis:** Representative snapshots of BOM sheets and operational cards are sent to the ML service to determine column coordinates and matching keys.
3. **Card Processing:** Workers parse cards in parallel directly from the ZIP stream, translate text, and write intermediate JSON parts list and translated sheets to disk.
4. **Aggregation:** Card data is merged with the master BOM sheet using pandas, calculating deficits and surplus materials.
5. **Packaging:** Results are zipped into a final download package, and temporary files are cleaned up.

### Database Division
* **FastAPI (Async context):** Uses `aiosqlite` via `app/db/async_repository.py`.
* **Celery workers (Sync context):** Uses standard library `sqlite3` via `app/db/sync_repository.py` to prevent event loop overhead. Transactions are written in `BEGIN IMMEDIATE` mode to be WAL-safe.

---

## Directory Structure

```
backend/
├── app/
│   ├── api/                           # HTTP routing and validation layers
│   │   └── v1/
│   ├── core/                          # System configurations, exceptions, storage management
│   ├── db/                            # Database models, connection handling, and repositories
│   ├── schemas/                       # Pydantic schemas for request/response serialization
│   ├── services/                      # Core business logic (parsers, adapters, matching)
│   ├── worker/                        # Celery application initialization and tasks pipeline
│   │   └── tasks/
│   └── main.py                        # FastAPI entrypoint
├── tests/                             # Test suite (unit, integration, e2e)
├── Dockerfile                         # Application container description
├── docker-compose.yml                 # Local development multi-container orchestration
├── pyproject.toml
└── uv.lock
```

---

## Setup and Installation

### Prerequisites
* Python 3.12+
* [uv](https://github.com/astral-sh/uv) (Python package manager)
* Redis server (running locally or in Docker)

### Installation
1. Install dependencies:
   ```bash
   uv sync
   ```

2. Copy the example environment file and configure parameters:
   ```bash
   cp .env.example .env
   ```

3. Configure environment variables in `.env`:
   * `DB_URL`: SQLite connection URL (e.g., `sqlite:///./dev.db`)
   * `REDIS_URL`: Redis server URL (e.g., `redis://localhost:6379/0`)
   * `ML_SERVICE_URL`: Endpoint of the ML service (e.g., `http://localhost:8000`)
   * `STORAGE_PATH`: Directory path for saving upload and result files

---

## Running the Application

### 1. Start the API Server
Run the FastAPI development server:
```bash
uv run uvicorn app.main:app --reload
```
The API documentation will be available at `http://127.0.0.1:8000/docs`.

### 2. Start the Celery Worker
Run the background worker instance:
```bash
uv run celery -A app.worker.celery_app worker --loglevel=info
```

---

## Development Commands

### Running Tests
Execute the entire pytest suite:
```bash
uv run pytest
```

### Linting and Formatting
Check, autofix code quality issues, and format the code style using Ruff:
```bash
uv run ruff check . --fix
uv run ruff format .
```

### Type Checking
Run strict mypy static type checking:
```bash
uv run mypy .
```
All files must fully comply with mypy strict type hinting guidelines.

---

## Key Constraints for Developers

* **No zipfile.extractall():** Never unpack the 1.3 GB zip archive to disk. Stream individual card bytes directly from the archive using `zipfile.open()`.
* **SQLite WAL Writes:** All database writes inside Celery tasks must use the synchronous repository and the `BEGIN IMMEDIATE` transaction block to avoid locking/concurrency errors.
* **Response Streaming:** Download endpoints for results must stream binary data using `StreamingResponse` or `FileResponse` to avoid loading massive archives into memory.

