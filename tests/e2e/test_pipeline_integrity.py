import pytest
import aiosqlite
from pathlib import Path

from app.db import async_repository
from app.services.result_service import ResultService


@pytest.mark.asyncio
async def test_pipeline_happy_path(temp_db_path: str, mock_storage_path: Path):
    """Сквозной тест: создание задачи → отметка файлов → done → проверка результатов."""
    async with aiosqlite.connect(temp_db_path) as db:
        job_id = await async_repository.create_job(db)

        # Имитируем успешную загрузку файлов
        await async_repository.update_file_upload(db, job_id, "bom", "/tmp/bom.xlsx", True)
        await async_repository.update_file_upload(db, job_id, "archive", "/tmp/archive.zip", True)
        await async_repository.update_job_status(db, job_id, "done")

        # Создаём фейковые результаты
        job_dir: Path = mock_storage_path / str(job_id)
        job_dir.mkdir(parents=True, exist_ok=True)
        (job_dir / "diff.xlsx").write_bytes(b"dummy diff content")
        (job_dir / "translated_cards.zip").write_bytes(b"dummy cards content")

        # Проверяем сервис результатов
        await ResultService.validate_job_ready(db, job_id)
        diff_path = ResultService.get_result_path(job_id, "diff")
        assert diff_path.exists()