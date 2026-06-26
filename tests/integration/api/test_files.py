from fastapi.testclient import TestClient


def test_chunk_upload_idempotency(
    api_client: TestClient, temp_db_path: str, mock_storage_path
):
    """Тест idempotency чанковой загрузки (повторная отправка того же чанка)."""
    # Создаём задачу
    response = api_client.post("/api/v1/jobs")
    job_id = response.json()["id"]

    chunk_data = b"test chunk data " * 100

    # Первый upload
    resp1 = api_client.put(
        f"/api/v1/jobs/{job_id}/files/bom/chunks/0",
        content=chunk_data,
        headers={"X-Total-Chunks": "2"},
    )
    assert resp1.status_code == 200

    # Повторный upload того же чанка — должен быть 200 OK
    resp2 = api_client.put(
        f"/api/v1/jobs/{job_id}/files/bom/chunks/0",
        content=chunk_data,
        headers={"X-Total-Chunks": "2"},
    )
    assert resp2.status_code == 200

    # Загрузка чанка другого размера — ошибка
    resp3 = api_client.put(
        f"/api/v1/jobs/{job_id}/files/bom/chunks/0",
        content=chunk_data + b"extra",
        headers={"X-Total-Chunks": "2"},
    )
    assert resp3.status_code == 422
    assert resp3.json()["error"]["code"] == "CHUNK_CORRUPTED"
