"""Модуль кэширования результатов обработки.

Инкрементальная обработка: при повторном запуске обрабатываются
только изменённые файлы. Кэш на базе file hash (xxhash или md5).

Использование:
    cache = ProcessingCache(".burlak_cache")
    if not cache.needs_processing(file_path):
        return cache.get_cached_result(file_path)
    result = do_processing(file_path)
    cache.store_result(file_path, result)
"""

from __future__ import annotations

import hashlib
import json
import logging
import os
import time
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Set, Tuple

logger = logging.getLogger(__name__)


def _file_hash(file_path: str, chunk_size: int = 8192) -> str:
    """Вычислить hash файла (MD5 для скорости, не для безопасности).

    Args:
        file_path: Путь к файлу.
        chunk_size: Размер чанка для чтения.

    Returns:
        Hex-строка хеша.
    """
    h = hashlib.md5()
    try:
        with open(file_path, 'rb') as f:
            while True:
                chunk = f.read(chunk_size)
                if not chunk:
                    break
                h.update(chunk)
        return h.hexdigest()
    except OSError:
        return ""


@dataclass
class CacheEntry:
    """Одна запись в кэше."""
    file_path: str
    file_hash: str
    file_size: int
    mtime: float
    processed_at: float
    output_files: List[str] = field(default_factory=list)
    metadata: Dict[str, Any] = field(default_factory=dict)


class ProcessingCache:
    """Кэш результатов обработки файлов.

    Структура:
        cache_dir/
        ├── index.json          # Основной индекс: file_path -> CacheEntry
        └── results/            # Кэшированные результаты (опционально)

    Алгоритм проверки:
        1. Быстрая проверка: mtime + file_size (миллисекунды)
        2. Точная проверка: file_hash (если mtime изменился)
    """

    def __init__(self, cache_dir: str):
        self.cache_dir = cache_dir
        self.index_path = os.path.join(cache_dir, "index.json")
        self._index: Dict[str, CacheEntry] = {}
        self._load_index()

    def _load_index(self) -> None:
        """Загрузить индекс из файла."""
        if not os.path.isfile(self.index_path):
            return
        try:
            with open(self.index_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            for path, entry_data in data.items():
                self._index[path] = CacheEntry(
                    file_path=entry_data.get("file_path", path),
                    file_hash=entry_data.get("file_hash", ""),
                    file_size=entry_data.get("file_size", 0),
                    mtime=entry_data.get("mtime", 0.0),
                    processed_at=entry_data.get("processed_at", 0.0),
                    output_files=entry_data.get("output_files", []),
                    metadata=entry_data.get("metadata", {}),
                )
            logger.debug("Загружен кэш: %d записей", len(self._index))
        except (json.JSONDecodeError, OSError) as e:
            logger.warning("Не удалось загрузить кэш: %s", e)
            self._index = {}

    def _save_index(self) -> None:
        """Сохранить индекс в файл."""
        try:
            os.makedirs(self.cache_dir, exist_ok=True)
            data = {}
            for path, entry in self._index.items():
                data[path] = {
                    "file_path": entry.file_path,
                    "file_hash": entry.file_hash,
                    "file_size": entry.file_size,
                    "mtime": entry.mtime,
                    "processed_at": entry.processed_at,
                    "output_files": entry.output_files,
                    "metadata": entry.metadata,
                }
            with open(self.index_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except OSError as e:
            logger.warning("Не удалось сохранить кэш: %s", e)

    def needs_processing(self, file_path: str) -> bool:
        """Проверить, нужно ли обрабатывать файл.

        Быстрая проверка по mtime + size, затем по hash.

        Args:
            file_path: Путь к файлу.

        Returns:
            True если файл нужно обработать (новый или изменённый).
        """
        abs_path = os.path.abspath(file_path)

        if abs_path not in self._index:
            return True

        entry = self._index[abs_path]

        # Быстрая проверка: mtime + size
        try:
            stat = os.stat(abs_path)
            if stat.st_mtime != entry.mtime or stat.st_size != entry.file_size:
                # Файл изменился — проверяем hash
                current_hash = _file_hash(abs_path)
                if current_hash != entry.file_hash:
                    return True
                # Hash совпал — обновляем mtime/size
                entry.mtime = stat.st_mtime
                entry.file_size = stat.st_size
                return False
        except OSError:
            return True

        return False

    def get_cached_result(self, file_path: str) -> Optional[CacheEntry]:
        """Получить кэшированный результат.

        Args:
            file_path: Путь к файлу.

        Returns:
            CacheEntry если есть кэш, None если нет.
        """
        abs_path = os.path.abspath(file_path)
        return self._index.get(abs_path)

    def store_result(
        self,
        file_path: str,
        output_files: Optional[List[str]] = None,
        metadata: Optional[Dict[str, Any]] = None,
    ) -> None:
        """Сохранить результат обработки в кэш.

        Args:
            file_path: Путь к исходному файлу.
            output_files: Список созданных выходных файлов.
            metadata: Дополнительные метаданные.
        """
        abs_path = os.path.abspath(file_path)
        try:
            stat = os.stat(abs_path)
            self._index[abs_path] = CacheEntry(
                file_path=abs_path,
                file_hash=_file_hash(abs_path),
                file_size=stat.st_size,
                mtime=stat.st_mtime,
                processed_at=time.time(),
                output_files=output_files or [],
                metadata=metadata or {},
            )
        except OSError as e:
            logger.warning("Не удалось сохранить в кэш %s: %s", file_path, e)

    def remove_entry(self, file_path: str) -> None:
        """Удалить запись из кэша."""
        abs_path = os.path.abspath(file_path)
        self._index.pop(abs_path, None)

    def get_stale_files(self, current_files: List[str]) -> List[str]:
        """Найти файлы в кэше, которых больше нет на диске.

        Args:
            current_files: Список текущих файлов.

        Returns:
            Список путей, которые нужно удалить из кэша.
        """
        current_set = {os.path.abspath(f) for f in current_files}
        return [p for p in self._index if p not in current_set]

    def cleanup_stale(self, current_files: List[str]) -> int:
        """Удалить из кэша записи для несуществующих файлов.

        Returns:
            Количество удалённых записей.
        """
        stale = self.get_stale_files(current_files)
        for path in stale:
            del self._index[path]
        if stale:
            logger.info("Очищено %d устаревших записей из кэша", len(stale))
        return len(stale)

    def save(self) -> None:
        """Сохранить индекс на диск."""
        self._save_index()

    @property
    def size(self) -> int:
        """Количество записей в кэше."""
        return len(self._index)

    def get_stats(self) -> Dict[str, Any]:
        """Получить статистику кэша."""
        return {
            "total_entries": len(self._index),
            "cache_dir": self.cache_dir,
            "index_exists": os.path.isfile(self.index_path),
        }
