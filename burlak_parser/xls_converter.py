"""Модуль конвертации .xls → .xlsx через LibreOffice.

Использует LibreOffice headless mode для конвертации legacy .xls файлов
в современный .xlsx формат с сохранением изображений, форматирования
и структуры данных.

Преимущества перед xlrd:
  - Сохраняет изображения (встроенные OLE-объекты, рисунки)
  - Сохраняет conditional formatting
  - Сохраняет merged cells корректно
  - Поддерживает .xls (BIFF) форматы старых версий Excel

Примечание:
  - Требуется установленный LibreOffice в системе
  - Конвертация выполняется через subprocess (headless mode)
  - Временные файлы создаются в указанной директории
"""

from __future__ import annotations

import logging
import os
import shutil
import subprocess
import tempfile
from typing import Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

# Кэш путей к LibreOffice binary
_libreoffice_path: Optional[str] = None


def find_libreoffice() -> Optional[str]:
    """Найти путь к LibreOffice binary.

    Проверяет:
      1. Переменную окружения LIBREOFFICE_PATH
      2. which libreoffice / which soffice
      3. Типичные пути установки

    Returns:
        Путь к LibreOffice или None если не найден.
    """
    global _libreoffice_path
    if _libreoffice_path is not None:
        return _libreoffice_path

    # 1. Переменная окружения
    env_path = os.environ.get("LIBREOFFICE_PATH")
    if env_path and os.path.isfile(env_path):
        _libreoffice_path = env_path
        return _libreoffice_path

    # 2. which libreoffice / soffice
    for cmd in ("libreoffice", "soffice"):
        try:
            result = subprocess.run(
                ["which", cmd],
                capture_output=True, text=True, timeout=10,
            )
            if result.returncode == 0 and result.stdout.strip():
                path = result.stdout.strip()
                if os.path.isfile(path):
                    _libreoffice_path = path
                    return _libreoffice_path
        except (subprocess.TimeoutExpired, FileNotFoundError):
            continue

    # 3. Типичные пути
    typical_paths = [
        "/usr/bin/libreoffice",
        "/usr/bin/soffice",
        "/usr/local/bin/libreoffice",
        "/usr/local/bin/soffice",
        "/snap/bin/libreoffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
        "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
    ]
    for path in typical_paths:
        if os.path.isfile(path):
            _libreoffice_path = path
            return _libreoffice_path

    return None


def is_libreoffice_available() -> bool:
    """Проверить, доступен ли LibreOffice."""
    return find_libreoffice() is not None


def convert_xls_to_xlsx(
    xls_path: str,
    output_dir: Optional[str] = None,
    timeout: int = 120,
) -> Optional[str]:
    """Конвертировать .xls файл в .xlsx через LibreOffice.

    Args:
        xls_path: Путь к исходному .xls файлу.
        output_dir: Директория для сохранения результата.
                    Если None — создаётся временная директория.
        timeout: Таймаут конвертации в секундах.

    Returns:
        Путь к сконвертированному .xlsx файлу или None при ошибке.
    """
    lo_path = find_libreoffice()
    if lo_path is None:
        logger.warning(
            "LibreOffice не найден. Конвертация .xls → .xlsx невозможна. "
            "Установите LibreOffice или установите переменную окружения "
            "LIBREOFFICE_PATH."
        )
        return None

    if not os.path.isfile(xls_path):
        logger.error("Файл не найден: %s", xls_path)
        return None

    ext = os.path.splitext(xls_path)[1].lower()
    if ext != ".xls":
        logger.debug("Файл не .xls, конвертация не требуется: %s", xls_path)
        return None

    # Создаём временную директорию для вывода, если не указана
    temp_dir_created = False
    if output_dir is None:
        output_dir = tempfile.mkdtemp(prefix="burlak_xls_convert_")
        temp_dir_created = True
    else:
        os.makedirs(output_dir, exist_ok=True)

    try:
        logger.info(
            "Конвертация .xls → .xlsx: %s → %s",
            os.path.basename(xls_path), output_dir,
        )

        # LibreOffice headless конвертация
        cmd = [
            lo_path,
            "--headless",
            "--convert-to", "xlsx",
            "--outdir", output_dir,
            xls_path,
        ]

        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout,
        )

        if result.returncode != 0:
            logger.warning(
                "LibreOffice конвертация завершилась с кодом %d: %s",
                result.returncode, result.stderr[:500] if result.stderr else "",
            )
            # Проверяем, создался ли файл несмотря на код возврата
            # (LibreOffice有时 возвращает 1 при成功的 конвертации)

        # Ищем сконвертированный файл
        base_name = os.path.splitext(os.path.basename(xls_path))[0]
        xlsx_path = os.path.join(output_dir, f"{base_name}.xlsx")

        if os.path.isfile(xlsx_path):
            file_size = os.path.getsize(xlsx_path)
            logger.info(
                "Конвертация успешна: %s (%.1f MB)",
                os.path.basename(xlsx_path), file_size / (1024 * 1024),
            )
            return xlsx_path

        # Fallback: ищем любой .xlsx файл в output_dir с похожим именем
        for fn in os.listdir(output_dir):
            if fn.endswith(".xlsx") and base_name[:10] in fn:
                found_path = os.path.join(output_dir, fn)
                logger.info("Найден сконвертированный файл: %s", fn)
                return found_path

        logger.warning(
            "Сконвертированный .xlsx файл не найден в %s. "
            "LibreOffice вывод: %s",
            output_dir, result.stdout[:500] if result.stdout else "",
        )
        return None

    except subprocess.TimeoutExpired:
        logger.error(
            "Таймаут конвертации .xls → .xlsx (%d сек): %s",
            timeout, os.path.basename(xls_path),
        )
        return None
    except FileNotFoundError:
        logger.error("LibreOffice не найден: %s", lo_path)
        return None
    except Exception as e:
        logger.error("Ошибка конвертации .xls → .xlsx: %s", e)
        return None


def convert_xls_files_batch(
    xls_files: List[str],
    temp_dir: str,
    max_workers: int = 2,
) -> Dict[str, str]:
    """Конвертировать список .xls файлов в .xlsx.

    Args:
        xls_files: Список путей к .xls файлам.
        temp_dir: Директория для сконвертированных файлов.
        max_workers: Максимальное количество параллельных конвертаций.

    Returns:
        Словарь {оригинальный_путь: путь_к_сконвертированному_xlsx}.
        Файлы, которые не удалось сконвертировать, отсутствуют в словаре.
    """
    if not xls_files:
        return {}

    if not is_libreoffice_available():
        logger.warning(
            "LibreOffice недоступен. %d .xls файлов не будут конвертированы.",
            len(xls_files),
        )
        return {}

    os.makedirs(temp_dir, exist_ok=True)
    converted: Dict[str, str] = {}

    logger.info(
        "Конвертация %d .xls файлов через LibreOffice...",
        len(xls_files),
    )

    for xls_path in xls_files:
        xlsx_path = convert_xls_to_xlsx(xls_path, output_dir=temp_dir)
        if xlsx_path:
            converted[xls_path] = xlsx_path
            logger.info(
                "  ✓ %s → %s",
                os.path.basename(xls_path),
                os.path.basename(xlsx_path),
            )
        else:
            logger.warning(
                "  ✗ Конвертация не удалась: %s",
                os.path.basename(xls_path),
            )

    logger.info(
        "Конвертация завершена: %d/%d успешно",
        len(converted), len(xls_files),
    )
    return converted
