"""Shared pytest configuration — автоматическая очистка временных файлов.

При завершении тестовой сессии удаляет все .xlsx/.xls файлы
и временные директории, созданные тестовыми хелперами в системе /tmp/.
"""

from __future__ import annotations

import glob
import logging
import os
import shutil
import tempfile
from typing import List


_KNOWN_PREFIXES: List[str] = [
    # test_bom_parser.py
    "bom_test_",
    "bom_empty_",
    # test_card_parser.py
    "card_test_",
    "card_multi_",
    "card_empty_",
    "card_notable2_",
    "card_notable3_",
    "card_maxrow0_",
    "card_findxls_",
    "card_xlstest_",
    "card_xlsfloat_",
    "card_xlsno_",
    "card_xlscorrupt_",
    "xlrempty_",
    "xlrint_",
    "xlrreg_",
    "xlrex_",
    "card_dir_",
    "card_nest_",
    "walk_temp_",
    "walk_clean_",
    "walk_keep_",
    "walk_nest_",
    "walk_dup_",
    "walk_corrupt_",
    "find_autoextract_",
    "find_zip_",
    "find_mix_",
    "parse_par_val_",
    "parse_svc_",
    "parse_svc_err_",
    "parse_seq_",
    "parse_par_",
    # test_splitter.py
    "splitter_test_",
    # test_report_generator.py
    "report_test_",
]


def _cleanup_temp_files() -> int:
    """Удалить все временные файлы/директории с известными префиксами из /tmp/.

    Returns:
        Количество удалённых записей.
    """
    temp_dir = tempfile.gettempdir()
    removed = 0

    for prefix in _KNOWN_PREFIXES:
        for entry in glob.glob(os.path.join(temp_dir, f"{prefix}*")):
            try:
                if os.path.isfile(entry) or os.path.islink(entry):
                    os.remove(entry)
                elif os.path.isdir(entry):
                    shutil.rmtree(entry, ignore_errors=True)
                removed += 1
            except Exception:
                pass

    return removed


def pytest_sessionfinish(session) -> None:  # type: ignore[no-untyped-def]
    """Pytest hook: очистка temp-файлов после завершения тестовой сессии."""
    removed = _cleanup_temp_files()
    if removed > 0:
        logger = logging.getLogger("burlak_parser.tests")
        logger.debug("Удалено временных файлов/директорий: %d", removed)
