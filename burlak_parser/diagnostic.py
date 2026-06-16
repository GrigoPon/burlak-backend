"""Модуль диагностики и отладки парсинга BOM и операционных карт.

Выводит промежуточные сырые данные в JSON-файлы для визуальной проверки:
  - BOM_parsed_dump.json — что парсер нашёл в BOM
  - OC_parsed_dump.json — какие карты и детали извлечены
  - schema_detection.json — как парсер определил структуру файлов

Позволяет человеку глазами просмотреть выжимку и убедиться,
что парсер ничего не упустил и не придумал лишнего.

Режим работы:
  python -m burlak_parser.main --bom BOM.xlsx --cards ./cards/ --diagnostic
"""

from __future__ import annotations

import json
import logging
import os
from dataclasses import asdict, dataclass, field
from datetime import datetime
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)


@dataclass
class SheetSchema:
    """Схема одного листа Excel (результат эвристического анализа)."""
    sheet_name: str
    header_rows: List[int]
    column_types: Dict[str, int]
    config_columns: List[int]
    data_start_row: int
    total_rows: int
    total_columns: int
    is_bom_candidate: bool
    is_service_sheet: bool
    graphic_number_column: int = 0
    notes: List[str] = field(default_factory=list)


@dataclass
class FileSchema:
    """Схема одного Excel-файла (результат эвристического анализа)."""
    file_path: str
    file_name: str
    file_type: str  # "bom", "operational_card", "unknown"
    sheets: List[SheetSchema]
    detected_card_number: str = ""
    processing_notes: List[str] = field(default_factory=list)


@dataclass
class BOMDumpEntry:
    """Одна деталь из BOM-дампа."""
    part_number: str
    part_number_original: str
    name_cn: str
    name_en: str
    applicable_configs: List[str]
    quantities: Dict[str, float]


@dataclass
class CardDumpEntry:
    """Одна деталь из операционной карты (дамп)."""
    part_number: str
    part_number_original: str
    quantity: float
    source_card: str
    source_sheet: str
    operation_name: str = ""
    graphic_number: str = ""


@dataclass
class CardFileDump:
    """Дамп одного файла операционной карты."""
    file_path: str
    file_name: str
    card_number: str
    sheets: List[Dict[str, Any]]
    parts: List[CardDumpEntry]
    aggregated_parts: Dict[str, float]
    is_service_file: bool = False


class DiagnosticDumper:
    """Сервис генерации диагностических дампов.

    Используется в режиме --diagnostic для вывода промежуточных данных
    парсера в JSON-файлы, доступные для визуальной проверки.
    """

    def __init__(self, output_dir: str):
        """Инициализировать дампер.

        Args:
            output_dir: Директория для сохранения дампов.
        """
        self.output_dir = output_dir
        os.makedirs(output_dir, exist_ok=True)
        self._schema_data: List[FileSchema] = []
        self._bom_data: List[BOMDumpEntry] = []
        self._card_data: List[CardFileDump] = []

    def dump_file_schema(
        self,
        file_path: str,
        file_type: str,
        sheets_info: List[Dict[str, Any]],
        card_number: str = "",
        notes: Optional[List[str]] = None,
    ) -> None:
        """Сохранить схему файла.

        Args:
            file_path: Путь к файлу.
            file_type: Тип файла ("bom", "operational_card", "unknown").
            sheets_info: Информация о листах.
            card_number: Номер карты (если определён).
            notes: Примечания по обработке.
        """
        schema = FileSchema(
            file_path=file_path,
            file_name=os.path.basename(file_path),
            file_type=file_type,
            sheets=[
                SheetSchema(
                    sheet_name=s.get("name", ""),
                    header_rows=s.get("header_rows", []),
                    column_types=s.get("column_types", {}),
                    config_columns=s.get("config_columns", []),
                    data_start_row=s.get("data_start_row", 0),
                    total_rows=s.get("total_rows", 0),
                    total_columns=s.get("total_columns", 0),
                    is_bom_candidate=s.get("is_bom_candidate", False),
                    is_service_sheet=s.get("is_service_sheet", False),
                    graphic_number_column=s.get("graphic_number_column", 0),
                    notes=s.get("notes", []),
                )
                for s in sheets_info
            ],
            detected_card_number=card_number,
            processing_notes=notes or [],
        )
        self._schema_data.append(schema)
        logger.debug("Schema recorded for: %s", os.path.basename(file_path))

    def dump_bom_entry(
        self,
        part_number: str,
        part_number_original: str,
        name_cn: str = "",
        name_en: str = "",
        applicable_configs: Optional[List[str]] = None,
        quantities: Optional[Dict[str, float]] = None,
    ) -> None:
        """Добавить запись в BOM-дамп.

        Args:
            part_number: Очищенный парт-номер.
            part_number_original: Оригинальный парт-номер.
            name_cn: Название (китайский).
            name_en: Название (английский).
            applicable_configs: Список комплектаций.
            quantities: Количества по комплектациям.
        """
        entry = BOMDumpEntry(
            part_number=part_number,
            part_number_original=part_number_original,
            name_cn=name_cn,
            name_en=name_en,
            applicable_configs=applicable_configs or [],
            quantities=quantities or {},
        )
        self._bom_data.append(entry)

    def dump_card_file(
        self,
        file_path: str,
        card_number: str,
        sheets_info: List[Dict[str, Any]],
        parts: List[Dict[str, Any]],
        aggregated: Dict[str, float],
        is_service: bool = False,
    ) -> None:
        """Сохранить дамп файла операционной карты.

        Args:
            file_path: Путь к файлу.
            card_number: Номер карты.
            sheets_info: Информация о листах.
            parts: Список извлечённых деталей.
            aggregated: Агрегированные количества.
            is_service: Флаг служебного файла.
        """
        card_dump = CardFileDump(
            file_path=file_path,
            file_name=os.path.basename(file_path),
            card_number=card_number,
            sheets=sheets_info,
            parts=[
                CardDumpEntry(
                    part_number=p.get("part_number", ""),
                    part_number_original=p.get("part_number_original", ""),
                    quantity=p.get("quantity", 0.0),
                    source_card=p.get("source_card", ""),
                    source_sheet=p.get("source_sheet", ""),
                    operation_name=p.get("operation_name", ""),
                    graphic_number=p.get("graphic_number", ""),
                )
                for p in parts
            ],
            aggregated_parts=aggregated,
            is_service_file=is_service,
        )
        self._card_data.append(card_dump)
        logger.debug("Card dump recorded for: %s", os.path.basename(file_path))

    def save_all(self) -> Dict[str, str]:
        """Сохранить все дампы в JSON-файлы.

        Returns:
            Словарь {тип_дампа: путь_к_файлу}.
        """
        outputs: Dict[str, str] = {}
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        # 1. Schema dump
        schema_path = os.path.join(self.output_dir, "schema_detection.json")
        self._save_json(schema_path, {
            "generated_at": timestamp,
            "total_files": len(self._schema_data),
            "files": [self._to_dict(s) for s in self._schema_data],
        })
        outputs["schema"] = schema_path
        logger.info("Schema dump saved: %s", schema_path)

        # 2. BOM dump
        bom_path = os.path.join(self.output_dir, "BOM_parsed_dump.json")
        self._save_json(bom_path, {
            "generated_at": timestamp,
            "total_parts": len(self._bom_data),
            "parts": [self._to_dict(e) for e in self._bom_data],
        })
        outputs["bom"] = bom_path
        logger.info("BOM dump saved: %s (%d parts)", bom_path, len(self._bom_data))

        # 3. OC (operational cards) dump
        oc_path = os.path.join(self.output_dir, "OC_parsed_dump.json")
        self._save_json(oc_path, {
            "generated_at": timestamp,
            "total_files": len(self._card_data),
            "files": [self._to_dict(f) for f in self._card_data],
        })
        outputs["oc"] = oc_path
        logger.info("OC dump saved: %s (%d files)", oc_path, len(self._card_data))

        return outputs

    @staticmethod
    def _save_json(path: str, data: Any) -> None:
        """Сохранить данные в JSON-файл с красивым форматированием."""
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2, default=str)

    @staticmethod
    def _to_dict(obj: Any) -> Any:
        """Конвертировать dataclass в dict (рекурсивно)."""
        if hasattr(obj, "__dataclass_fields__"):
            return asdict(obj)
        if isinstance(obj, dict):
            return {k: DiagnosticDumper._to_dict(v) for k, v in obj.items()}
        if isinstance(obj, (list, tuple)):
            return [DiagnosticDumper._to_dict(item) for item in obj]
        return obj


def create_diagnostic_from_bom(
    bom_data: Any,
    output_dir: str,
) -> str:
    """Создать диагностический дамп из BOM-данных.

    Args:
        bom_data: BOMData из bom_parser.
        output_dir: Директория для дампов.

    Returns:
        Путь к созданному дампу.
    """
    dumper = DiagnosticDumper(output_dir)

    # Записываем каждую деталь BOM
    for pn, part in bom_data.parts.items():
        quantities = {}
        for config_name in bom_data.config_names:
            qty = bom_data.config_quantities.get(config_name, {}).get(pn, 0.0)
            if qty > 0:
                quantities[config_name] = qty

        dumper.dump_bom_entry(
            part_number=pn,
            part_number_original=part.part_number,
            name_cn=part.name_cn,
            name_en=part.name_en,
            applicable_configs=part.applicable_configs,
            quantities=quantities,
        )

    # Сохраняем с метаданными конфигураций
    outputs = dumper.save_all()
    bom_path = outputs.get("bom", "")

    # Добавляем config_names и config_quantities в дамп
    if bom_path and os.path.exists(bom_path):
        with open(bom_path, "r", encoding="utf-8") as f:
            bom_json = json.load(f)
        bom_json["config_names"] = bom_data.config_names
        bom_json["config_quantities"] = {
            cn: dict(qtys) for cn, qtys in bom_data.config_quantities.items()
        }
        with open(bom_path, "w", encoding="utf-8") as f:
            json.dump(bom_json, f, ensure_ascii=False, indent=2, default=str)

    return bom_path


def create_diagnostic_from_cards(
    cards_data: Any,
    output_dir: str,
) -> str:
    """Создать диагностический дамп из данных операционных карт.

    Args:
        cards_data: CardsData из card_parser.
        output_dir: Директория для дампов.

    Returns:
        Путь к созданному дампу.
    """
    dumper = DiagnosticDumper(output_dir)

    for result in cards_data.card_results:
        parts_list = []
        for cp in result.parts:
            parts_list.append({
                "part_number": cp.part_number,
                "part_number_original": cp.part_number,
                "quantity": cp.quantity,
                "source_card": cp.source_card,
                "source_sheet": cp.source_sheet,
                "operation_name": "",
                "graphic_number": "",
            })

        sheets_info = []
        for si in result.sheets:
            sheets_info.append({
                "name": si.sheet_name,
                "card_number": si.card_number,
                "operation_name": si.operation_name,
                "is_valid": si.is_valid,
                "has_data": si.has_data,
            })

        dumper.dump_card_file(
            file_path=result.file_path,
            card_number=result.card_number,
            sheets_info=sheets_info,
            parts=parts_list,
            aggregated=result.aggregated_parts,
            is_service=result.is_service_file,
        )

    outputs = dumper.save_all()
    return outputs.get("oc", "")
