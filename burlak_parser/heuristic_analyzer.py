"""
Эвристический анализатор Excel-документов — ядро универсального парсинга
BOM-листов и операционных карт автомобильного производства.

Не зависит от конкретных брендов, форматов, индексов колонок и префиксов.
Определяет структуру документа динамически через анализ:
  - Заголовков колонок (по словарю синонимов на 3 языках)
  - Содержимого ячеек (типы данных, паттерны part-number)
  - Расположения таблиц и границ данных
  - Имён файлов и названий листов

Поддерживает: китайский, английский, русский языки.

Алгоритмы:
  1. find_header_rows — поиск строки/строк заголовков (оценка по ключевым словам)
  2. detect_column_types — классификация колонок (part_no, name, qty, config, meta)
  3. find_data_region — определение границ таблицы данных
  4. extract_card_number — извлечение номера операционной карты из любых источников
  5. build_global_name_dict — сбор всех парт-номеров и названий по всему документу
"""

from __future__ import annotations

import logging
import os
import re
from typing import Any, Dict, List, Optional, Set, Tuple

from burlak_parser.normalizer import (
    clean_part_number,
    is_valid_part_number,
)

logger = logging.getLogger(__name__)


# ═══════════════════════════════════════════════════════════════════════
# СЛОВАРИ СИНОНИМОВ для определения колонок
# ═══════════════════════════════════════════════════════════════════════    # --- Парт-номер / Код детали ---
PART_NO_KEYWORDS: List[str] = [
    # Китайский
    "零件号", "零部件件号", "零件编码", "物料编码", "件号", "物料号",
    "料号", "零部件代号", "代号", "代码", "编码", "物料代码",
    "零件号(中文）", "零件号(中文)", "零件号（中文）",
    # Английский
    "part no", "partno", "part number", "part_no", "part#",
    "part number(中文）", "part number(中文)",
    "item code", "material code", "material no", "material number",
    "component code", "component number", "code",
    # Русский
    "код детали", "номер детали", "деталь", "код",
    "артикул", "каталожный номер",
]

# Ключевые слова, которые НЕ должны быть в колонке part_no
# (мета-колонки, содержащие «код» или «номер», но не являющиеся part-номером)
PART_NO_ANTI_KEYWORDS: List[str] = [
    "cpac", "fnd", "gpc", "поставщик", "supplier",
    "vehicle", "материал", "описание",
    "серийный", "serial",  # серийный номер, serial number — не part_no
]

# --- Название детали ---
NAME_KEYWORDS: List[str] = [
    # Китайский
    "零件名称", "零部件名称", "物料名称", "物料描述", "描述",
    "名称", "物料名称/描述", "材料名称",
    "零件名称(中文）", "零件名称(中文)", "零件名称（中文）",
    "零件名称(英文）", "零件名称(英文)", "零件名称（英文）",
    "物料描述（中文）", "物料描述(中文）",
    "物料描述（英文）", "物料描述(英文）",
    # Английский
    "part name", "description", "material description",
    "item description", "component name",
    "part name(cn)", "part name(en)", "part name(cn）", "part name(en）",
    "name(cn)", "name(en)", "name（cn）", "name（en）",
    # Русский
    "наименование", "наименование детали",
    "название", "описание", "наименование детали",
    "описание детали",
]

# --- Ключевые слова, которые НЕ должны быть в колонке name_cn ---
# (колонки с factory/supplier/manufacturer — это НЕ название детали)
NAME_ANTI_KEYWORDS: List[str] = [
    "工厂", "厂家", "供应商", "制造商", "生产商", "模块",
    "factory", "manufacturer", "supplier", "module",
    "завод", "производитель", "поставщик",
]

# --- Количество ---
QTY_KEYWORDS: List[str] = [
    # Китайский
    "用量", "数量", "单车用量", "每车用量", "数量/用量", "标配数量",
    "用量/数量", "单位用量",
    # Английский
    "qty", "quantity", "usage", "qty per",
    # Русский
    "количество", "кол-во", "расход", "норма",
]

# Ключевые слова, которые НЕ должны быть в колонке qty
# (колонки, содержащие "количество" НО другого типа)
QTY_ANTI_KEYWORDS: List[str] = [
    # Китайский
    "工具数量",   # tool quantity — не количество деталей
    "扭矩数量",   # torque quantity
    "工具",       # tool
    "模具数量",   # mould/die quantity
    "工装数量",   # fixture quantity
]

# --- Стандартные служебные колонки (не комплектации) ---
META_KEYWORDS: List[str] = [
    # Китайский
    "序号", "行号", "修订", "版本", "层级", "等级",
    "标识", "发运", "采购", "度量单位", "uom",
    "gpc", "fnd", "物料状态", "来源车间", "使用工厂",
    "目标车间", "供应商", "供应商代码", "供应商名称",
    "生产工厂", "供货工厂", "制造工厂", "装配工厂", "安装工厂",
    "mwo", "mwo单号", "生效日期", "失效日期",
    "整车物料号", "变更单号", "eop", "eos",
    "零件成熟度", "物料组", "物料组描述",
    "cpac编码", "cpac描述", "品牌", "车系",
    "卸货工厂",
    "备注", "说明", "附注", "注",
    "分类", "类别", "车型",
    "状态号", "模块状态", "供货状态", "PBOM供货",
    "平台属性", "设计层次", "装配层次",
    "工位范围", "工位", "工序",
    "零件质量", "IA编码", "货源", "货源描述", "结构货源",
    "单车用量", "组件数量", "发动机附件",
    # Торсионные/моментные колонки — НЕ комплектации
    "扭矩", "力矩", "动态扭矩", "残余扭矩", "扭矩角度",
    "扭矩关重", "扭矩说明", "扭矩监控",
    "图示编号", "图纸编号", "图纸号", "示意图编号",
    # Английский
    "serial no", "serial no.", "serial", "seq", "sequence",
    "revision", "rev", "version", "ver", "level",
    "ship", "purchase", "uom", "unit", "gpc code",
    "fnd code", "make/buy", "source shop", "using plant",
    "target shop", "supplier", "supplier code", "supplier name",
    "mwo", "effective date", "expire date",
    "vehicle material", "vehicle model",
    "remark", "note", "notes", "comment",
    "category", "classification", "type",
    "logo", "identification", "id",
    "torque", "nm", "n·m", "moment",
    # Русский
    "примечание", "комментарий", "завод",
    "поставщик", "дата", "статус",
    "система", "узел", "подразделение",
    "расход на один автомобиль", "количество компонентов",
]

# Только колонки с ФИКСИРОВАННЫМ текстом (не данные, не конфиги)
STRICT_META_KEYWORDS: List[str] = [
    "序号", "修订", "版本", "度量单位", "uom", "gpc", "fnd",
    "零件成熟度", "make/buy", "cpac编码",
    "serial no", "serial", "revision", "level",
    "变更记录", "文件编号",
]

# Колонка с номером схемы/операции (для привязки деталей к операционным картам)
GRAPHIC_NUMBER_KEYWORDS: List[str] = [
    # Китайский
    "图示编号", "图号", "图纸编号", "图纸号", "示意图编号",
    "工序号", "工位号", "工位编号",
    # Английский
    "graphic number", "drawing number", "drawing no", "drawing no.",
    "operation number", "operation no", "operation no.",
    "station number", "station no",
    # Русский
    "номер схемы", "номер операции", "номер чертежа",
    "код операции", "позиция схемы",
]

# Слова для определения "служебный лист" (не BOM)
SERVICE_SHEET_KEYWORDS: List[str] = [
    "封面", "目录", "记录表", "空表", "范本", "填写说明",
    "содержание", "обложка",
    # Дополнительные мета-листы (не BOM)
    "变更记录", "变更",   # change log
    "汇总",               # summary
    "原稿",               # draft
    "分装",               # sub-assembly (Changan: 分装明细)
    "分总成",             # sub-assembly list (Changan: 分总成明细)
    "申请",               # application/request
    "路线",               # routing
    "ebom",               # engineering BOM (another view, not the main)
    "mbom",               # manufacturing BOM (another view, not the main)
    "bom原稿",            # BOM draft
    "bom汇总",            # BOM summary
    "bom变更",            # BOM change
    "物料号汇总",         # material summary
]


# ═══════════════════════════════════════════════════════════════════════
# ПАТТЕРНЫ
# ═══════════════════════════════════════════════════════════════════════

# Паттерн для извлечения номера операционной карты из текста/имени файла:
#   - SQRT, SQR, SWM, G01, T1L, JETOUR и любые другие буквенные префиксы
#   - Комбинации с цифрами, дефисами, может содержать подпаттерны типа -A-AS-
#   - Чисто цифровые коды (минимум 2 цифры) или буква + 2+ цифры
CARD_NUMBER_RE = re.compile(
    r"(?:"
    r"  [A-Za-z]+\d*[A-Za-z]*(?:-\d+)+[\w-]*"         # SQRT1L-17-AS-04001, SQRT1JL-17-AS-01
    r"|"
    r"  [A-Za-z]{1,3}\d{3,}[\w-]*"                     # A001, G01, T1L
    r"|"
    r"  (?:TP|процесс|операция|card)\s*[-]?\s*\d+"     # TP-123, операция 5
    r")",
    re.IGNORECASE | re.VERBOSE,
)

# Паттерн для номеров операций в имени файла (любой бренд):
#   - Цифры (минимум 2) в начале имени или после префикса
OP_NUMBER_IN_FILENAME_RE = re.compile(
    r"(?:^|[-_\s])(\d{2,})(?:[-_\s]|$)"
)

# Паттерн для "буква + 2+ цифры" в начале имени
LETTERS_DIGITS_RE = re.compile(r"^([A-Za-z]{1,3}\d{2,})")

# Паттерн для префикс-AS-паттерна (например G01-A-AS- или SQRT1L-A-AS-)
PREFIX_AS_RE = re.compile(
    r"^([A-Za-z0-9]+[-_])?[A-Za-z]?[-_]?AS[-_]?\d+",
    re.IGNORECASE,
)

# Китайские иероглифы
CJK_RE = re.compile(r"[\u4e00-\u9fff\u3400-\u4dbf]")

# Русские буквы
CYRILLIC_RE = re.compile(r"[\u0400-\u04ff]")

# Артефакты кодировки Excel XML
XML_HEX_RE = re.compile(r"_x[0-9a-fA-F]{4}_")
XML_ARTIFACTS_RE = re.compile(r"(_x[0-9a-fA-F]{4}_|\r\n|[\r\n])", re.IGNORECASE)


# ═══════════════════════════════════════════════════════════════════════
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ═══════════════════════════════════════════════════════════════════════

def normalize_text(text: str) -> str:
    """Привести текст к нижнему регистру, удалить лишние пробелы и артефакты кодировки."""
    s = str(text).lower()
    # Удаляем артефакты кодировки Excel XML (carriage return)
    s = XML_HEX_RE.sub("", s)
    s = s.replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
    return " ".join(s.split())


def clean_cell_text(text: Any) -> str:
    """Очистить текст ячейки от артефактов кодировки Excel XML.

    Удаляет _x000d_, _x000A_, \r, \n (в любом регистре) и берёт только ПЕРВУЮ часть
    (когда в ячейке два значения: китайское + английское через \n).
    """
    if text is None:
        return ""
    s = str(text).strip()
    match = XML_ARTIFACTS_RE.search(s)
    if match:
        s = s[:match.start()]
    return s.strip()


def looks_like_part_number(value: Any) -> float:
    """Проверить значение ячейки на принадлежность к парт-номеру.

    Возвращает float (0.0–1.0) — уверенность.
    """
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return 0.1  # Числа редко бывают part-номерами (но могут быть)
    s = str(value).strip()
    if not s or len(s) < 3:
        return 0.0
    # Оценка
    score = 0.0
    # Содержит буквы + цифры
    has_alpha = bool(re.search(r"[A-Za-z]", s))
    has_digit = bool(re.search(r"\d", s))
    has_cjk = bool(CJK_RE.search(s))
    has_cyrillic = bool(CYRILLIC_RE.search(s))
    # Если есть кириллица или иероглифы — маловероятно, что это part-no
    if has_cjk or has_cyrillic:
        score -= 0.3
    if has_alpha and has_digit:
        score += 0.6
    elif has_digit and len(s) >= 6:
        score += 0.3
    # Наличие дефисов — признак part-no
    if "-" in s:
        score += 0.2
    # Служебные слова
    if any(kw in s.lower() for kw in ["零件", "物料", "部件", "part", "компонент"]):
        return 0.1  # Это заголовок, не номер
    return min(max(score, 0.0), 1.0)


def looks_like_name(value: Any) -> float:
    """Проверить значение ячейки на принадлежность к названию детали.

    Возвращает float (0.0–1.0) — уверенность.
    """
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return 0.0  # Числа не бывают названиями
    s = str(value).strip()
    if not s or len(s) < 2:
        return 0.0
    if is_valid_part_number(s):
        return 0.1  # Похоже на part-no, а не название
    score = 0.0
    has_cjk = bool(CJK_RE.search(s))
    has_cyrillic = bool(CYRILLIC_RE.search(s))
    has_alpha = bool(re.search(r"[A-Za-z]", s))
    has_digit = bool(re.search(r"\d", s))
    # Наличие иероглифов или кириллицы — сильный признак названия
    if has_cjk:
        score += 0.5
    if has_cyrillic:
        score += 0.4
    # Только буквы (без цифр) — похоже на название
    if has_alpha and not has_digit:
        score += 0.3
    # Длинный текст — признак названия
    if len(s) > 10:
        score += 0.2
    # Если содержит и буквы и цифры — может быть и названием
    if has_alpha and has_digit:
        score -= 0.1
    return min(max(score, 0.0), 1.0)


def looks_like_quantity(value: Any) -> float:
    """Проверить значение ячейки на принадлежность к количеству.

    Возвращает float (0.0–1.0) — уверенность.
    """
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return 1.0  # Число — почти всегда количество
    s = str(value).strip()
    if not s:
        return 0.0
    # Строка, представляющая число
    try:
        float(s.replace(",", "."))
        return 0.9
    except ValueError:
        pass
    # 'S' или '-' — возможные значения VIN-разбивки
    if s.upper() == "S" or s == "-":
        return 0.2  # Не количество, а "такая же"
    return 0.0


def extract_card_number_from_filepath(file_path: str) -> str:
    """Извлечь номер операционной карты из пути к файлу.

    Универсальный алгоритм:
      1. Берём базовое имя файла без расширения
      2. Ищем известные паттерны номеров
      3. Если паттерн не найден — возвращаем базовое имя

    Примеры:
      "SQRT1L-17-AS-04001-20点扫描" -> "SQRT1L-17-AS-04001"
      "G01-AS-05001-Установка" -> "G01-AS-05001"
      "A123-Контроль" -> "A123"
      "038-Установка двери" -> "038"
      "TP-0123-Main" -> "TP-0123"
    """
    basename = os.path.basename(file_path)
    name_no_ext = os.path.splitext(basename)[0]

    # Попытка 1: Полный паттерн карты (SQRT1L-17-AS-04001 и т.д.)
    match = CARD_NUMBER_RE.match(name_no_ext)
    if match:
        return _normalize_card_number(match.group(0).strip("- "))

    # Попытка 2: Префикс-AS-паттерн
    match = PREFIX_AS_RE.match(name_no_ext)
    if match:
        prefix = match.group(0).strip("- ")
        if prefix:
            return prefix

    # Попытка 3: Буква + 2+ цифры в начале
    match = LETTERS_DIGITS_RE.match(name_no_ext)
    if match:
        return match.group(1)

    # Попытка 4: Цифры (минимум 2) в начале имени
    match = OP_NUMBER_IN_FILENAME_RE.match(name_no_ext)
    if match:
        return match.group(1)

    # Возвращаем базовое имя
    return name_no_ext


def _normalize_card_number(card_no: str) -> str:
    """Нормализовать номер карты, исправляя распространённые ошибки данных.

    Исправления:
      - SSQRT → SQRT (двойная S в начале — опечатка в исходнике)
    """
    if card_no.upper().startswith("SSQRT"):
        normalized = "SQRT" + card_no[5:]
        logger.debug("Нормализация номера карты: %s → %s", card_no, normalized)
        return normalized
    return card_no


def _is_numeric_string(s: str) -> bool:
    """Проверить, представляет ли строка число (целое или дробное)."""
    try:
        float(s.replace(",", "."))
        return True
    except (ValueError, TypeError):
        return False


# ═══════════════════════════════════════════════════════════════════════
# ОСНОВНОЙ КЛАСС
# ═══════════════════════════════════════════════════════════════════════

class HeuristicAnalyzer:
    """Эвристический анализатор Excel-листов.

    Позволяет динамически определять структуру BOM и операционных карт
    без привязки к конкретным форматам брендов.
    """

    # Максимальное количество строк для сканирования заголовков
    MAX_HEADER_SCAN_ROWS = int(os.environ.get("BURLAK_MAX_HEADER_SCAN_ROWS", 30))
    # Максимальная ширина сканирования колонок (для SWM-формата, где qty может быть в C30)
    MAX_COL_SCAN_WIDTH = int(os.environ.get("BURLAK_MAX_COL_SCAN_WIDTH", 40))
    # Минимальный порог уверенности для определения колонки
    CONFIDENCE_THRESHOLD = float(os.environ.get("BURLAK_CONFIDENCE_THRESHOLD", 0.3))

    # Бонус к header score для листов с BOM-подобными именами
    _BOM_SHEET_NAME_KEYWORDS: Tuple[str, ...] = (
        "bom", "总装", "涂装", "焊装", "零部件", "附件",
        "сборка", "комплект", "список деталей",
    )

    @staticmethod
    def find_header_rows(
        ws: Any,
        max_rows: Optional[int] = None,
        sheet_name: str = "",
    ) -> List[int]:
        """Найти строки заголовков в листе.

        Анализирует первые max_rows строк, вычисляя для каждой
        'header score' на основе ключевых слов.

        Args:
            ws: Лист Excel.
            max_rows: Максимум строк для сканирования.
            sheet_name: Имя листа (для бонуса к score у BOM-подобных листов).

        Returns:
            Список номеров строк-кандидатов (отсортирован по убыванию score).
            Пустой список, если ничего не найдено.
        """
        scores: List[Tuple[int, float]] = []
        max_col = ws.max_column or 50

        if max_rows is None:
            max_rows = HeuristicAnalyzer.MAX_HEADER_SCAN_ROWS

        # Бонус к score для листов с BOM-подобными именами
        name_bonus = 0.0
        if sheet_name:
            name_lower = sheet_name.lower()
            for kw in HeuristicAnalyzer._BOM_SHEET_NAME_KEYWORDS:
                if kw in name_lower:
                    name_bonus = 0.05
                    break

        for row_idx in range(1, min(max_rows + 1, (ws.max_row or 100) + 1)):
            row_values = [
                str(HeuristicAnalyzer.get_cell_value(ws, row_idx, c) or "")
                for c in range(1, min(max_col + 1, 30))
            ]

            if not any(v.strip() for v in row_values):
                continue

            score = HeuristicAnalyzer._score_header_row(row_values)
            if score > 0:
                scores.append((row_idx, score))

        # Сортировка по убыванию score
        scores.sort(key=lambda x: -x[1])

        # Применяем бонус к лучшему score
        if scores and name_bonus > 0:
            scores[0] = (scores[0][0], scores[0][1] + name_bonus)

        # Возвращаем только строки со score выше порога
        # Пониженный порог (0.18) для поддержки BOM-листов с нестандартными заголовками
        threshold = 0.18
        if scores:
            threshold = max(scores[0][1] * 0.4, 0.18)

        result = [r for r, s in scores if s >= threshold]

        if result:
            logger.debug(
                "Найдены строки заголовков: %s (score: %s)",
                result, [f"{s:.2f}" for _, s in scores if s >= threshold],
            )
        else:
            logger.debug("Строки заголовков не найдены")

        return result

    @staticmethod
    def _score_header_row(row_values: List[str]) -> float:
        """Оценить, насколько строка похожа на заголовок.

        Учитывает:
          - Количество совпадений с ключевыми словами (part_no, name, qty)
          - Плотность непустых ячеек
          - Разнообразие типов содержимого
        """
        if not row_values:
            return 0.0

        part_no_matches = 0
        name_matches = 0
        qty_matches = 0
        meta_matches = 0
        non_empty = 0
        total = len(row_values)

        for val in row_values:
            v = normalize_text(val)
            if not v:
                continue
            non_empty += 1

            # Проверка на ключевые слова
            if any(kw in v for kw in PART_NO_KEYWORDS):
                part_no_matches += 1
            if any(kw in v for kw in NAME_KEYWORDS):
                name_matches += 1
            if any(kw in v for kw in QTY_KEYWORDS):
                qty_matches += 1
            if any(kw in v for kw in META_KEYWORDS):
                meta_matches += 1

        # Плотность непустых ячеек
        density = non_empty / max(total, 1)

        # Общий счёт
        score = (
            part_no_matches * 2.0 +
            name_matches * 1.5 +
            qty_matches * 1.5 +
            meta_matches * 0.5 +
            density * 0.5
        )

        # Нормализация
        max_possible = total * 2.0
        return score / max_possible

    @staticmethod
    def detect_column_types(ws: Any, header_rows: List[int]) -> Dict[str, int]:
        """Определить типы колонок на основе заголовков и содержимого.

        Анализирует заголовки и проверяет типы данных в ячейках.

        Returns:
            Словарь: {
                'part_no': int (номер колонки),
                'name_cn': int,
                'name_en': int,
                'qty': int,
                'config_start': int (первая колонка комплектации),
            }
            Нулевые значения означают, что колонка не найдена.
        """
        max_col = ws.max_column or 100
        header_row = header_rows[0] if header_rows else 1

        col_types: Dict[str, int] = {}
        column_scores: Dict[str, List[Tuple[int, float]]] = {
            'part_no': [],
            'name_cn': [],
            'name_en': [],
            'qty': [],
        }

        # Сбор заголовков (из нескольких строк, если есть)
        header_texts: Dict[int, str] = {}
        for c in range(1, max_col + 1):
            texts = []
            for hr in header_rows:
                v = HeuristicAnalyzer.get_cell_value(ws, hr, c)
                if v:
                    texts.append(str(v))
            header_texts[c] = " ".join(texts).strip().lower()

        # Фаза 1: Оценка колонок по заголовкам
        for c in range(1, max_col + 1):
            text = header_texts[c]
            if not text:
                continue

            # Сначала проверяем, не является ли колонка заведомо мета-колонкой
            # (содержит ключевые слова-антипаттерны, исключающие part_no)
            is_anti_part_no = any(kw in text for kw in PART_NO_ANTI_KEYWORDS)

            # Проверка на part_no (только если не анти-паттерн)
            if not is_anti_part_no:
                best_kw = None
                best_score = 0.0
                for kw in PART_NO_KEYWORDS:
                    if kw.lower() in text:
                        specificity = min(len(kw), 6) / 6.0
                        score = 0.7 + 0.3 * specificity
                        if score > best_score:
                            best_score = score
                            best_kw = kw
                if best_kw is not None:
                    column_scores['part_no'].append((c, best_score))
                else:
                    # Fallback: fuzzy match
                    for kw in PART_NO_KEYWORDS:
                        kw_norm = kw.lower().replace(" ", "").replace("(", "").replace(")", "")
                        text_norm = text.replace(" ", "").replace("(", "").replace(")", "")
                        if kw_norm in text_norm:
                            specificity = min(len(kw), 6) / 6.0
                            score = 0.5 + 0.3 * specificity
                            if score > best_score:
                                best_score = score
                                best_kw = kw
                    if best_kw is not None:
                        column_scores['part_no'].append((c, best_score))

            # Проверка на name_cn (с китайскими иероглифами)
            is_cn_name = False
            best_name_kw = None
            for kw in NAME_KEYWORDS:
                if kw.lower() in text:
                    if best_name_kw is None or len(kw) > len(best_name_kw):
                        best_name_kw = kw
            if best_name_kw is not None:
                kw = best_name_kw
                # Проверяем на анти-паттерны (factory/supplier)
                is_anti_name = any(ak in text for ak in NAME_ANTI_KEYWORDS)
                if not is_anti_name:
                    # Проверяем на русский/английский
                    has_cjk = bool(CJK_RE.search(text))
                    has_cyrillic = bool(CYRILLIC_RE.search(text))
                    cn_hints = ["中文", "cn", "(chinese)", "chinese"]
                    is_cn = has_cjk or any(h in text for h in cn_hints)
                    is_en = has_cyrillic or "英文" in text or "en)" in text or "(en" in text

                    specificity = min(len(kw), 6) / 6.0
                    base_score = 0.7 + 0.3 * specificity

                    if is_cn and not is_en:
                        column_scores['name_cn'].append((c, base_score))
                    elif is_en or has_cyrillic:
                        column_scores['name_en'].append((c, base_score))
                    elif "英文" in text or "英文）" in text or "en)" in text or "en）" in text:
                        column_scores['name_en'].append((c, 1.0))
                    elif "中文" in text or "中文）" in text or "中文)" in text:
                        column_scores['name_cn'].append((c, 1.0))
                    else:
                        column_scores['name_cn'].append((c, base_score))
                    is_cn_name = True

            if not is_cn_name:
                # Проверка на описание (description)
                text_lower = text.lower()
                if "descript" in text_lower or "наимен" in text_lower or "описан" in text_lower:
                    column_scores['name_cn'].append((c, 0.7))

            # Проверка на qty (с анти-ключевыми словами)
            # Исключаем колонки, содержащие QTY_ANTI_KEYWORDS
            is_anti_qty = any(ak.lower() in text for ak in QTY_ANTI_KEYWORDS)
            if not is_anti_qty:
                for kw in QTY_KEYWORDS:
                    if kw.lower() in text:
                        column_scores['qty'].append((c, 1.0))
                        break

        # Фаза 2: Верификация содержимым (проверяем ячейки с данными)
        data_start = header_rows[-1] + 1 if header_rows else 2
        sample_end = min(data_start + 30, (ws.max_row or data_start + 30) + 1)

        # Для каждого типа колонки проверяем содержимое
        for col_type in ['part_no', 'name_cn', 'name_en', 'qty']:
            current_scores = column_scores[col_type]
            verified_scores: List[Tuple[int, float]] = []

            for c, score in current_scores:
                if score >= 0.9:
                    verified_scores.append((c, score))
                    continue

                # Выборка содержимого
                part_no_hits = 0
                name_hits = 0
                qty_hits = 0
                total_samples = 0

                for r in range(data_start, sample_end):
                    v = HeuristicAnalyzer.get_cell_value(ws, r, c)
                    if v is None or (isinstance(v, str) and not v.strip()):
                        continue
                    total_samples += 1

                    pn_score = looks_like_part_number(v)
                    nm_score = looks_like_name(v)
                    qt_score = looks_like_quantity(v)

                    if pn_score > 0.6:
                        part_no_hits += 1
                    if nm_score > 0.6:
                        name_hits += 1
                    if qt_score > 0.8:
                        qty_hits += 1

                if total_samples > 0:
                    pn_ratio = part_no_hits / total_samples
                    nm_ratio = name_hits / total_samples
                    qt_ratio = qty_hits / total_samples

                    # Корректировка score на основе содержимого
                    if col_type == 'part_no' and pn_ratio > 0.3:
                        score = max(score, 0.7)
                    elif col_type == 'name_cn' and nm_ratio > 0.3:
                        score = max(score, 0.6)
                    elif col_type == 'name_en' and nm_ratio > 0.2:
                        score = max(score, 0.5)
                    elif col_type == 'qty' and qt_ratio > 0.3:
                        score = max(score, 0.7)

                verified_scores.append((c, score))

            column_scores[col_type] = verified_scores

        # Выбор лучших кандидатов
        for col_type in ['part_no', 'name_cn', 'name_en', 'qty']:
            scores = column_scores[col_type]
            if scores:
                # Сортировка по score, затем выбираем первый
                scores.sort(key=lambda x: (-x[1], x[0]))
                best_col, best_score = scores[0]
                if best_score >= HeuristicAnalyzer.CONFIDENCE_THRESHOLD:
                    key_map = {
                        'part_no': 'part_no',
                        'name_cn': 'name_cn',
                        'name_en': 'name_en',
                        'qty': 'qty',
                    }
                    col_types[key_map[col_type]] = best_col

        # Фаза 3: Fallback по содержимому для part_no (с проверкой заголовков!)
        if 'part_no' not in col_types:
            col_types['part_no'] = HeuristicAnalyzer._find_part_no_by_content(
                ws, data_start, sample_end, max_col, header_texts,
            )

        # Фаза 4: Fallback для имени
        if 'name_cn' not in col_types and 'name_en' not in col_types:
            name_col = HeuristicAnalyzer._find_name_by_content(ws, data_start, sample_end, max_col)
            if name_col:
                col_types['name_cn'] = name_col

        # Фаза 5: Если name_en найден, но контент — русский/китайский → переназначаем в name_cn
        # (но только если заголовок НЕ содержит явных английских маркеров)
        name_en_col = col_types.get('name_en', 0)
        name_cn_col = col_types.get('name_cn', 0)
        if name_en_col and not name_cn_col:
            # Проверяем заголовок на явные английские маркеры
            header_text = header_texts.get(name_en_col, "")
            has_en_marker = any(
                m in header_text
                for m in ["英文", "english", "en)", "en）", "(en", "（en", "inglés"]
            )
            if not has_en_marker:
                # Проверяем контент name_en: если там кириллица или CJK → это name_cn
                has_cjk_cyrillic = False
                for r in range(data_start, sample_end):
                    v = HeuristicAnalyzer.get_cell_value(ws, r, name_en_col)
                    if v and isinstance(v, str):
                        if CJK_RE.search(v) or CYRILLIC_RE.search(v):
                            has_cjk_cyrillic = True
                            break
                if has_cjk_cyrillic:
                    col_types['name_cn'] = name_en_col
                    del col_types['name_en']

        logger.debug(
            "Определены колонки: part_no=%s, name_cn=%s, name_en=%s, qty=%s",
            col_types.get('part_no'), col_types.get('name_cn'),
            col_types.get('name_en'), col_types.get('qty'),
        )
        return col_types

    @staticmethod
    def _get_part_no_keywords() -> List[str]:
        """Вернуть список PART_NO_KEYWORDS для внешнего использования.

        Нужно для card_parser._collect_raw_rows, где нет прямого импорта PART_NO_KEYWORDS.
        """
        return PART_NO_KEYWORDS

    # Cache for merged cell maps: id(ws) -> {(row, col): (top_row, top_col)}
    _merged_cell_cache: Dict[int, Dict[Tuple[int, int], Tuple[int, int]]] = {}

    @classmethod
    def _build_merged_cell_map(cls, ws: Any) -> Dict[Tuple[int, int], Tuple[int, int]]:
        """Build a map from non-top-left merged cells to their top-left source.

        In openpyxl, only the top-left cell of a merged range has a value.
        All other cells return None. This map allows resolving None cells
        to their merged source.
        """
        cache_key = id(ws)
        if cache_key in cls._merged_cell_cache:
            return cls._merged_cell_cache[cache_key]

        merged_map: Dict[Tuple[int, int], Tuple[int, int]] = {}
        try:
            ranges = ws.merged_cells.ranges
            for mr in ranges:
                top_row = mr.min_row
                top_col = mr.min_col
                for r in range(mr.min_row, mr.max_row + 1):
                    for c in range(mr.min_col, mr.max_col + 1):
                        if r != top_row or c != top_col:
                            merged_map[(r, c)] = (top_row, top_col)
        except (AttributeError, TypeError, IndexError):
            pass

        cls._merged_cell_cache[cache_key] = merged_map
        return merged_map

    @staticmethod
    def get_cell_value(ws: Any, row: int, col: int) -> Any:
        """Получить значение ячейки через универсальный API (worksheet/excel_sheet).

        Работает как с openpyxl.Worksheet, так и с ExcelSheet (из card_parser).
        Автоматически отбрасывает NaN и Infinity значения.
        Поддерживает merged cells — если ячейка является частью merged range,
        возвращает значение из верхней левой ячейки.
        """
        import math
        try:
            if hasattr(ws, 'cell_value'):
                val = ws.cell_value(row, col)
            else:
                val = ws.cell(row=row, column=col).value

            # If value is None, check merged cell map
            if val is None and row > 0 and col > 0:
                merged_map = HeuristicAnalyzer._build_merged_cell_map(ws)
                source = merged_map.get((row, col))
                if source is not None:
                    top_row, top_col = source
                    if hasattr(ws, 'cell_value'):
                        val = ws.cell_value(top_row, top_col)
                    else:
                        val = ws.cell(row=top_row, column=top_col).value

            if isinstance(val, float) and (math.isnan(val) or math.isinf(val)):
                return None
            return val
        except Exception as e:
            logger.debug("get_cell_value error at row=%d, col=%d: %s", row, col, e)
            return None

    @staticmethod
    def is_cell_strike(ws: Any, row: int, col: int) -> bool:
        """Проверить, зачеркнут ли шрифт в ячейке.

        Всегда возвращает False для:
          - xlrd (нет данных о зачёркивании)
          - Ячеек без явного шрифта
          - Исключений
          - col <= 0 (невалидный индекс колонки)

        Безопасна для:
          - font.strike = None (не установлен) -> False
          - font.strike = True/False -> соответствующее значение
          - font.strike = "sngStrike"/"dblStrike" (строковые значения openpyxl) -> True
          - font = None -> False
        """
        if col <= 0:
            return False
        try:
            # Handle card_parser's ExcelSheet wrapper
            if hasattr(ws, "_ws") and hasattr(ws, "_engine"):
                if ws._engine != "openpyxl":
                    return False
                cell = ws._ws.cell(row=row, column=col)
                if cell is None:
                    return False
                font = cell.font
                if font is None:
                    return False
                strike_val = getattr(font, 'strike', None)
                if strike_val is None:
                    return False
                if isinstance(strike_val, str):
                    return strike_val.lower() in ("sngstrike", "dblstrike", "true")
                return bool(strike_val)

            # openpyxl Worksheet
            if hasattr(ws, 'cell'):
                cell = ws.cell(row=row, column=col)
                if cell is None:
                    return False
                font = cell.font
                if font is None:
                    return False
                strike_val = getattr(font, 'strike', None)
                if strike_val is None:
                    return False
                if isinstance(strike_val, str):
                    return strike_val.lower() in ("sngstrike", "dblstrike", "true")
                return bool(strike_val)
            return False
        except Exception as e:
            logger.debug("is_cell_strike error at row=%d, col=%d: %s", row, col, e)
            return False

    @staticmethod
    def get_strike_rows(ws: Any, rows: range, cols: list) -> set:
        """Batch-detect rows with strikethrough in any of the given columns.

        Returns a set of row numbers where at least one cell has strikethrough.
        This is much faster than calling is_cell_strike() per cell.
        """
        strike_rows: set = set()
        for row in rows:
            for col in cols:
                try:
                    if hasattr(ws, "_ws") and hasattr(ws, "_engine"):
                        if ws._engine != "openpyxl":
                            break
                        cell = ws._ws.cell(row=row, column=col)
                    elif hasattr(ws, 'cell'):
                        cell = ws.cell(row=row, column=col)
                    else:
                        break
                    if cell is None:
                        continue
                    font = cell.font
                    if font is None:
                        continue
                    strike_val = getattr(font, 'strike', None)
                    if strike_val is None:
                        continue
                    is_strike = False
                    if isinstance(strike_val, str):
                        is_strike = strike_val.lower() in ("sngstrike", "dblstrike", "true")
                    else:
                        is_strike = bool(strike_val)
                    if is_strike:
                        strike_rows.add(row)
                        break  # Already found for this row, no need to check other cols
                except Exception:
                    continue
        return strike_rows


    @staticmethod
    def _find_part_no_by_content(
        ws: Any, start_row: int, end_row: int, max_col: int,
        header_texts: Optional[Dict[int, str]] = None,
    ) -> int:
        """Fallback: найти колонку парт-номера по содержимому ячеек.

        Использует ДВА подхода:
          1. Стандартный: looks_like_part_number (буквы+цифры, дефисы).
          2. Data Profiling: если колонка без распознаваемого заголовка
             содержит >50% альфа-цифровых значений длиной 8-15 символов —
             классифицирует её как Part Number кандидат.

        Args:
            ws: Лист
            start_row, end_row: Диапазон строк для анализа
            max_col: Максимальная колонка
            header_texts: Заголовки колонок (для исключения мета-колонок)

        Returns:
            Номер колонки или 0.
        """
        col_scores: Dict[int, float] = {}

        for c in range(1, max_col + 1):
            # Исключаем колонки, заголовок которых — заведомо служебный
            if header_texts:
                text = header_texts.get(c, "")
                if text:
                    is_meta = False
                    for kw in STRICT_META_KEYWORDS:
                        if kw.lower() in text:
                            is_meta = True
                            break
                    if not is_meta:
                        for kw in META_KEYWORDS:
                            if kw.lower() in text:
                                is_meta = True
                                break
                    if is_meta:
                        continue

            # Собираем непустые значения
            values: List[str] = []
            for r in range(start_row, end_row):
                v = HeuristicAnalyzer.get_cell_value(ws, r, c)
                if v is not None:
                    s = str(v).strip()
                    if s:
                        values.append(s)

            if len(values) <= 2:
                continue

            # ── Pre-check: Skip columns with too many date-like values ──
            date_re = re.compile(
                r"^(\d{1,2}[./\-]\d{1,2}[./\-]\d{2,4}"
                r"|\d{4}[./\-]\d{1,2}[./\-]\d{1,2}"
                r"|\d{1,2}\s*[а-яА-ЯёЁ]{3,8}\s*\d{2,4}"
                r"|\d{1,2}\s+[a-zA-Z]{3,8}\s*\d{2,4})$"
            )
            date_hits = sum(1 for v in values if date_re.match(v))
            if date_hits / len(values) > 0.3:
                continue

            # ── Подход 1: Стандартный (looks_like_part_number) ──
            pn_hits = sum(1 for v in values if looks_like_part_number(v) > 0.6)
            pn_ratio = pn_hits / len(values)

            # ── Подход 2: Data Profiling (альфа-цифровые 8-15 символов) ──
            # Если заголовок не определён или не содержит мета-ключевых слов,
            # пробуем проджектировать колонку как part_no
            has_header = bool(header_texts and header_texts.get(c, ""))
            alpha_numeric_hits = 0
            for val in values:
                # Очищаем от распространённых разделителей
                cleaned = val.replace("-", "").replace(".", "").replace("_", "").replace("/", "").replace(" ", "")
                if not cleaned:
                    continue
                # Проверяем: содержит И буквы И цифры, длина 8-15
                has_alpha = bool(re.search(r"[A-Za-z]", cleaned))
                has_digit = bool(re.search(r"\d", cleaned))
                length_ok = 8 <= len(cleaned) <= 15
                if has_alpha and has_digit and length_ok:
                    # Дополнительная проверка: не слишком много разных символов (не UUID/GUID)
                    unique_chars = len(set(cleaned))
                    if unique_chars >= 4:  # Минимум 4 уникальных символа (не повторяющийся паттерн)
                        alpha_numeric_hits += 1

            an_ratio = alpha_numeric_hits / len(values) if values else 0

            # Комбинированный score
            combined_score = max(pn_ratio, an_ratio)

            # Штраф за CJK/cyrillic в значениях (это названия, не part-no)
            cjk_hits = sum(1 for v in values if bool(CJK_RE.search(str(v))))
            cjk_ratio = cjk_hits / len(values) if values else 0
            if cjk_ratio > 0.3:
                combined_score *= 0.3

            if combined_score > 0.3:
                col_scores[c] = combined_score

        if col_scores:
            best = max(col_scores, key=col_scores.get)
            if col_scores[best] > 0.3:
                logger.info(
                    "Колонка part_no найдена по содержимому + data profiling: "
                    "%d (score=%.2f, values=%d)",
                    best, col_scores[best], sum(1 for _ in range(start_row, end_row)
                                                 if HeuristicAnalyzer.get_cell_value(ws, _, best) is not None),
                )
                return best
        return 0

    @staticmethod
    def _find_name_by_content(ws: Any, start_row: int, end_row: int, max_col: int) -> int:
        """Fallback: найти колонку названия по содержимому ячеек."""
        col_scores: Dict[int, float] = {}
        for c in range(1, max_col + 1):
            hits = 0
            total = 0
            for r in range(start_row, end_row):
                v = HeuristicAnalyzer.get_cell_value(ws, r, c)
                if v is None:
                    continue
                total += 1
                if looks_like_name(v) > 0.5:
                    hits += 1
            if total > 2:
                ratio = hits / total
                # Исключаем колонки с part-number
                pn_ratio = sum(1 for r in range(start_row, end_row)
                               if looks_like_part_number(HeuristicAnalyzer.get_cell_value(ws, r, c)) > 0.6) / max(total, 1)
                if ratio > 0.3 and pn_ratio < 0.3:
                    col_scores[c] = ratio - pn_ratio * 0.5
        if col_scores:
            best = max(col_scores, key=col_scores.get)
            if col_scores[best] > 0.2:
                logger.info("Колонка name найдена по содержимому: %d (score=%.2f)", best, col_scores[best])
                return best
        return 0

    @staticmethod
    def detect_config_columns(ws: Any, header_rows: List[int], col_types: Dict[str, int]) -> List[int]:
        """Определить колонки комплектаций.

        Алгоритм:
          1. Берём все колонки, не определённые как part_no/name/qty/meta
          2. Проверяем их на наличие числовых значений (количеств)
          3. Фильтруем VIN-разбивку (колонки без чисел)
          4. Возвращаем отсортированный список колонок-комплектаций

        Args:
            ws: Лист Excel
            header_rows: Найденные строки заголовков
            col_types: Определённые типы колонок

        Returns:
            Список номеров колонок комплектаций.
        """
        max_col = ws.max_column or 200
        header_row = header_rows[0] if header_rows else 1
        known_cols = {v for v in col_types.values() if v > 0}

        # Собираем заголовки всех колонок
        headers: Dict[int, str] = {}
        for c in range(1, max_col + 1):
            v = HeuristicAnalyzer.get_cell_value(ws, header_row, c)
            if v is not None:
                headers[c] = normalize_text(str(v))

        # Колонки-кандидаты: справа от part_no, исключая известные
        part_no_col = col_types.get('part_no', 1)
        candidates: List[int] = []
        last_named_col = part_no_col  # последняя колонка с заголовком
        for c in range(part_no_col + 1, max_col + 1):
            if c in known_cols:
                continue
            header_text = headers.get(c, "")

            # Проверка на служебные/мета-колонки
            is_meta = False
            if header_text:
                for kw in STRICT_META_KEYWORDS:
                    if kw.lower() in header_text:
                        is_meta = True
                        break
                if is_meta:
                    continue

                for kw in META_KEYWORDS:
                    if kw.lower() in header_text:
                        is_meta = True
                        break
                if is_meta:
                    continue
                last_named_col = c
            else:
                # Колонка без заголовка — добавляем только если она правее
                # последней колонки с заголовком (зона конфигураций).
                if c <= last_named_col:
                    continue  # та же или левее — пропускаем как мета-колонку

            candidates.append(c)

        # Проверка на числовые значения и валидность контента
        data_start = header_rows[-1] + 1 if header_rows else 2
        sample_end = min(data_start + 50, (ws.max_row or data_start + 50) + 1)

        # Паттерны, которые НЕ являются валидными значениями комплектаций
        _torque_range_re = re.compile(r"\d+\s*[±\-]\s*\d+")
        _bolt_pattern_re = re.compile(r"^[Mm]\d")
        _text_heavy_re = re.compile(r"[^\d\s.,;:]", re.UNICODE)

        col_has_numbers: Dict[int, bool] = {}
        col_has_real_numbers: Dict[int, bool] = {}  # actual numeric values (not just S/-)
        for c in candidates:
            has_valid_config = False
            has_real_number = False
            invalid_hits = 0
            total_non_empty = 0
            for r in range(data_start, sample_end):
                v = HeuristicAnalyzer.get_cell_value(ws, r, c)
                if v is None or (isinstance(v, str) and not v.strip()):
                    continue
                total_non_empty += 1

                if isinstance(v, (int, float)) and v > 0:
                    has_valid_config = True
                    has_real_number = True
                    continue
                if isinstance(v, str):
                    stripped = v.strip()
                    if stripped in ('S', 's', 'Y', 'y', '–', '-', ''):
                        has_valid_config = True
                        continue
                    # Reject torque ranges: "40-50", "50±5", "1.6±0.1"
                    if _torque_range_re.search(stripped):
                        invalid_hits += 1
                        continue
                    # Reject bolt designations: "M6x16", "M8"
                    if _bolt_pattern_re.match(stripped):
                        invalid_hits += 1
                        continue
                    # Reject text-heavy values (not numeric config)
                    try:
                        val = float(stripped)
                        if val > 0:
                            has_valid_config = True
                            has_real_number = True
                            continue
                    except ValueError:
                        pass
                    # If mostly non-numeric text, not a config column
                    if len(stripped) > 3 and _text_heavy_re.search(stripped):
                        invalid_hits += 1

            # Отбрасываем столбцы, где все значения одинаковые (коды заводов
            # типа "1020" повторяются в каждой строке — это мета-данные, а не комплектации),
            # НО только если это значение не является валидным маркером комплектации (S/-/Y/число).
            unique_data_vals = set()
            data_rows_checked = 0
            val_counter: Dict[str, int] = {}
            for r in range(data_start, sample_end):
                v = HeuristicAnalyzer.get_cell_value(ws, r, c)
                if v is not None and str(v).strip():
                    # Пропускаем значения, которые выглядят как подзаголовки
                    # (текст на Eng/Chi без цифр — типичная строка-подзаголовок)
                    sv = str(v).strip()
                    if len(sv) > 2 and not any(ch.isdigit() for ch in sv):
                        continue
                    unique_data_vals.add(sv)
                    val_counter[sv] = val_counter.get(sv, 0) + 1
                    data_rows_checked += 1
            # Skip if >70% of values are identical (factory codes, dates, etc.)
            if data_rows_checked >= 5 and len(unique_data_vals) >= 1:
                most_common_count = max(val_counter.values()) if val_counter else 0
                most_common_val = max(val_counter, key=val_counter.get) if val_counter else ""
                dup_ratio = most_common_count / data_rows_checked
                if dup_ratio > 0.7:
                    is_valid_marker = (
                        most_common_val.upper() in ('S', 'Y') or
                        most_common_val in ('-', '\u2013', '\u2014') or
                        _is_numeric_string(most_common_val)
                    )
                    if not is_valid_marker:
                        col_has_numbers[c] = False
                        col_has_real_numbers[c] = False
                        continue

            # Reject column if too many invalid values or no valid config values
            if total_non_empty > 0 and invalid_hits / total_non_empty > 0.3:
                col_has_numbers[c] = False
                col_has_real_numbers[c] = False
            else:
                col_has_numbers[c] = has_valid_config
                col_has_real_numbers[c] = has_real_number

        # Определяем границу VIN-разбивки (используем real_numbers — колонки только с S/- это маркеры)
        first_non_numeric: Optional[int] = None
        found_numeric = False
        for c in candidates:
            if col_has_real_numbers.get(c, False):
                found_numeric = True
            elif found_numeric:
                first_non_numeric = c
                break

        config_cols: List[int] = []
        if first_non_numeric is not None:
            # Берём ОБЕ группы: numeric (числовые.qty) + non-numeric (S/- маркеры).
            # Уникальные-значения фильтр уже отсеял мета-колонки (коды заводов и т.д.)
            numeric_cols = [c for c in candidates if c < first_non_numeric and col_has_numbers.get(c, False)]
            non_numeric_cols = [c for c in candidates if c >= first_non_numeric and col_has_numbers.get(c, False)]
            config_cols = numeric_cols + non_numeric_cols
            logger.debug(
                "VIN-разбивка с колонки %d. Numeric: %d, Non-numeric: %d. Всего: %d",
                first_non_numeric, len(numeric_cols), len(non_numeric_cols), len(config_cols),
            )
        else:
            config_cols = [c for c in candidates if col_has_numbers.get(c, False)]

        # Если ничего не найдено — берём все кандидаты
        if not config_cols and candidates:
            config_cols = list(candidates)

        logger.debug("Колонки комплектаций: %s (всего %d)", config_cols, len(config_cols))
        return config_cols

    @staticmethod
    def find_graphic_number_column(
        ws: Any,
        header_rows: List[int],
        col_types: Dict[str, int],
    ) -> int:
        """Найти колонку с номером схемы/операции (图示编号 / Graphic Number).

        Используется для привязки деталей из BOM к операционным картам
        (формат Changan: последний столбец таблицы содержит номер операции).

        Алгоритм:
          1. Ищем колонку по ключевым словам GRAPHIC_NUMBER_KEYWORDS
          2. Если не нашли по заголовку — ищем по содержимому (паттерн DP-CH-A01 и т.д.)
          3. Возвращаем номер колонки или 0 если не найдена.

        Args:
            ws: Лист Excel.
            header_rows: Найденные строки заголовков.
            col_types: Определённые типы колонок.

        Returns:
            Номер колонки или 0.
        """
        max_col = ws.max_column or 50
        known_cols = {v for v in col_types.values() if v > 0}

        # Фаза 1: Поиск по заголовкам
        for c in range(1, max_col + 1):
            if c in known_cols:
                continue
            for hr in header_rows:
                v = HeuristicAnalyzer.get_cell_value(ws, hr, c)
                if v is not None:
                    text = str(v).strip().lower()
                    for kw in GRAPHIC_NUMBER_KEYWORDS:
                        if kw.lower() in text:
                            logger.debug(
                                "Колонка 图示编号 найдена по заголовку: колонка %d, '%s'",
                                c, kw,
                            )
                            return c

        # Фаза 2: Поиск по содержимому (fallback — очень мягкий паттерн)
        # Любой непустой строке, содержащей буквы И цифры (но не чисто число),
        # и при этом не похожей на количество.
        data_start = header_rows[-1] + 1 if header_rows else 2
        sample_end = min(data_start + 30, (ws.max_row or data_start) + 1)

        for c in range(1, max_col + 1):
            if c in known_cols:
                continue
            hits = 0
            total = 0
            for r in range(data_start, sample_end):
                v = HeuristicAnalyzer.get_cell_value(ws, r, c)
                if v is None:
                    continue
                total += 1
                s = str(v).strip()
                if not s or len(s) < 2:
                    continue
                # Отсеиваем чистые числа (количество) и совсем короткие строки
                try:
                    float(s.replace(",", "."))
                    continue  # Это число — не graphic number
                except ValueError:
                    pass
                # Содержит буквы И цифры (любой порядок, любые разделители)
                has_alpha = bool(re.search(r"[A-Za-z]", s))
                has_digit = bool(re.search(r"\d", s))
                if has_alpha and has_digit:
                    hits += 1
            if total > 2 and hits / total > 0.3:
                logger.debug(
                    "Колонка 图示编号 найдена по содержимому: колонка %d (hits=%.2f)",
                    c, hits / total,
                )
                return c

        return 0

    @staticmethod
    def extract_card_number_from_sheet(ws: Any, file_path: str, max_scan_rows: int = 15) -> str:
        """Извлечь номер карты из содержимого листа или имени файла.

        Алгоритм:
          1. Сканируем первые max_scan_rows строк на наличие номера карты
             (ищем паттерны SQRT... или код операции)
          2. Если найден — возвращаем
          3. Если нет — извлекаем из имени файла

        Args:
            ws: Лист Excel
            file_path: Путь к файлу (для fallback)
            max_scan_rows: Количество строк для сканирования

        Returns:
            Номер карты.
        """
        max_col = min(ws.max_column or 20, 20)

        for r in range(1, min(max_scan_rows + 1, (ws.max_row or max_scan_rows) + 1)):
            for c in range(1, max_col + 1):
                v = HeuristicAnalyzer.get_cell_value(ws, r, c)
                if v is not None:
                    # Strip _x000d_ / _x000A_ / \r\n artifacts from WPS Office
                    text = str(v)
                    text = XML_HEX_RE.sub("", text)
                    text = text.replace('\r', '').replace('\n', ' ').strip()
                    match = CARD_NUMBER_RE.search(text)
                    if match:
                        card_no = match.group(0).strip("- ")
                        card_no = _normalize_card_number(card_no)
                        logger.debug("Номер карты из содержимого листа: %s", card_no)
                        return card_no

        # Fallback: из имени файла
        card_no = extract_card_number_from_filepath(file_path)
        logger.debug("Номер карты из имени файла: %s", card_no)
        return card_no

    @staticmethod
    def _is_false_positive_part_no(val: str, keyword: str) -> bool:
        """Check if a PART_NO_KEYWORD match is a false positive.

        Handles cases like:
          - "更改文件号" contains "件号" but is NOT a part number column
          - "变更记录" contains "记录" but is NOT a part number column

        Returns True if the match should be REJECTED.
        """
        if not val:
            return False

        # Compound-word prefixes that invalidate certain short keywords
        # When these precede a keyword, the compound has a different meaning
        _COMPOUND_PREFIXES: Dict[str, List[str]] = {
            "件号": ["更改", "文件", "变更", "修订", "版本"],
        }

        prefixes = _COMPOUND_PREFIXES.get(keyword, [])
        for prefix in prefixes:
            if prefix in val and keyword in val:
                # Check that the prefix appears BEFORE the keyword
                prefix_pos = val.find(prefix)
                kw_pos = val.find(keyword)
                if prefix_pos >= 0 and kw_pos >= 0 and prefix_pos < kw_pos:
                    # The keyword is part of a compound word — reject
                    return True

        return False

    @staticmethod
    def find_part_table(
        ws: Any,
        start_row: int = 1,
    ) -> Optional[Tuple[int, int, int, int]]:
        """Найти таблицу деталей в листе.

        Анализирует строки в поиске заголовков таблицы деталей
        (part_no, qty, name).

        Поддерживает 2-строчные заголовки: если name/qty не найдены в той же
        строке, что и part_no — продолжает поиск на следующих 3-5 строках
        (характерно для SWM карт).

        Args:
            ws: Лист Excel.
            start_row: Номер строки, с которой начинать поиск (для многооперационных листов).

        Returns:
            (header_row, part_no_col, qty_col, name_col) или None.
        """
        max_row = ws.max_row or 200
        max_col = ws.max_column or 50

        scan_width = HeuristicAnalyzer.MAX_COL_SCAN_WIDTH

        for row_idx in range(start_row, max_row + 1):
            row_values: List[str] = []
            for col_idx in range(1, min(max_col + 1, scan_width)):
                v = HeuristicAnalyzer.get_cell_value(ws, row_idx, col_idx)
                row_values.append(str(v).strip().lower() if v is not None else "")

            if not any(row_values):
                continue

            # Строка-заголовок должна содержать минимум 2 непустых ячейки
            non_empty_count = sum(1 for v in row_values if v)
            if non_empty_count < 2:
                continue

            # Оценка строки как заголовка таблицы деталей
            has_part_no = any(
                not HeuristicAnalyzer._is_false_positive_part_no(v, kw)
                for v in row_values
                for kw in PART_NO_KEYWORDS
                if kw in v
            )
            if not has_part_no:
                continue

            # Определяем колонки (из сканированного диапазона колонок)
            part_no_col: Optional[int] = None
            qty_col: Optional[int] = None
            name_col: Optional[int] = None

            for kw in PART_NO_KEYWORDS:
                if part_no_col is not None:
                    break
                for col_idx, val in enumerate(row_values, 1):
                    if kw in val and len(val) < 50:
                        if HeuristicAnalyzer._is_false_positive_part_no(val, kw):
                            continue
                        part_no_col = col_idx
                        break

            if part_no_col is None:
                continue

            for col_idx, val in enumerate(row_values, 1):
                if qty_col is None and not any(ak in val for ak in QTY_ANTI_KEYWORDS):
                    if any(kw in val for kw in QTY_KEYWORDS):
                        qty_col = col_idx
                if name_col is None and any(kw in val for kw in NAME_KEYWORDS):
                    # Skip anti-keywords (factory/supplier)
                    if not any(ak in val for ak in NAME_ANTI_KEYWORDS):
                        name_col = col_idx

            # ── Multi-row header scan (BELOW) ──
            # Если qty или name не найдены в той же строке — ищем на следующих 10 строках
            # (характерно для SWM карт, где part_no на R21, а name/qty на R28 — отступ 7 строк)
            if qty_col is None or name_col is None:
                for scan_offset in range(1, min(11, max_row - row_idx + 1)):
                    scan_row = row_idx + scan_offset
                    scan_values: List[str] = []
                    for col_idx in range(1, min(max_col + 1, scan_width)):
                        v = HeuristicAnalyzer.get_cell_value(ws, scan_row, col_idx)
                        scan_values.append(str(v).strip().lower() if v is not None else "")

                    if not any(scan_values):
                        continue

                    # Проверяем на part_no в сканируемой строке — НЕ забираем её как qty/name
                    has_pn_scan = any(
                        not HeuristicAnalyzer._is_false_positive_part_no(sv, kw)
                        for sv in scan_values
                        for kw in PART_NO_KEYWORDS
                        if kw in sv
                    )
                    if has_pn_scan:
                        # Это новый заголовок — останавливаем поиск
                        break

                    if qty_col is None:
                        for col_idx, val in enumerate(scan_values, 1):
                            if any(kw in val for kw in QTY_KEYWORDS):
                                qty_col = col_idx
                                break
                    if name_col is None:
                        for col_idx, val in enumerate(scan_values, 1):
                            if any(kw in val for kw in NAME_KEYWORDS):
                                name_col = col_idx
                                break

                    if qty_col is not None and name_col is not None:
                        logger.debug(
                            "Найдены name/qty на строке %d (multi-row header below)",
                            scan_row,
                        )
                        break

            # ── Row-above header scan ──
            # Если qty/name всё ещё не найдены — ищем ВЫШЕ part_no-заголовка
            # (для SWM-формата, где qty может быть в той же строке, но правее лимита,
            #  а заголовок qty мог быть расположен в строке над part_no)
            if qty_col is None or name_col is None:
                for scan_offset in range(1, min(6, row_idx)):
                    scan_row = row_idx - scan_offset
                    scan_values: List[str] = []
                    for col_idx in range(1, min(max_col + 1, scan_width)):
                        v = HeuristicAnalyzer.get_cell_value(ws, scan_row, col_idx)
                        scan_values.append(str(v).strip().lower() if v is not None else "")

                    if not any(scan_values):
                        continue

                    # Не забираем строку с part_no как qty/name
                    has_pn_above = any(
                        not HeuristicAnalyzer._is_false_positive_part_no(sv, kw)
                        for sv in scan_values
                        for kw in PART_NO_KEYWORDS
                        if kw in sv
                    )
                    if has_pn_above:
                        continue

                    if qty_col is None:
                        for col_idx, val in enumerate(scan_values, 1):
                            if any(kw in val for kw in QTY_KEYWORDS):
                                qty_col = col_idx
                                break
                    if name_col is None:
                        for col_idx, val in enumerate(scan_values, 1):
                            if any(kw in val for kw in NAME_KEYWORDS):
                                name_col = col_idx
                                break

                    if qty_col is not None and name_col is not None:
                        logger.debug(
                            "Найдены name/qty на строке %d (row-above header)",
                            scan_row,
                        )
                        break

            result = (
                row_idx,
                part_no_col,
                qty_col or 0,
                name_col or 0,
            )
            logger.debug(
                "Таблица деталей: строка %d, part_no=%s, qty=%s, name=%s",
                row_idx, part_no_col, qty_col, name_col,
            )
            return result

        return None

    @staticmethod
    def is_service_sheet(sheet_name: str) -> bool:
        """Проверить, является ли лист служебным (не BOM, не операционная карта)."""
        name_lower = sheet_name.lower()
        for kw in SERVICE_SHEET_KEYWORDS:
            if kw in name_lower:
                return True
        return False

    @staticmethod
    def is_sheet_bom_candidate(
        ws: Any,
        min_configs: int = 2,
        sheet_name: str = "",
    ) -> bool:
        """Проверить, является ли лист кандидатом на BOM-данные.

        Анализирует:
          - Наличие строки заголовков с part_no
          - Количество строк с данными (> 5)
          - Наличие минимум min_configs колонок комплектаций с числами
          - Имя листа не должно содержать служебных ключевых слов

        Args:
            ws: Лист Excel
            min_configs: Минимальное количество колонок комплектаций.
                         Для основного BOM-листа должно быть >= 2.
                         Спец-листы (零部件附件) могут иметь 1 qty-колонку.
                         Многостраничные BOM (SWM: 涂装/焊装) могут иметь 0-1.
            sheet_name: Имя листа для дополнительной фильтрации.
        """
        # Фильтрация по имени листа
        if sheet_name and HeuristicAnalyzer.is_service_sheet(sheet_name):
            return False

        header_rows = HeuristicAnalyzer.find_header_rows(ws, sheet_name=sheet_name)
        if not header_rows:
            return False

        col_types = HeuristicAnalyzer.detect_column_types(ws, header_rows)
        if 'part_no' not in col_types:
            return False

        # Проверяем количество данных
        data_start = header_rows[-1] + 1
        data_rows = (ws.max_row or 0) - data_start + 1
        if data_rows < 3:
            return False

        # Проверяем наличие как минимум min_configs колонок комплектаций
        config_cols = HeuristicAnalyzer.detect_config_columns(ws, header_rows, col_types)
        if len(config_cols) >= min_configs:
            return True

        # Спец-листы (附件) с одной qty-колонкой
        # (в т.ч. когда ВСЕ колонки классифицированы как part_no/name/qty,
        #  и config_cols пуст — qty-колонка сама служит конфигурацией)
        qty_col = col_types.get('qty', 0)
        has_name = 'name_cn' in col_types or 'name_en' in col_types
        if qty_col > 0 and has_name:
            return True

        # Многостраничные BOM (SWM: 涂装BOM/焊装BOM):
        # Если лист содержит part_no + name + qty — принимаем как BOM-лист
        # даже при 0 конфигурационных колонках.
        # Данные будут агрегированы через config_quantities[sheet_name].
        part_no_col = col_types.get('part_no', 0)
        if qty_col > 0 and part_no_col > 0:
            return True

        return False

    @staticmethod
    def analyze_bom_sheet(
        ws: Any,
        min_configs: int = 2,
        sheet_name: str = "",
    ) -> Optional[Tuple[List[int], Dict[str, int], List[int]]]:
        """Analyze a sheet and return (header_rows, col_types, config_cols) if it's a BOM candidate.

        Returns None if the sheet is not a BOM candidate.
        This avoids duplicate calls to find_header_rows/detect_column_types/detect_config_columns.
        """
        if sheet_name and HeuristicAnalyzer.is_service_sheet(sheet_name):
            return None

        header_rows = HeuristicAnalyzer.find_header_rows(ws, sheet_name=sheet_name)
        if not header_rows:
            return None

        col_types = HeuristicAnalyzer.detect_column_types(ws, header_rows)
        if 'part_no' not in col_types:
            return None

        data_start = header_rows[-1] + 1
        data_rows = (ws.max_row or 0) - data_start + 1
        if data_rows < 3:
            return None

        config_cols = HeuristicAnalyzer.detect_config_columns(ws, header_rows, col_types)
        if len(config_cols) >= min_configs:
            return (header_rows, col_types, config_cols)

        qty_col = col_types.get('qty', 0)
        has_name = 'name_cn' in col_types or 'name_en' in col_types
        if qty_col > 0 and has_name:
            return (header_rows, col_types, config_cols)

        part_no_col = col_types.get('part_no', 0)
        if qty_col > 0 and part_no_col > 0:
            return (header_rows, col_types, config_cols)

        return None

        return False

    @staticmethod
    def extract_operation_name(ws: Any, table_header_row: int) -> str:
        """Извлечь название операции из шапки листа (выше таблицы деталей).

        Ищет текстовые строки с китайскими иероглифами,
        не содержащие служебных ключевых слов.
        """
        max_col = min(ws.max_column or 20, 10)
        service_kws = [
            "作业指导书", "文件编号", "工具/夹具", "版本", "发行时间",
            "关键点", "车间", "序号", "变更记录", "物料清单",
            "说明性符号", "编制", "校对", "审核", "批准", "无",
        ]

        for r in range(1, min(table_header_row, 15)):
            for c in range(1, max_col + 1):
                v = HeuristicAnalyzer.get_cell_value(ws, r, c)
                if v is None:
                    continue
                text = str(v).strip()
                # Пропускаем служебные
                if any(kw in text for kw in service_kws):
                    continue
                # Ищем текст с иероглифами
                if len(text) > 3 and bool(CJK_RE.search(text)):
                    # Возвращаем первое неслужебное название
                    if "作业要素" in text:
                        # Ищем рядом
                        for check_c in range(c + 1, min(c + 3, max_col + 1)):
                            nv = HeuristicAnalyzer.get_cell_value(ws, r, check_c)
                            if nv and len(str(nv).strip()) > 1 and "作业要素" not in str(nv):
                                return str(nv).strip()
                    else:
                        return text
        return ""

    @staticmethod
    def build_global_name_dict(
        ws: Any,
        part_no_col: int,
        name_cn_col: int,
        name_en_col: int,
        header_row: int,
    ) -> Dict[str, Tuple[str, str]]:
        """Построить глобальный словарь part_number -> (name_cn, name_en).

        Сканирует ВСЕ строки данных (не только для конкретной комплектации),
        собирая названия для КАЖДОГО найденного парт-номера.

        Args:
            ws: Лист Excel
            part_no_col: Колонка с парт-номерами
            name_cn_col: Колонка с названием (кит.)
            name_en_col: Колонка с названием (англ.)
            header_row: Строка заголовков

        Returns:
            Словарь {part_number: (name_cn, name_en)}
        """
        name_dict: Dict[str, Tuple[str, str]] = {}
        data_start = header_row + 1
        max_row = ws.max_row or data_start

        for row_idx in range(data_start, max_row + 1):
            if HeuristicAnalyzer.is_cell_strike(ws, row_idx, part_no_col):
                continue
            pn = HeuristicAnalyzer.get_cell_value(ws, row_idx, part_no_col)
            if pn is None:
                continue
            pn_str = str(pn).strip()
            if not pn_str or pn_str.startswith("~$"):
                continue
            pn_clean = clean_part_number(pn_str)
            if not pn_clean or len(pn_clean) < 3:
                continue

            name_cn = ""
            if name_cn_col:
                nc = HeuristicAnalyzer.get_cell_value(ws, row_idx, name_cn_col)
                if nc is not None:
                    name_cn = clean_cell_text(nc)

            name_en = ""
            if name_en_col:
                ne = HeuristicAnalyzer.get_cell_value(ws, row_idx, name_en_col)
                if ne is not None:
                    name_en = clean_cell_text(ne)

            # Сохраняем названия (не перезаписываем пустыми)
            if pn_clean in name_dict:
                existing_cn, existing_en = name_dict[pn_clean]
                if not existing_cn and name_cn:
                    existing_cn = name_cn
                if not existing_en and name_en:
                    existing_en = name_en
                name_dict[pn_clean] = (existing_cn, existing_en)
            else:
                name_dict[pn_clean] = (name_cn, name_en)

        return name_dict


# ═══════════════════════════════════════════════════════════════════════
# УДОБНАЯ ФУНКЦИЯ: ИЗВЛЕЧЕНИЕ НОМЕРА КАРТЫ (ИЗ ФАЙЛА ИЛИ СОДЕРЖИМОГО)
# ═══════════════════════════════════════════════════════════════════════

def extract_card_number(file_path: str, ws: Optional[Any] = None) -> str:
    """Извлечь номер операционной карты.

    Если передан ws — сначала пытается извлечь из содержимого листа.
    Fallback: извлекает из имени файла.

    Args:
        file_path: Путь к файлу.
        ws: Лист Excel (опционально).

    Returns:
        Номер карты.
    """
    if ws is not None:
        return HeuristicAnalyzer.extract_card_number_from_sheet(ws, file_path)
    return extract_card_number_from_filepath(file_path)
