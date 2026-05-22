from __future__ import annotations

import argparse
import json
import os
import re
import urllib.error
import urllib.request
from dataclasses import dataclass, asdict
from difflib import SequenceMatcher
from functools import lru_cache
from pathlib import Path
from typing import Callable, Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from pypdf import PdfReader


DEEPSEEK_API_BASE = os.environ.get("DEEPSEEK_API_BASE", "https://api.deepseek.com").rstrip("/")
DEEPSEEK_API_URL = os.environ.get("DEEPSEEK_API_URL", f"{DEEPSEEK_API_BASE}/chat/completions")
DEEPSEEK_MODEL = os.environ.get("DEEPSEEK_MODEL", "deepseek-chat")
AI_WARNINGS: list[str] = []


def clear_ai_warnings() -> None:
    AI_WARNINGS.clear()


def get_ai_warnings() -> list[str]:
    return list(AI_WARNINGS)


def record_ai_warning(message: str) -> None:
    if message not in AI_WARNINGS:
        AI_WARNINGS.append(message)


def deepseek_api_url() -> str:
    api_url = os.environ.get("DEEPSEEK_API_URL", "").strip()
    if api_url:
        return api_url
    api_base = os.environ.get("DEEPSEEK_API_BASE", DEEPSEEK_API_BASE).strip().rstrip("/")
    return f"{api_base}/chat/completions"


def load_env_file(path: Path | None = None) -> None:
    env_path = path or Path(__file__).resolve().parent / ".env"
    if not env_path.exists():
        return
    for line in env_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip().lstrip("\ufeff")
        value = value.strip().strip('"').strip("'")
        if key and not os.environ.get(key):
            os.environ[key] = value


STOP_WORDS = {
    "для",
    "при",
    "под",
    "или",
    "без",
    "над",
    "мм",
    "м2",
    "м3",
    "шт",
    "уп",
    "упак",
    "рул",
    "кг",
    "л",
    "лист",
    "наружная",
    "наружный",
    "серый",
    "белый",
    "оцинкованный",
    "заказ",
    "цвет",
    "расход",
    "аналог",
}

SERVICE_WORDS = {
    "доставка",
    "доставки",
    "транспортные",
    "транспортная",
    "транспорт",
    "услуга",
    "услуги",
    "разгрузка",
    "подъем",
}

SYNONYM_GROUPS = {
    "саморез": {"саморез", "саморезы", "шуруп", "шурупы", "метиз", "метизы", "крепеж", "крепежный", "винт", "клопы", "гм"},
    "гипсокартон": {"гипсокартон", "гкл", "сапфир"},
    "мембрана": {"мембрана", "ветрозащита", "ветро", "влагозащита", "пароизоляция", "пароизоляционная", "гидро", "изоспан"},
    "утеплитель": {"утеплитель", "теплоизоляция", "изоляция", "вата", "минвата", "каменная", "стеклянного"},
    "профиль": {"профиль", "каркас", "планка", "направляющая", "поперечная", "угловой", "j", "l", "f", "омега", "пи"},
    "плита": {"плита", "панель", "панели", "аквапанель", "акустическая", "потолочная"},
    "клей": {"клей", "клеевая", "клеевой", "штукатурно", "смесь", "смеси", "цементная"},
    "грунтовка": {"грунтовка", "грунт", "тифенгрунт", "ct17", "ст17"},
    "гидроизоляция": {"гидроизоляция", "гидроизоляционный", "флехендихт", "флэхендихт", "обмазочная"},
    "пленка": {"пленка", "пленка", "полиэтиленовая", "укрывочная", "микрон", "мкм"},
    "затирка": {"затирка", "затирочная", "ce40", "се40"},
    "сетка": {"сетка", "стеклосетка", "стеклотканевая", "фасадная"},
    "рейка": {"рейка", "рейки", "каркас", "норма", "norma", "т24", "t24", "планка"},
    "гребенка": {"гребенка", "bts", "bt", "вт"},
    "подвес": {"подвес", "подвесов", "нониус", "европодвес"},
    "соединитель": {"соединитель", "соед", "элем", "элемент"},
    "герметик": {"герметик", "силикон", "силиконовый"},
    "плитка": {"плитка", "керамогранит", "керамическая"},
    "краска": {"краска", "окраска", "акриловая"},
    "штукатурка": {"штукатурка", "штукатурки", "декоративная", "камешковая"},
    "водосток": {"водосток", "водосточная", "сток", "слив", "колено", "труба", "муфта", "хомут"},
    "osb": {"osb", "осп", "osb3", "osb-3"},
}

SYNONYM_BY_TOKEN = {
    token: canonical
    for canonical, variants in SYNONYM_GROUPS.items()
    for token in variants
}

KNOWN_BRANDS = {
    "knauf": "кнауф",
    "кнауф": "кнауф",
    "volma": "волма",
    "волма": "волма",
    "grandline": "grandline",
    "грандлайн": "grandline",
    "изоспан": "изоспан",
    "технониколь": "технониколь",
    "техновент": "техновент",
    "церезит": "церезит",
    "ceresit": "церезит",
    "rockfon": "rockfon",
    "рокфон": "rockfon",
    "albes": "albes",
    "албес": "albes",
}


@dataclass
class RequestItem:
    pos: str
    name: str
    specs: str = ""
    unit: str = ""
    qty: float | None = None


@dataclass
class SupplierItem:
    supplier: str
    source: str
    row_no: str
    name: str
    qty: float | None = None
    unit: str = ""
    price: float | None = None
    total: float | None = None
    delivery: str = ""


@dataclass
class Match:
    supplier_item: SupplierItem
    request_pos: str | None
    score: float
    status: str
    reason: str


STATUS_LABELS = {
    "auto": "автоматически",
    "review": "на проверку",
    "unmatched": "не сопоставлено",
    "manual": "подтверждено на проверке",
    "service": "доставка/услуга",
}


def status_label(status: str) -> str:
    return STATUS_LABELS.get(status, status)


@lru_cache(maxsize=20000)
def fix_mojibake(text: str) -> str:
    if not text or not re.search(r"(Р[°-џ]|\u0098|вЂ)", text):
        return text
    try:
        raw = bytearray()
        for char in text:
            code = ord(char)
            if code <= 255:
                raw.append(code)
            else:
                raw.extend(char.encode("cp1251"))
        fixed = raw.decode("utf-8")
    except (UnicodeEncodeError, UnicodeDecodeError, ValueError):
        return text

    bad_before = len(re.findall(r"(Р[°-џ]|\u0098|вЂ)", text))
    bad_after = len(re.findall(r"(Р[°-џ]|\u0098|вЂ)", fixed))
    if bad_after < bad_before:
        return fixed
    return text


@lru_cache(maxsize=40000)
def clean_text(value) -> str:
    if value is None:
        return ""
    text = re.sub(r"\s+", " ", str(value).replace("\xa0", " ")).strip()
    for _ in range(2):
        fixed = fix_mojibake(text)
        if fixed == text:
            break
        text = fixed
    return text


def parse_number(value) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = clean_text(value)
    text = text.replace(" ", "").replace("\u202f", "").replace(",", ".")
    text = re.sub(r"[^\d.\-]", "", text)
    if not text or text in {"-", ".", "-."}:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def money_to_float(text: str) -> float | None:
    return parse_number(text)


@lru_cache(maxsize=40000)
def normalize(text: str) -> str:
    text = clean_text(text).lower().replace("ё", "е")
    text = text.replace("x", "х").replace("*", "х")
    text = re.sub(r"(?<=\d),(?=\d)", ".", text)
    text = re.sub(r"[^0-9a-zа-я.х]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


@lru_cache(maxsize=40000)
def tokens(text: str) -> frozenset[str]:
    base_tokens = {
        token
        for token in normalize(text).split()
        if len(token) > 1 and token not in STOP_WORDS
    }
    expanded = set(base_tokens)
    for token in base_tokens:
        canonical = SYNONYM_BY_TOKEN.get(token)
        if canonical:
            expanded.add(canonical)
    return frozenset(expanded)


@lru_cache(maxsize=40000)
def is_service_text(text: str) -> bool:
    text_tokens = tokens(text)
    normalized = normalize(text)
    if text_tokens & SERVICE_WORDS:
        return True
    return any(marker in normalized for marker in ["доставка", "транспортные услуги", "транспортная услуга"])


def is_service_item(item: SupplierItem) -> bool:
    return is_service_text(item.name)


@lru_cache(maxsize=40000)
def item_categories(text: str) -> frozenset[str]:
    return frozenset(SYNONYM_BY_TOKEN[token] for token in tokens(text) if token in SYNONYM_BY_TOKEN)


def normalized_unit(unit: str) -> str:
    raw = clean_text(unit).lower().replace("²", "2").replace("^2", "2")
    raw = raw.replace(" ", "").replace(".", "")
    if raw in {"м2", "квм", "квм2"}:
        return "m2"
    if raw in {"м", "мп", "пм"}:
        return "m"
    if raw in {"шт", "штук"}:
        return "pcs"
    if raw in {"лист", "листы"}:
        return "sheet"
    if raw in {"рул", "рулон", "рулоны"}:
        return "roll"
    if raw in {"уп", "упак", "упаковка"}:
        return "pack"
    if raw in {"кг"}:
        return "kg"
    unit_norm = normalize(unit)
    if unit_norm in {"м2", "м.2", "кв.м", "кв", "м²"}:
        return "m2"
    if unit_norm in {"м", "м.п", "п.м", "мп"}:
        return "m"
    if unit_norm in {"шт", "штук"}:
        return "pcs"
    if unit_norm in {"лист", "листы"}:
        return "sheet"
    if unit_norm in {"рул", "рулон", "рулоны"}:
        return "roll"
    if unit_norm in {"уп", "упак", "упаковка"}:
        return "pack"
    if unit_norm in {"кг"}:
        return "kg"
    return unit_norm


@lru_cache(maxsize=40000)
def area_m2_from_text(text: str) -> float | None:
    normalized = normalize(text)
    area_patterns = [
        r"(?:s|площадь)\s*=?\s*(\d+(?:\.\d+)?)\s*(?:м2|м²|кв\.?м)",
        r"(\d+(?:\.\d+)?)\s*(?:м2|м²|кв\.?м)",
    ]
    for pattern in area_patterns:
        match = re.search(pattern, normalized)
        if match:
            value = parse_number(match.group(1))
            if value:
                return value
    dim = re.search(r"(\d{3,4})\s*х\s*(\d{3,4})", normalized)
    if dim:
        first = parse_number(dim.group(1))
        second = parse_number(dim.group(2))
        if first and second:
            return round(first * second / 1_000_000, 4)
    return None


def quantities_compatible(request: RequestItem, offer: SupplierItem) -> bool:
    if request.qty is None or offer.qty is None:
        return False
    req_unit = normalized_unit(request.unit)
    offer_unit = normalized_unit(offer.unit)
    if req_unit == offer_unit:
        return abs(request.qty - offer.qty) <= max(0.01, request.qty * 0.02)
    req_area = area_m2_from_text(request.name)
    offer_area = area_m2_from_text(offer.name)
    if req_unit == "m2" and offer_unit in {"sheet", "roll", "pack", "pcs"} and offer_area:
        return abs(request.qty - offer.qty * offer_area) <= max(0.05, request.qty * 0.12)
    if offer_unit == "m2" and req_unit in {"sheet", "roll", "pack", "pcs"} and req_area:
        return abs(request.qty * req_area - offer.qty) <= max(0.05, offer.qty * 0.12)
    return False


@lru_cache(maxsize=40000)
def numeric_tokens(text: str) -> frozenset[str]:
    return frozenset(m.group(0).replace(",", ".") for m in re.finditer(r"\d+(?:[,.]\d+)?", text))


@lru_cache(maxsize=40000)
def brands(text: str) -> frozenset[str]:
    source = normalize(text)
    found = set()
    for raw, canonical in KNOWN_BRANDS.items():
        if raw in source:
            found.add(canonical)
    return frozenset(found)


def match_score(request_name: str, supplier_name: str) -> tuple[float, str]:
    req_norm = normalize(request_name)
    sup_norm = normalize(supplier_name)
    seq = SequenceMatcher(None, req_norm, sup_norm).ratio()

    req_tokens = tokens(request_name)
    sup_tokens = tokens(supplier_name)
    union = req_tokens | sup_tokens
    jaccard = len(req_tokens & sup_tokens) / len(union) if union else 0

    req_nums = numeric_tokens(request_name)
    sup_nums = numeric_tokens(supplier_name)
    num_union = req_nums | sup_nums
    num_score = len(req_nums & sup_nums) / len(num_union) if num_union else 0.4
    req_categories = item_categories(request_name)
    sup_categories = item_categories(supplier_name)
    category_overlap = bool(req_categories & sup_categories)

    req_brands = brands(request_name)
    sup_brands = brands(supplier_name)
    brand_conflict = bool(req_brands and sup_brands and req_brands.isdisjoint(sup_brands))

    score = 0.42 * seq + 0.43 * jaccard + 0.15 * num_score
    if category_overlap:
        score += 0.14
    if brand_conflict:
        score -= 0.22

    reasons = []
    if brand_conflict:
        reasons.append(f"разные бренды: {', '.join(sorted(req_brands))} / {', '.join(sorted(sup_brands))}")
    if score < 0.62:
        reasons.append("низкая уверенность сопоставления")
    return max(0, min(1, score)), "; ".join(reasons)


def find_header_row(ws, required_terms: Iterable[str]) -> tuple[int, dict[str, int]] | None:
    required = [term.lower() for term in required_terms]
    for row in ws.iter_rows():
        values = [clean_text(cell.value).lower() for cell in row]
        joined = " ".join(values)
        if all(term in joined for term in required):
            return row[0].row, {values[i]: i + 1 for i in range(len(values)) if values[i]}
    return None


def find_column(ws, header_row: int, variants: Iterable[str]) -> int | None:
    variants = [v.lower() for v in variants]
    for cell in ws[header_row]:
        value = clean_text(cell.value).lower()
        if any(variant in value for variant in variants):
            return cell.column
    return None


def find_column_preferred(
    ws,
    header_row: int,
    exact: Iterable[str],
    contains: Iterable[str],
    exclude: Iterable[str] = (),
) -> int | None:
    exact_values = [value.lower() for value in exact]
    contains_values = [value.lower() for value in contains]
    exclude_values = [value.lower() for value in exclude]
    fallback: int | None = None
    for cell in ws[header_row]:
        value = clean_text(cell.value).lower()
        if not value or any(marker in value for marker in exclude_values):
            continue
        if value in exact_values:
            return cell.column
        if fallback is None and any(marker in value for marker in contains_values):
            fallback = cell.column
    return fallback


def read_request_xlsx(path: Path) -> list[RequestItem]:
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    found = find_header_row(ws, ["описание", "количество"])
    if not found:
        found = find_header_row(ws, ["описание", "объем"])
    if not found:
        raise ValueError(f"Не нашел строку заголовков заявки в {path.name}")

    header_row, _ = found
    pos_col = find_column(ws, header_row, ["№", "номер"]) or 1
    name_col = find_column(ws, header_row, ["описание закупаемой", "описание"])
    specs_col = find_column(ws, header_row, ["технические характеристики", "гост"])
    unit_col = find_column(ws, header_row, ["ед. измерения", "ед измерения"])
    qty_col = find_column(ws, header_row, ["необходимый объем", "количество"])

    if not name_col or not qty_col:
        raise ValueError(f"Не нашел колонки описания/количества в заявке {path.name}")

    items: list[RequestItem] = []
    for row in range(header_row + 1, ws.max_row + 1):
        pos = clean_text(ws.cell(row, pos_col).value)
        name = clean_text(ws.cell(row, name_col).value)
        unit = clean_text(ws.cell(row, unit_col).value) if unit_col else ""
        qty = parse_number(ws.cell(row, qty_col).value) if qty_col else None
        specs = clean_text(ws.cell(row, specs_col).value) if specs_col else ""
        if not name:
            continue
        if qty is None and not unit:
            continue
        if not pos:
            pos = str(len(items) + 1)
        items.append(RequestItem(pos=pos, name=name, specs=specs, unit=unit, qty=qty))
    return items


def supplier_name_from_file(path: Path) -> str:
    stem = path.stem
    known = {
        "петрович": "Петрович",
        "petrovich": "Петрович",
        "авангард": "Авангард-строй",
        "avangard": "Авангард-строй",
        "1350": "Авангард-строй",
        "еврострой": "Евростройгрупп",
        "eurostroy": "Евростройгрупп",
        "euro": "Евростройгрупп",
        "грандлайн": "Грандлайн",
        "grandline": "Грандлайн",
    }
    lower = stem.lower()
    for marker, name in known.items():
        if marker in lower:
            return name
    return stem[:45]


def read_supplier_xlsx(path: Path, supplier: str | None = None) -> list[SupplierItem]:
    wb = load_workbook(path, data_only=True)
    supplier = supplier or supplier_name_from_file(path)
    items: list[SupplierItem] = []

    for ws in wb.worksheets:
        table_header = find_header_row(ws, ["товар", "кол-во"])
        if table_header:
            header_row, _ = table_header
            name_col = find_column_preferred(
                ws,
                header_row,
                exact=["товар", "наименование"],
                contains=["номенклатура", "наименование", "товар"],
                exclude=["код товара", "код"],
            )
            qty_col = find_column_preferred(ws, header_row, exact=["кол-во", "количество"], contains=["кол-во", "количество"])
            unit_col = find_column_preferred(ws, header_row, exact=["ед.", "ед"], contains=["ед."])
            price_col = find_column_preferred(ws, header_row, exact=["цена"], contains=["цена"])
            total_col = find_column(ws, header_row, ["сумма", "стоимость"])
            pos_col = find_column(ws, header_row, ["№"]) or 1
            if name_col and qty_col and price_col:
                for row in range(header_row + 1, ws.max_row + 1):
                    name = clean_text(ws.cell(row, name_col).value)
                    if not name or "итого" in name.lower() or "всего" in name.lower():
                        continue
                    price = parse_number(ws.cell(row, price_col).value)
                    total = parse_number(ws.cell(row, total_col).value) if total_col else None
                    if price is None and total is None:
                        continue
                    items.append(
                        SupplierItem(
                            supplier=supplier,
                            source=path.name,
                            row_no=clean_text(ws.cell(row, pos_col).value) or str(row),
                            name=name,
                            qty=parse_number(ws.cell(row, qty_col).value),
                            unit=clean_text(ws.cell(row, unit_col).value) if unit_col else "",
                            price=price,
                            total=total,
                        )
                    )
                continue

        request_like_header = find_header_row(ws, ["описание", "предельная цена"])
        if request_like_header:
            header_row, _ = request_like_header
            pos_col = find_column(ws, header_row, ["№"]) or 1
            name_col = find_column(ws, header_row, ["описание закупаемой", "описание"])
            qty_col = find_column(ws, header_row, ["необходимый объем", "количество"])
            unit_col = find_column(ws, header_row, ["ед. измерения", "ед измерения"])
            price_col = find_column(ws, header_row, ["предельная цена", "цена"])
            total_col = find_column(ws, header_row, ["стоимость", "сумма"])
            delivery_col = find_column(ws, header_row, ["срок"])
            if name_col and price_col:
                for row in range(header_row + 1, ws.max_row + 1):
                    name = clean_text(ws.cell(row, name_col).value)
                    price = parse_number(ws.cell(row, price_col).value)
                    total = parse_number(ws.cell(row, total_col).value) if total_col else None
                    if not name or (price is None and total is None):
                        continue
                    items.append(
                        SupplierItem(
                            supplier=supplier,
                            source=path.name,
                            row_no=clean_text(ws.cell(row, pos_col).value) or str(row),
                            name=name,
                            qty=parse_number(ws.cell(row, qty_col).value) if qty_col else None,
                            unit=clean_text(ws.cell(row, unit_col).value) if unit_col else "",
                            price=price,
                            total=total,
                            delivery=clean_text(ws.cell(row, delivery_col).value) if delivery_col else "",
                        )
                    )
    return items


ROW_RE = re.compile(
    r"^\s*(?P<row>\d+)\s+(?:(?P<code>\d{4,}|[0-9]{4,}\s+заказ)\s+)?"
    r"(?P<name>.*?)\s+(?P<qty>\d[\d\s]*(?:[,.]\d+)?)\s+"
    r"(?P<unit>[A-Za-zА-Яа-яЁё.]+)\s+"
    r"(?P<money>\d[\d\s]*,\d{2}(?:\s+\d[\d\s]*,\d{2}){1,3})\s*$"
)


def clean_pdf_name_line(line: str) -> str:
    line = clean_text(line)
    line = re.sub(r"^\d{5,}\s+", "", line)
    line = re.sub(r"^\d{5,}\s+заказ\s+", "", line)
    return line


def is_product_continuation(line: str) -> bool:
    text = clean_text(line)
    if not text:
        return False
    lower = text.lower()
    if any(marker in lower for marker in ["итого", "ндс", "всего к оплате", "страница", "поставщик", "покупатель"]):
        return False
    if re.search(r"\d[\d\s]*,\d{2}\s+\d[\d\s]*,\d{2}", text):
        return False
    return bool(re.search(r"[A-Za-zА-Яа-яЁё]", text))


def parse_layout_pdf(path: Path, supplier: str | None = None) -> list[SupplierItem]:
    supplier = supplier or supplier_name_from_file(path)
    reader = PdfReader(str(path))
    text_parts = []
    for page in reader.pages:
        try:
            text_parts.append(page.extract_text(extraction_mode="layout") or "")
        except TypeError:
            text_parts.append(page.extract_text() or "")
    lines = "\n".join(text_parts).splitlines()

    delivery = ""
    for line in lines:
        if "срок готовности" in line.lower() or "срок поставки" in line.lower():
            parts = re.split(r":", line, maxsplit=1)
            delivery = clean_text(parts[-1]) if len(parts) > 1 else clean_text(line)
            break

    items: list[SupplierItem] = []
    in_table = False
    pending: list[str] = []
    current: SupplierItem | None = None

    def flush_pending_to_current() -> None:
        nonlocal pending, current
        if current and pending:
            current.name = clean_text(current.name + " " + " ".join(clean_pdf_name_line(x) for x in pending))
        pending = []

    for line in lines:
        lower = line.lower()
        if "товары" in lower and "сумма" in lower:
            in_table = True
            pending = []
            continue
        if not in_table:
            continue
        if "итого:" in lower or "всего к оплате" in lower:
            break

        match = ROW_RE.match(line)
        if match:
            if current:
                items.append(current)
            money_values = [money_to_float(x) for x in re.findall(r"\d[\d\s]*,\d{2}", match.group("money"))]
            money_values = [x for x in money_values if x is not None]
            price = money_values[1] if len(money_values) >= 3 else (money_values[0] if money_values else None)
            total = money_values[-1] if money_values else None
            name_parts = [clean_pdf_name_line(x) for x in pending]
            name_parts.append(clean_pdf_name_line(match.group("name")))
            current = SupplierItem(
                supplier=supplier,
                source=path.name,
                row_no=match.group("row"),
                name=clean_text(" ".join(name_parts)),
                qty=parse_number(match.group("qty")),
                unit=clean_text(match.group("unit")),
                price=price,
                total=total,
                delivery=delivery,
            )
            pending = []
            continue

        if current and is_product_continuation(line):
            current.name = clean_text(current.name + " " + clean_pdf_name_line(line))
        elif is_product_continuation(line):
            pending.append(line)

    if current:
        items.append(current)
    return items


def read_offer(path: Path, supplier: str | None = None) -> list[SupplierItem]:
    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        return read_supplier_xlsx(path, supplier)
    if suffix == ".pdf":
        return parse_layout_pdf(path, supplier)
    if suffix == ".xls":
        raise ValueError(
            f"{path.name}: старый .xls не поддержан текущим Python-окружением. "
            "Сохраните файл как .xlsx или PDF и запустите повторно."
        )
    raise ValueError(f"{path.name}: неподдержанный формат КП")


def candidate_request_indexes(offer: SupplierItem, request_items: list[RequestItem], token_index: dict[str, set[int]], num_index: dict[str, set[int]]) -> set[int]:
    offer_tokens = tokens(offer.name)
    offer_nums = numeric_tokens(offer.name)
    candidates: set[int] = set()
    for token in offer_tokens:
        candidates.update(token_index.get(token, set()))
    for number in offer_nums:
        candidates.update(num_index.get(number, set()))

    if not candidates:
        return set(range(len(request_items)))

    req_brands_by_index = {idx: brands(request_items[idx].name) for idx in candidates}
    offer_brands = brands(offer.name)
    if offer_brands:
        brand_matches = {idx for idx, req_brands in req_brands_by_index.items() if req_brands and not req_brands.isdisjoint(offer_brands)}
        if brand_matches:
            candidates = brand_matches

    if len(candidates) < 6:
        number_matches = set()
        for number in offer_nums:
            number_matches.update(num_index.get(number, set()))
        candidates.update(number_matches)
    return candidates or set(range(len(request_items)))


def call_deepseek(payload: dict) -> dict | None:
    api_key = os.environ.get("DEEPSEEK_API_KEY", "").strip()
    if not api_key:
        return None

    body = {
        "model": os.environ.get("DEEPSEEK_MODEL", DEEPSEEK_MODEL),
        "messages": [
            {
                "role": "system",
                "content": (
                    "Ты сопоставляешь строительные материалы из заявки и КП. "
                    "Учитывай синонимы, бренды, размеры, единицы измерения и смысл товара. "
                    "Примеры: метиз/шуруп/саморез/крепеж близкие группы; ГКЛ/гипсокартон близкие; "
                    "лист, рулон и упаковка могут соответствовать м2, если площадь указана в названии. "
                    "Не сопоставляй разные бренды, если бренд принципиален. "
                    "Ответь только JSON без пояснений."
                ),
            },
            {
                "role": "user",
                "content": json.dumps(
                    {
                        "task": (
                            "Для каждой строки offer_positions выбери pos из request_positions или null. "
                            "Если точного совпадения нет, но есть 2-5 вероятных вариантов, верни лучший request_pos и добавь alternatives. "
                            "Верни JSON: {\"matches\":[{\"offer_id\": число, \"request_pos\": строка или null, "
                            "\"confidence\": число от 0 до 1, \"reason\": коротко, "
                            "\"alternatives\":[{\"request_pos\": строка, \"confidence\": число, \"reason\": коротко}]}]}"
                        ),
                        **payload,
                    },
                    ensure_ascii=False,
                ),
            },
        ],
        "response_format": {"type": "json_object"},
        "temperature": 0,
    }
    request = urllib.request.Request(
        deepseek_api_url(),
        data=json.dumps(body, ensure_ascii=False).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    timeout = float(os.environ.get("DEEPSEEK_TIMEOUT", "25"))
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            raw = response.read().decode("utf-8")
    except urllib.error.HTTPError as exc:
        if exc.code == 402:
            record_ai_warning("DeepSeek не выполнил проверку: на аккаунте нет оплаченного баланса.")
        elif exc.code in {401, 403}:
            record_ai_warning("DeepSeek не выполнил проверку: API-ключ не принят сервисом.")
        else:
            record_ai_warning(f"DeepSeek не выполнил проверку: API вернул ошибку {exc.code}.")
        return None
    except (urllib.error.URLError, TimeoutError, OSError, ValueError):
        record_ai_warning("DeepSeek не выполнил проверку: нет соединения с API или истекло время ожидания.")
        return None

    try:
        content = json.loads(raw)["choices"][0]["message"]["content"]
        return json.loads(content)
    except (KeyError, IndexError, TypeError, json.JSONDecodeError):
        record_ai_warning("DeepSeek ответил в неожиданном формате, проверка ИИ пропущена.")
        return None


ProgressCallback = Callable[[int, int, str], None]


def improve_matches_with_deepseek(
    request_items: list[RequestItem],
    matches: list[Match],
    progress_callback: ProgressCallback | None = None,
) -> list[Match]:
    if not os.environ.get("DEEPSEEK_API_KEY", "").strip():
        return matches

    request_payload = [
        {
            "pos": item.pos,
            "name": item.name,
            "specs": item.specs,
            "unit": item.unit,
            "qty": item.qty,
            "area_m2_hint": area_m2_from_text(item.name),
        }
        for item in request_items
    ]
    by_pos = {item.pos: item for item in request_items}
    ai_scope = os.environ.get("DEEPSEEK_AI_SCOPE", "all").strip().lower()
    if ai_scope == "all":
        need_ai = [(idx, match) for idx, match in enumerate(matches) if match.status != "service"]
    else:
        need_ai = [
            (idx, match)
            for idx, match in enumerate(matches)
            if match.status != "service" and (match.status == "review" or not match.request_pos)
        ]
    max_rows = int(os.environ.get("DEEPSEEK_MAX_AI_ROWS", "300"))
    batch_size = int(os.environ.get("DEEPSEEK_BATCH_SIZE", "20"))
    updated = list(matches)
    total_to_check = min(len(need_ai), max_rows)
    if progress_callback:
        progress_callback(0, total_to_check, "ИИ готовит позиции к проверке")
    if len(need_ai) > max_rows:
        record_ai_warning(f"DeepSeek проверил только первые {max_rows} строк КП из {len(need_ai)} по текущему лимиту.")

    for start in range(0, total_to_check, batch_size):
        batch = need_ai[start : start + batch_size]
        offer_payload = [
            {
                "offer_id": idx,
                "supplier": match.supplier_item.supplier,
                "row_no": match.supplier_item.row_no,
                "name": match.supplier_item.name,
                "unit": match.supplier_item.unit,
                "qty": match.supplier_item.qty,
                "area_m2_hint": area_m2_from_text(match.supplier_item.name),
                "price": match.supplier_item.price,
                "current_request_pos": match.request_pos,
                "current_score": round(match.score, 3),
            }
            for idx, match in batch
        ]
        result = call_deepseek({"request_positions": request_payload, "offer_positions": offer_payload})
        done_count = min(start + len(batch), total_to_check)
        if progress_callback:
            progress_callback(done_count, total_to_check, f"ИИ проверил {done_count} из {total_to_check} строк КП")
        if not result:
            continue
        for item in result.get("matches", []):
            try:
                offer_id = int(item.get("offer_id"))
            except (TypeError, ValueError):
                continue
            if offer_id < 0 or offer_id >= len(updated):
                continue
            request_pos = clean_text(item.get("request_pos"))
            confidence = parse_number(item.get("confidence")) or 0
            if confidence > 1:
                confidence /= 100
            alternatives = []
            for alt in item.get("alternatives") or []:
                alt_pos = clean_text(alt.get("request_pos"))
                alt_confidence = parse_number(alt.get("confidence")) or 0
                if alt_confidence > 1:
                    alt_confidence /= 100
                alt_reason = clean_text(alt.get("reason") or "")
                if alt_pos in by_pos:
                    alternatives.append((alt_pos, alt_confidence, alt_reason))
            alternatives = sorted(alternatives, key=lambda alt: alt[1], reverse=True)[:5]
            if (not request_pos or request_pos not in by_pos) and alternatives:
                request_pos, confidence, _ = alternatives[0]
            if not request_pos or request_pos not in by_pos or confidence < 0.38:
                continue
            old = updated[offer_id]
            reason = clean_text(item.get("reason") or "")
            alt_text = "; ".join(
                f"{pos} ({round(conf * 100)}%)" for pos, conf, _ in alternatives if pos != request_pos
            )
            if alt_text:
                reason = clean_text(f"{reason}. Другие варианты: {alt_text}")
            if confidence >= 0.72:
                updated[offer_id] = Match(old.supplier_item, request_pos, confidence, "auto", "")
            else:
                updated[offer_id] = Match(
                    old.supplier_item,
                    request_pos,
                    confidence,
                    "review",
                    f"ИИ предлагает проверить: {reason}" if reason else "ИИ предлагает проверить совпадение",
                )
    return updated


def build_matches(
    request_items: list[RequestItem],
    supplier_items: list[SupplierItem],
    progress_callback: ProgressCallback | None = None,
) -> list[Match]:
    clear_ai_warnings()
    load_env_file()
    matches: list[Match] = []
    token_index: dict[str, set[int]] = {}
    num_index: dict[str, set[int]] = {}
    for idx, request in enumerate(request_items):
        for token in tokens(request.name):
            token_index.setdefault(token, set()).add(idx)
        for number in numeric_tokens(request.name):
            num_index.setdefault(number, set()).add(idx)

    for offer in supplier_items:
        if is_service_item(offer):
            matches.append(Match(offer, None, 0, "service", "доставка или услуга: учитывается отдельно и не входит в процент сопоставления товаров"))
            continue
        best_item = None
        best_score = 0.0
        best_reason = ""
        for idx in candidate_request_indexes(offer, request_items, token_index, num_index):
            request = request_items[idx]
            score, reason = match_score(request.name, offer.name)
            if score > best_score:
                best_score = score
                best_item = request
                best_reason = reason
        if best_item and best_score >= 0.36:
            status = "auto" if best_score >= 0.62 and not best_reason else "review"
            matches.append(Match(offer, best_item.pos, best_score, status, best_reason))
        elif best_item and (
            best_score >= 0.26
            or (best_score >= 0.20 and not item_categories(best_item.name).isdisjoint(item_categories(offer.name)))
        ):
            matches.append(Match(offer, best_item.pos, best_score, "review", best_reason or "возможное совпадение, требуется проверка"))
        else:
            matches.append(Match(offer, None, best_score, "unmatched", "не найдено надежное совпадение"))
    return improve_matches_with_deepseek(request_items, matches, progress_callback)


def write_review(path: Path, matches: list[Match], request_items: list[RequestItem]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Проверка"
    headers = [
        "Статус",
        "Позиция заявки (исправить при необходимости)",
        "Позиция заявки - текст",
        "Поставщик",
        "Строка КП",
        "Позиция КП",
        "Кол-во КП",
        "Ед.",
        "Цена",
        "Сумма",
        "Уверенность",
        "Причина",
        "Источник",
    ]
    ws.append(headers)
    by_pos = {item.pos: item for item in request_items}
    for match in matches:
        request = by_pos.get(match.request_pos or "")
        ws.append(
            [
                status_label(match.status),
                match.request_pos or "",
                request.name if request else "",
                match.supplier_item.supplier,
                match.supplier_item.row_no,
                match.supplier_item.name,
                match.supplier_item.qty,
                match.supplier_item.unit,
                match.supplier_item.price,
                match.supplier_item.total,
                round(match.score, 3),
                match.reason,
                match.supplier_item.source,
            ]
        )
    style_sheet(ws)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    widths = [14, 22, 48, 22, 12, 62, 12, 8, 12, 14, 12, 34, 32]
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    wb.save(path)


def read_review_overrides(path: Path) -> dict[tuple[str, str, str], str | None]:
    wb = load_workbook(path, data_only=True)
    ws = wb["Проверка"] if "Проверка" in wb.sheetnames else wb.active
    overrides: dict[tuple[str, str, str], str | None] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[3]:
            continue
        request_pos = clean_text(row[1])
        supplier = clean_text(row[3])
        row_no = clean_text(row[4])
        source = clean_text(row[12])
        key = (supplier, row_no, source)
        if request_pos.lower() in {"", "-", "нет", "skip", "не сравнивать"}:
            overrides[key] = None
        else:
            overrides[key] = request_pos
    return overrides


def apply_overrides(matches: list[Match], overrides: dict[tuple[str, str, str], str | None]) -> list[Match]:
    result: list[Match] = []
    for match in matches:
        key = (match.supplier_item.supplier, match.supplier_item.row_no, match.supplier_item.source)
        if key in overrides:
            request_pos = overrides[key]
            status = "manual" if request_pos else "unmatched"
            reason = "подтверждено на проверке" if request_pos else "оставлено без сопоставления"
            result.append(Match(match.supplier_item, request_pos, match.score, status, reason))
        else:
            result.append(match)
    return result


def style_sheet(ws) -> None:
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    thin = Side(style="thin", color="B7B7B7")
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(bold=True)


def fmt_qty(value: float | None) -> str | float:
    return "" if value is None else value


def delivery_mark(delivery: str) -> str:
    text = clean_text(delivery)
    if not text:
        return "-"
    if "налич" in text.lower() or "склад" in text.lower():
        return "✓"
    return text


def write_final(path: Path, request_items: list[RequestItem], matches: list[Match]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Сводка"

    suppliers = sorted({m.supplier_item.supplier for m in matches})
    by_request_supplier: dict[tuple[str, str], list[Match]] = {}
    unmatched: list[SupplierItem] = []
    service_items: list[SupplierItem] = []
    for match in matches:
        if match.request_pos:
            by_request_supplier.setdefault((match.request_pos, match.supplier_item.supplier), []).append(match)
        elif match.status == "service":
            service_items.append(match.supplier_item)
        else:
            unmatched.append(match.supplier_item)

    left_header = PatternFill("solid", fgColor="9FB6C9")
    supplier_fills = [
        PatternFill("solid", fgColor="D9D9D9"),
        PatternFill("solid", fgColor="B6D7A8"),
        PatternFill("solid", fgColor="C9DAF8"),
        PatternFill("solid", fgColor="FCE5CD"),
        PatternFill("solid", fgColor="D9EAD3"),
    ]
    supplier_light_fills = [
        PatternFill("solid", fgColor="EFEFEF"),
        PatternFill("solid", fgColor="E2F0D9"),
        PatternFill("solid", fgColor="EAF2FF"),
        PatternFill("solid", fgColor="FFF2E5"),
        PatternFill("solid", fgColor="EEF7E8"),
    ]
    green = PatternFill("solid", fgColor="C6EFCE")
    red = PatternFill("solid", fgColor="FFC7CE")
    yellow = PatternFill("solid", fgColor="FFF2CC")
    gray = PatternFill("solid", fgColor="F2F2F2")
    white = PatternFill("solid", fgColor="FFFFFF")
    thin = Side(style="thin", color="8D99A6")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def style_range(row_from: int, row_to: int, col_from: int, col_to: int, fill: PatternFill | None = None) -> None:
        for sheet_row in ws.iter_rows(min_row=row_from, max_row=row_to, min_col=col_from, max_col=col_to):
            for cell in sheet_row:
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                if fill:
                    cell.fill = fill

    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    ws.merge_cells(start_row=1, start_column=2, end_row=2, end_column=2)
    ws.merge_cells(start_row=1, start_column=3, end_row=2, end_column=3)
    ws.cell(1, 1, "позиция\nв заявке")
    ws.cell(1, 2, "Описание закупаемой позиции")
    ws.cell(1, 3, "кол-во")
    style_range(1, 2, 1, 3, left_header)
    for col_idx in range(1, 4):
        ws.cell(1, col_idx).font = Font(bold=True, color="1F2933")

    col = 4
    supplier_start_cols: dict[str, int] = {}
    supplier_fill_by_name: dict[str, PatternFill] = {}
    supplier_light_by_name: dict[str, PatternFill] = {}
    for idx, supplier in enumerate(suppliers):
        supplier_start_cols[supplier] = col
        header_fill = supplier_fills[idx % len(supplier_fills)]
        light_fill = supplier_light_fills[idx % len(supplier_light_fills)]
        supplier_fill_by_name[supplier] = header_fill
        supplier_light_by_name[supplier] = light_fill
        ws.merge_cells(start_row=1, start_column=col, end_row=2, end_column=col + 4)
        ws.cell(1, col, supplier)
        ws.cell(1, col).font = Font(bold=True, color="1F2933", size=12)
        style_range(1, 2, col, col + 4, header_fill)
        col += 5

    ws.row_dimensions[1].height = 42
    ws.row_dimensions[2].height = 42
    row = 3
    for request in request_items:
        desc_row = row
        label_row = row + 1
        value_row = row + 2

        ws.cell(label_row, 1, "позиция\nв заявке")
        ws.cell(value_row, 1, request.pos)
        ws.merge_cells(start_row=desc_row, start_column=2, end_row=value_row, end_column=2)
        ws.cell(desc_row, 2, request.name)
        ws.merge_cells(start_row=desc_row, start_column=3, end_row=value_row, end_column=3)
        qty_text = fmt_qty(request.qty)
        ws.cell(desc_row, 3, f"{qty_text} {request.unit}".strip())
        style_range(desc_row, value_row, 1, 3, white)
        ws.cell(label_row, 1).font = Font(size=7)
        ws.cell(value_row, 1).font = Font(size=10)
        ws.cell(desc_row, 2).font = Font(bold=True)
        ws.cell(desc_row, 2).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.cell(desc_row, 3).font = Font(bold=True)

        prices: list[tuple[str, float]] = []
        selected: dict[str, Match] = {}
        for supplier in suppliers:
            candidates = by_request_supplier.get((request.pos, supplier), [])
            if candidates:
                chosen = sorted(
                    candidates,
                    key=lambda item: (item.status != "manual", -(item.supplier_item.price or 0), -item.score),
                )[0]
                selected[supplier] = chosen
                if chosen.supplier_item.price is not None:
                    prices.append((supplier, chosen.supplier_item.price))

        min_price = min((price for _, price in prices), default=None)
        max_price = max((price for _, price in prices), default=None)

        for supplier in suppliers:
            start = supplier_start_cols[supplier]
            match = selected.get(supplier)
            style_range(desc_row, value_row, start, start + 4, supplier_light_by_name[supplier])
            if not match:
                ws.merge_cells(start_row=desc_row, start_column=start, end_row=value_row, end_column=start + 4)
                cell = ws.cell(desc_row, start, "нет в КП")
                cell.font = Font(italic=True, color="667085", size=10)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.fill = gray
                continue
            subheaders = ["позиция\nв счете", "цена", "кол-во", "стоимость", "срок"]
            for offset, header in enumerate(subheaders):
                ws.cell(label_row, start + offset, header)
                ws.cell(label_row, start + offset).font = Font(bold=True, size=9)
            offer = match.supplier_item
            ws.merge_cells(start_row=desc_row, start_column=start, end_row=desc_row, end_column=start + 4)
            ws.cell(desc_row, start, offer.name)
            ws.cell(value_row, start, offer.row_no)
            ws.cell(value_row, start + 1, offer.price)
            offer_qty = fmt_qty(offer.qty)
            ws.cell(value_row, start + 2, f"{offer_qty} {offer.unit}".strip())
            ws.cell(value_row, start + 3, offer.total)
            ws.cell(value_row, start + 4, delivery_mark(offer.delivery))
            qty_ok = quantities_compatible(request, offer)
            if offer.price is not None and min_price is not None and abs(offer.price - min_price) < 0.0001:
                ws.cell(value_row, start + 1).fill = green
            if offer.price is not None and max_price is not None and abs(offer.price - max_price) < 0.0001 and max_price != min_price:
                ws.cell(value_row, start + 1).fill = red
            if qty_ok:
                ws.cell(value_row, start + 2).fill = green
            if ws.cell(value_row, start + 4).value == "✓":
                ws.cell(value_row, start + 4).fill = green
            needs_review = match.status == "review" or "разные бренды" in (match.reason or "")
            if needs_review:
                ws.cell(desc_row, start).fill = red if "разные бренды" in (match.reason or "") else yellow
                if not qty_ok:
                    ws.cell(value_row, start + 2).fill = yellow
        ws.row_dimensions[desc_row].height = 28
        ws.row_dimensions[label_row].height = 21
        ws.row_dimensions[value_row].height = 25
        row += 3

    for sheet_row in ws.iter_rows():
        for cell in sheet_row:
            cell.border = border
            if cell.alignment is None or cell.alignment == Alignment():
                cell.alignment = Alignment(wrap_text=True, vertical="center")

    widths = {1: 9, 2: 52, 3: 12}
    for col_idx, width in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    for col_idx in range(4, ws.max_column + 1):
        offset = (col_idx - 4) % 5
        width = [9, 13, 12, 14, 16][offset]
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    for price_col in range(5, ws.max_column + 1, 5):
        for price_cell in ws.iter_cols(min_col=price_col, max_col=price_col, min_row=3, max_row=ws.max_row):
            for cell in price_cell:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0.00 ₽'
    for total_col in range(7, ws.max_column + 1, 5):
        for total_cell in ws.iter_cols(min_col=total_col, max_col=total_col, min_row=3, max_row=ws.max_row):
            for cell in total_cell:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0.00 ₽'
    ws.freeze_panes = "D3"

    if unmatched:
        extra = wb.create_sheet("Не сопоставлено")
        headers = ["Поставщик", "Строка", "Позиция КП", "Кол-во", "Ед.", "Цена", "Сумма", "Источник"]
        extra.append(headers)
        for offer in unmatched:
            extra.append([offer.supplier, offer.row_no, offer.name, offer.qty, offer.unit, offer.price, offer.total, offer.source])
        style_sheet(extra)
        extra.column_dimensions["A"].width = 20
        extra.column_dimensions["B"].width = 10
        extra.column_dimensions["C"].width = 70
        extra.column_dimensions["D"].width = 12
        extra.column_dimensions["E"].width = 10
        extra.column_dimensions["F"].width = 14
        extra.column_dimensions["G"].width = 14
        extra.column_dimensions["H"].width = 24
        extra.freeze_panes = "A2"
        for col_letter in ("F", "G"):
            for cell in extra[col_letter][1:]:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0.00 ₽'

    if service_items:
        service = wb.create_sheet("Доставка и услуги")
        headers = ["Поставщик", "Строка", "Позиция КП", "Кол-во", "Ед.", "Цена", "Сумма", "Источник"]
        service.append(headers)
        for offer in service_items:
            service.append([offer.supplier, offer.row_no, offer.name, offer.qty, offer.unit, offer.price, offer.total, offer.source])
        style_sheet(service)
        service.column_dimensions["A"].width = 20
        service.column_dimensions["B"].width = 10
        service.column_dimensions["C"].width = 70
        service.column_dimensions["D"].width = 12
        service.column_dimensions["E"].width = 10
        service.column_dimensions["F"].width = 14
        service.column_dimensions["G"].width = 14
        service.column_dimensions["H"].width = 24
        service.freeze_panes = "A2"
        for col_letter in ("F", "G"):
            for cell in service[col_letter][1:]:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0.00 ₽'

    wb.save(path)


def dump_json(path: Path, request_items: list[RequestItem], supplier_items: list[SupplierItem], matches: list[Match]) -> None:
    data = {
        "request_items": [asdict(item) for item in request_items],
        "supplier_items": [asdict(item) for item in supplier_items],
        "matches": [
            {
                **asdict(match),
                "supplier_item": asdict(match.supplier_item),
            }
            for match in matches
        ],
    }
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Сравнение заявки/ТЗ с КП поставщиков и выпуск Excel-сводки.")
    parser.add_argument("--request", required=True, type=Path, help="Файл заявки .xlsx")
    parser.add_argument("--offers", required=True, nargs="+", type=Path, help="Файлы КП: .xlsx или текстовые .pdf")
    parser.add_argument("--out", required=True, type=Path, help="Финальный Excel-файл")
    parser.add_argument("--review", type=Path, help="Файл ручной проверки, который нужно создать")
    parser.add_argument("--review-in", type=Path, help="Заполненный файл ручной проверки для применения правок")
    parser.add_argument("--debug-json", type=Path, help="Сохранить извлеченные данные и сопоставления в JSON")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    request_items = read_request_xlsx(args.request)
    supplier_items: list[SupplierItem] = []
    errors = []
    for offer_path in args.offers:
        try:
            supplier_items.extend(read_offer(offer_path))
        except Exception as exc:  # noqa: BLE001 - CLI should report all bad input files together.
            errors.append(str(exc))
    if errors:
        print("Предупреждения по входным файлам:")
        for error in errors:
            print(f"- {error}")
    if not supplier_items:
        raise SystemExit("Не удалось извлечь ни одной позиции КП.")

    matches = build_matches(request_items, supplier_items)
    ai_warnings = get_ai_warnings()
    if ai_warnings:
        print("Предупреждения по ИИ:")
        for warning in ai_warnings:
            print(f"- {warning}")
    if args.review_in:
        matches = apply_overrides(matches, read_review_overrides(args.review_in))
    if args.review:
        args.review.parent.mkdir(parents=True, exist_ok=True)
        write_review(args.review, matches, request_items)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    write_final(args.out, request_items, matches)
    if args.debug_json:
        args.debug_json.parent.mkdir(parents=True, exist_ok=True)
        dump_json(args.debug_json, request_items, supplier_items, matches)

    print(f"Позиции заявки: {len(request_items)}")
    print(f"Позиции КП: {len(supplier_items)}")
    print(f"Сопоставлено автоматически/на проверку: {sum(1 for m in matches if m.request_pos)}")
    print(f"Несопоставлено: {sum(1 for m in matches if not m.request_pos)}")
    if args.review:
        print(f"Файл проверки: {args.review}")
    print(f"Финальный отчет: {args.out}")


if __name__ == "__main__":
    main()
