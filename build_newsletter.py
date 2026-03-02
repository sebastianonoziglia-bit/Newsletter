#!/usr/bin/env python3
"""Build a reusable newsletter HTML file from Excel or Google Sheets."""

from __future__ import annotations

import argparse
import csv
import html
import math
import re
import shutil
import sys
from dataclasses import dataclass
from datetime import date, datetime
from io import StringIO
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlsplit, urlunsplit
from urllib.request import urlopen

from openpyxl import Workbook, load_workbook
try:
    from premailer import transform as premailer_transform
except ImportError:  # pragma: no cover - optional dependency
    premailer_transform = None

MAX_POINTS = 10
NUMBER_PATTERN = re.compile(
    r"(?<!\w)(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?(?:[kKmMbBtT%])?(?!\w)"
)
GOOGLE_SHEET_ID_PATTERN = re.compile(r"/spreadsheets/d/([a-zA-Z0-9-_]+)")

DEFAULT_META = {
    "eyebrow": "Globalite Macro Brief",
    "main_title": "WEEKLY TOP 10 ARGUMENTS",
    "subtitle": "A clear weekly macro summary with the key arguments that matter.",
    "block_height": "925000",
    "max_supply_btc": "21000000",
    "circulating_supply_btc": "19960000",
    "hashrate_eh_s": "820",
    "hashrate_scale_eh_s": "1000",
    "snapshot_title": "At The Time Of Writing",
    "snapshot_intro": (
        "At the time of writing, these on-chain supply anchors provide the baseline context."
    ),
    "snapshot_note": "Figures are rounded and updated with each issue.",
    "tldr_title": "TL;DR",
    "tldr_content": (
        "Leverage reset first, liquidity expanded next, and structural adoption kept building."
    ),
    "conclusion_title": "GLOBALITE CONCLUSION",
    "conclusion_content": (
        "For deeper context on these points, visit globalite.co.\n"
        "Our team tracks macro shifts, liquidity, and positioning every week."
    ),
    "cta_url": "https://globalite.co",
    "cta_label": "globalite.co",
    "address_line": "Globalite, Lugano, Piazza dell'Indipendenza 3, CAP 6901",
    "footer_line": "Globalite Macro Brief - For internal distribution.",
    "hero_image_url": "public/hero.png",
    "footer_logo_url": "public/logotosite.png",
    "footer_x_icon": "public/x:twitter.png",
    "footer_linkedin_icon": "public/linkedin.png",
    "image_dir": ".",
    "auto_image_by_order": "true",
}

GRAPH_SETTINGS_ORDER = [
    "btc_price",
    "liquidations",
    "treasuries",
    "circulating_btc",
    "ownership",
]


DEFAULT_GRAPH_SETTINGS = {
    "btc_price": {
        "show": "yes",
        "title": "BTC Price",
        "comment": "",
        "comment_position": "below",
        "top_n": "60",
        "series_filter": "",
    },
    "liquidations": {
        "show": "yes",
        "title": "Liquidations",
        "comment": "",
        "comment_position": "below",
        "top_n": "6",
        "series_filter": "",
    },
    "treasuries": {
        "show": "yes",
        "title": "Treasuries (Top Holders)",
        "comment": "",
        "comment_position": "below",
        "top_n": "6",
        "series_filter": "",
    },
    "circulating_btc": {
        "show": "yes",
        "title": "Circulating BTC",
        "comment": "",
        "comment_position": "below",
        "top_n": "",
        "series_filter": "",
    },
    "ownership": {
        "show": "yes",
        "title": "Supply Ownership",
        "comment": "",
        "comment_position": "below",
        "top_n": "8",
        "series_filter": "",
    },
}


@dataclass
class Point:
    order: int
    title: str
    content: str
    image_path: str
    image_caption: str
    source: str


@dataclass
class BtcPricePoint:
    date_label: str
    price: float


@dataclass
class TreasuryBar:
    entity: str
    btc: float


@dataclass
class LiquidationBar:
    label: str
    longs: float
    shorts: float
    total: float


@dataclass
class CirculatingMetric:
    as_of_date: str
    circulating_supply_btc: float
    max_supply_btc: float
    note: str


@dataclass
class LiveBtcPrice:
    price: float
    date_label: str
    currency: str


@dataclass
class OwnershipSegment:
    category: str
    amount_btc: float
    percent: float
    color: str
    as_of_date: str


@dataclass
class GraphSetting:
    show: bool
    title: str
    comment: str
    comment_position: str
    top_n: int | None
    series_filter: str


class TabularSheet:
    """Small adapter that exposes CSV rows like an openpyxl worksheet."""

    def __init__(self, rows: list[list[str]]) -> None:
        self.rows = rows

    def iter_rows(
        self,
        min_row: int = 1,
        max_row: int | None = None,
        values_only: bool = False,
    ):
        start = max(min_row - 1, 0)
        stop = max_row if max_row is not None else len(self.rows)
        for row in self.rows[start:stop]:
            # This script only reads values (no cell objects).
            yield tuple(row)


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def parse_bool(value: str) -> bool:
    return normalize_text(value).lower() in {"1", "true", "yes", "y", "on"}


def normalize_asset_key(value: object) -> str:
    return re.sub(r"[^A-Z0-9]", "", normalize_text(value).upper())


def is_btc_asset(value: object) -> bool:
    key = normalize_asset_key(value)
    return key in {"BTC", "BITCOIN", "BTCUSD", "XBT", "XBTUSD"}


def parse_number(value: object, default: float = 0.0) -> float:
    text = normalize_text(value).replace(",", "")
    if not text:
        return default
    try:
        return float(text)
    except ValueError:
        return default


def parse_positive_int(value: object, default: int | None = None) -> int | None:
    text = normalize_text(value)
    if not text:
        return default
    try:
        parsed = int(float(text))
    except ValueError:
        return default
    if parsed <= 0:
        return default
    return parsed


def normalize_graph_key(value: object) -> str:
    key = normalize_text(value).lower().replace("-", "_").replace(" ", "_")
    aliases = {
        "btc": "btc_price",
        "btcprice": "btc_price",
        "btc_price": "btc_price",
        "price": "btc_price",
        "liquidation": "liquidations",
        "liquidations": "liquidations",
        "treasury": "treasuries",
        "treasuries": "treasuries",
        "circulating": "circulating_btc",
        "circulating_btc": "circulating_btc",
        "circulatingbtc": "circulating_btc",
        "ownership": "ownership",
        "distribution": "ownership",
    }
    return aliases.get(key, key)


def default_graph_settings() -> dict[str, GraphSetting]:
    settings: dict[str, GraphSetting] = {}
    for key in GRAPH_SETTINGS_ORDER:
        raw = DEFAULT_GRAPH_SETTINGS[key]
        settings[key] = GraphSetting(
            show=parse_bool(raw.get("show", "yes")),
            title=normalize_text(raw.get("title", "")),
            comment=normalize_text(raw.get("comment", "")),
            comment_position=normalize_text(raw.get("comment_position", "below")).lower() or "below",
            top_n=parse_positive_int(raw.get("top_n", ""), default=None),
            series_filter=normalize_text(raw.get("series_filter", "")),
        )
    return settings


def read_graph_settings(graph_settings_sheet) -> dict[str, GraphSetting]:
    settings = default_graph_settings()
    first_row = list(graph_settings_sheet.iter_rows(min_row=1, max_row=1, values_only=True))
    if not first_row:
        return settings

    headers = first_row[0]
    mapping: dict[str, int] = {}
    for index, header in enumerate(headers):
        key = normalize_text(header).lower()
        if key:
            mapping[key] = index

    if "graph_key" not in mapping:
        return settings

    for row in graph_settings_sheet.iter_rows(min_row=2, values_only=True):
        graph_key = normalize_graph_key(row_value(row, mapping["graph_key"]))
        if graph_key not in settings:
            continue

        current = settings[graph_key]
        show = (
            parse_bool(normalize_text(row_value(row, mapping["show"])))
            if "show" in mapping and normalize_text(row_value(row, mapping["show"]))
            else current.show
        )
        title = (
            normalize_text(row_value(row, mapping["title"]))
            if "title" in mapping and normalize_text(row_value(row, mapping["title"]))
            else current.title
        )
        comment = (
            normalize_text(row_value(row, mapping["comment"]))
            if "comment" in mapping
            else current.comment
        )
        comment_position = (
            normalize_text(row_value(row, mapping["comment_position"])).lower()
            if "comment_position" in mapping
            else current.comment_position
        )
        if comment_position not in {"above", "below"}:
            comment_position = "below"
        top_n = (
            parse_positive_int(row_value(row, mapping["top_n"]), default=current.top_n)
            if "top_n" in mapping
            else current.top_n
        )
        series_filter = (
            normalize_text(row_value(row, mapping["series_filter"]))
            if "series_filter" in mapping and normalize_text(row_value(row, mapping["series_filter"]))
            else current.series_filter
        )

        settings[graph_key] = GraphSetting(
            show=show,
            title=title,
            comment=comment,
            comment_position=comment_position,
            top_n=top_n,
            series_filter=series_filter,
        )

    return settings


def get_graph_setting(
    settings: dict[str, GraphSetting], graph_key: str
) -> GraphSetting:
    normalized = normalize_graph_key(graph_key)
    if normalized in settings:
        return settings[normalized]
    return default_graph_settings()[normalized]


def row_enabled(
    row: tuple[object, ...], mapping: dict[str, int]
) -> bool:
    for key in ("show", "include"):
        if key in mapping:
            raw = normalize_text(row_value(row, mapping[key]))
            if raw:
                return parse_bool(raw)
    return True


def row_matches_series_filter(
    row: tuple[object, ...],
    mapping: dict[str, int],
    setting: GraphSetting | None,
    candidate_columns: list[str],
) -> bool:
    if setting is None:
        return True
    filter_text = normalize_text(setting.series_filter).lower()
    if not filter_text:
        return True

    allowed = {item.strip().lower() for item in filter_text.split(",") if item.strip()}
    if not allowed:
        return True

    for column in candidate_columns:
        if column not in mapping:
            continue
        raw = normalize_text(row_value(row, mapping[column])).lower()
        if raw in allowed:
            return True
    return False


def parse_order(value: object, row_number: int) -> int:
    text = normalize_text(value)
    if not text:
        raise ValueError(f"Missing order value at points row {row_number}.")
    if isinstance(value, int):
        parsed = value
    elif isinstance(value, float):
        if not value.is_integer():
            raise ValueError(
                f"Order value must be a whole number at points row {row_number}."
            )
        parsed = int(value)
    else:
        if not text.isdigit():
            raise ValueError(
                f"Invalid order value '{text}' at points row {row_number}."
            )
        parsed = int(text)

    if parsed < 1 or parsed > MAX_POINTS:
        raise ValueError(
            f"Order value must be between 1 and {MAX_POINTS} at points row {row_number}."
        )
    return parsed


def extract_google_sheet_id(sheet_ref: str) -> str:
    value = normalize_text(sheet_ref)
    if not value:
        raise ValueError("Google Sheet URL or ID is required.")

    match = GOOGLE_SHEET_ID_PATTERN.search(value)
    if match:
        return match.group(1)

    if re.fullmatch(r"[a-zA-Z0-9-_]{20,}", value):
        return value

    raise ValueError(
        "Invalid --google-sheet value. Pass a full Google Sheet URL or a sheet ID."
    )


def normalize_table_rows(rows: list[list[str]]) -> list[list[str]]:
    if not rows:
        return []
    width = max(len(row) for row in rows)
    return [row + [""] * (width - len(row)) for row in rows]


def next_history_workbook_path(history_dir: Path) -> Path:
    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M")
    backup_path = history_dir / f"newsletter_{timestamp}.xlsx"
    if backup_path.exists():
        timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        backup_path = history_dir / f"newsletter_{timestamp}.xlsx"
    return backup_path


def write_google_snapshot_workbook(
    path: Path,
    meta_rows: list[list[str]],
    points_rows: list[list[str]],
    live_prices_rows: list[list[str]],
    btc_price_rows: list[list[str]],
    treasuries_rows: list[list[str]],
    circulating_rows: list[list[str]],
    liquidations_rows: list[list[str]],
    ownership_rows: list[list[str]],
    graph_settings_rows: list[list[str]],
) -> None:
    workbook = Workbook()
    meta_sheet = workbook.active
    meta_sheet.title = "meta"
    for row in meta_rows:
        meta_sheet.append(row)

    points_sheet = workbook.create_sheet("points")
    for row in points_rows:
        points_sheet.append(row)

    optional_sheets = [
        ("live_prices", live_prices_rows),
        ("BTC Price", btc_price_rows),
        ("Treasuries", treasuries_rows),
        ("Circulating BTC", circulating_rows),
        ("Liquidations", liquidations_rows),
        ("Distribution", ownership_rows),
        ("graph_settings", graph_settings_rows),
    ]
    for sheet_name, rows in optional_sheets:
        if not rows:
            continue
        sheet = workbook.create_sheet(sheet_name)
        for row in rows:
            sheet.append(row)

    workbook.save(path)


def fetch_google_sheet_rows(
    sheet_id: str,
    tab_name: str,
    required: bool = True,
) -> list[list[str]]:
    safe_tab = quote(normalize_text(tab_name), safe="")
    url = (
        f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?"
        f"tqx=out:csv&sheet={safe_tab}"
    )
    try:
        with urlopen(url, timeout=30) as response:
            payload = response.read().decode("utf-8-sig")
    except HTTPError as error:
        if not required and error.code in {400, 404}:
            return []
        raise RuntimeError(
            f"Could not load Google Sheet tab '{tab_name}' (HTTP {error.code}). "
            "Confirm sharing is enabled and tab names are correct."
        ) from error
    except URLError as error:
        raise RuntimeError(
            f"Network error while loading Google Sheet tab '{tab_name}': {error.reason}"
        ) from error

    text = payload.strip()
    if not text:
        if required:
            raise ValueError(f"Google Sheet tab '{tab_name}' is empty.")
        return []

    lowered = text.lower()
    if lowered.startswith("<!doctype html") or lowered.startswith("<html"):
        if not required:
            return []
        raise RuntimeError(
            f"Could not read Google Sheet tab '{tab_name}'. "
            "Share the sheet as viewable (at least 'Anyone with the link can view')."
        )
    if "google.visualization.query.setresponse" in lowered and "status\":\"error\"" in lowered:
        if not required:
            return []
        raise RuntimeError(
            f"Google Sheets returned an error for tab '{tab_name}'. "
            "Check that the tab exists and has access permissions."
        )

    rows = normalize_table_rows(list(csv.reader(StringIO(text))))
    if not rows:
        if required:
            raise ValueError(f"Google Sheet tab '{tab_name}' is empty.")
        return []
    return rows


def create_template_workbook(path: Path, force: bool = False) -> None:
    if path.exists() and not force:
        raise FileExistsError(f"Template already exists at: {path}")

    wb = Workbook()
    meta = wb.active
    meta.title = "meta"
    meta.append(["key", "value"])
    for key, value in DEFAULT_META.items():
        meta.append([key, value])

    points = wb.create_sheet("points")
    points.append(
        ["order", "title", "content", "image_path", "image_caption", "source"]
    )
    points.append(
        [
            1,
            "Liquidity stopped tightening",
            "QT pace has slowed materially.\n- Funding stress eased\n- Repo usage normalized",
            "",
            "Optional image caption",
            "Source: Example Research Desk",
        ]
    )
    points.append(
        [
            2,
            "Market leverage reset",
            "A broad deleveraging event removed excess risk without structural damage.",
            "",
            "",
            "",
        ]
    )

    graph_settings = wb.create_sheet("graph_settings")
    graph_settings.append(
        [
            "graph_key",
            "show",
            "title",
            "comment",
            "comment_position",
            "top_n",
            "series_filter",
        ]
    )
    for key in GRAPH_SETTINGS_ORDER:
        cfg = DEFAULT_GRAPH_SETTINGS[key]
        graph_settings.append(
            [
                key,
                cfg["show"],
                cfg["title"],
                cfg["comment"],
                cfg["comment_position"],
                cfg["top_n"],
                cfg["series_filter"],
            ]
        )

    live_prices = wb.create_sheet("live_prices")
    live_prices.append(
        [
            "date",
            "price",
            "open",
            "high",
            "low",
            "vol.",
            "change %",
            "market cap.",
            "currency",
            "asset",
            "outstanding shares",
            "tag",
        ]
    )
    live_prices.append(
        [
            "2026-02-19",
            96000,
            95500,
            96800,
            94800,
            1250000,
            0.5,
            1900000000000,
            "USD",
            "Bitcoin",
            "",
            "crypto",
        ]
    )

    btc_price = wb.create_sheet("BTC Price")
    btc_price.append(["date", "price", "asset", "show"])
    btc_price.append(["2026-02-18", 95800, "BTC", "yes"])
    btc_price.append(["2026-02-19", 96400, "BTC", "yes"])

    treasuries = wb.create_sheet("Treasuries")
    treasuries.append(["entity", "btc", "holder_group", "show"])
    treasuries.append(
        [
            "Strategy (MicroStrategy)",
            717131,
            "Public Companies",
            "yes",
        ]
    )
    treasuries.append(
        [
            "Marathon Digital Holdings Inc",
            52850,
            "Public Companies",
            "yes",
        ]
    )

    circulating = wb.create_sheet("Circulating BTC")
    circulating.append(["as_of_date", "circulating_supply_btc", "max_supply_btc", "note", "show"])
    circulating.append(
        [
            "2026-02-19",
            19960000,
            21000000,
            "Update circulating_supply_btc over time; keep max_supply_btc at 21,000,000 unless needed.",
            "yes",
        ]
    )

    liquidations = wb.create_sheet("Liquidations")
    liquidations.append(["label", "longs", "shorts", "total", "period_type", "show"])
    liquidations.append(
        [
            "Dec '25",
            0.62,
            0.38,
            1.00,
            "monthly",
            "yes",
        ]
    )
    liquidations.append(
        [
            "Jan '26",
            3.90,
            0.70,
            4.60,
            "monthly",
            "yes",
        ]
    )

    distribution = wb.create_sheet("Distribution")
    distribution.append(["category", "amount_btc", "percent", "color", "date", "show"])
    distribution.append(
        [
            "Individuals",
            13660000,
            65.1,
            "rgb(255, 66, 2)",
            "2025-12-25",
            "yes",
        ]
    )
    distribution.append(
        [
            "Funds & ETFs",
            1490000,
            7.1,
            "rgb(255, 140, 90)",
            "2025-12-25",
            "yes",
        ]
    )

    wb.save(path)


def read_meta(meta_sheet) -> dict[str, str]:
    meta: dict[str, str] = {}
    for row in meta_sheet.iter_rows(min_row=2, values_only=True):
        key = normalize_text(row[0] if len(row) > 0 else "")
        value = normalize_text(row[1] if len(row) > 1 else "")
        if key:
            meta[key] = value
    for key, value in DEFAULT_META.items():
        meta.setdefault(key, value)
    return meta


def header_index_map(points_sheet) -> dict[str, int]:
    headers = list(points_sheet.iter_rows(min_row=1, max_row=1, values_only=True))[0]
    mapping: dict[str, int] = {}
    for index, header in enumerate(headers):
        key = normalize_text(header).lower()
        if key:
            mapping[key] = index
    required = ["order", "title", "content", "image_path", "image_caption"]
    missing = [name for name in required if name not in mapping]
    if missing:
        raise ValueError(
            "Missing required columns in points sheet: " + ", ".join(missing)
        )
    return mapping


def read_points(points_sheet) -> list[Point]:
    mapping = header_index_map(points_sheet)
    points: list[Point] = []

    for row_number, row in enumerate(
        points_sheet.iter_rows(min_row=2, values_only=True), start=2
    ):
        order_text = normalize_text(row[mapping["order"]])
        title = normalize_text(row[mapping["title"]])
        content = normalize_text(row[mapping["content"]])
        image_path = normalize_text(row[mapping["image_path"]])
        image_caption = normalize_text(row[mapping["image_caption"]])
        source = (
            normalize_text(row[mapping["source"]])
            if "source" in mapping
            else ""
        )

        if not any([order_text, title, content, image_path, image_caption, source]):
            continue

        order = parse_order(order_text, row_number)
        if not title:
            raise ValueError(f"Missing title at points row {row_number}.")
        if not content:
            raise ValueError(f"Missing content at points row {row_number}.")

        points.append(
            Point(
                order=order,
                title=title,
                content=content,
                image_path=image_path,
                image_caption=image_caption,
                source=source,
            )
        )

    if not points:
        raise ValueError("No points found. Add at least 1 row in the points sheet.")

    points.sort(key=lambda item: item.order)
    orders = [item.order for item in points]
    duplicates = sorted({order for order in orders if orders.count(order) > 1})
    if duplicates:
        raise ValueError(
            "Duplicate order values found: " + ", ".join(str(value) for value in duplicates)
        )
    if len(points) > MAX_POINTS:
        raise ValueError(
            f"Found {len(points)} points. Max allowed is {MAX_POINTS}. "
            "Delete or merge entries to keep it to 10 or fewer."
        )

    return points


def header_mapping(sheet, required: list[str], sheet_name: str) -> dict[str, int]:
    first_row = list(sheet.iter_rows(min_row=1, max_row=1, values_only=True))
    if not first_row:
        raise ValueError(f"{sheet_name} sheet is empty.")

    mapping: dict[str, int] = {}
    for index, header in enumerate(first_row[0]):
        key = normalize_text(header).lower()
        if key:
            mapping[key] = index

    missing = [name for name in required if name not in mapping]
    if missing:
        raise ValueError(
            f"Missing required columns in {sheet_name} sheet: " + ", ".join(missing)
        )
    return mapping


def row_value(row: tuple[object, ...], index: int) -> object:
    if index < 0 or index >= len(row):
        return ""
    return row[index]


def parse_date_value(value: object) -> datetime | None:
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())

    text = normalize_text(value)
    if not text:
        return None

    if text.endswith("Z"):
        text = text[:-1]

    for fmt in (
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%m/%d/%Y",
        "%d/%m/%Y",
        "%m/%d/%y",
        "%Y-%m-%d %H:%M:%S",
    ):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue

    try:
        return datetime.fromisoformat(text)
    except ValueError:
        return None


def render_date_label(value: object) -> str:
    parsed = parse_date_value(value)
    if parsed is not None:
        return parsed.strftime("%Y-%m-%d")
    return normalize_text(value)


def format_as_of_date(value: object) -> str:
    parsed = parse_date_value(value)
    if parsed is not None:
        return parsed.strftime("%d %b %Y")
    return normalize_text(value)


def clean_entity_label(value: str, max_chars: int = 18) -> str:
    label = re.sub(r"\s*\(.*?\)", "", normalize_text(value))
    if len(label) <= max_chars:
        return label
    return label[: max_chars - 1].rstrip() + "…"


def read_live_btc_price(live_prices_sheet) -> LiveBtcPrice | None:
    mapping = header_mapping(live_prices_sheet, ["date", "asset"], "live_prices")
    close_index = mapping.get("close", mapping.get("price", -1))
    if close_index < 0:
        return None
    currency_index = mapping.get("currency", -1)

    latest: LiveBtcPrice | None = None
    latest_date: datetime | None = None

    for row in live_prices_sheet.iter_rows(min_row=2, values_only=True):
        if not is_btc_asset(row_value(row, mapping["asset"])):
            continue
        price = parse_number(row_value(row, close_index), default=-1)
        if price <= 0:
            continue

        date_raw = row_value(row, mapping["date"])
        date_label = render_date_label(date_raw)
        date_value = parse_date_value(date_raw)
        candidate = LiveBtcPrice(
            price=price,
            date_label=date_label,
            currency=normalize_text(row_value(row, currency_index)) or "USD",
        )

        if latest is None:
            latest = candidate
            latest_date = date_value
            continue

        if date_value and (latest_date is None or date_value >= latest_date):
            latest = candidate
            latest_date = date_value
        elif date_value is None and latest_date is None:
            latest = candidate

    return latest


def read_btc_price_points(
    price_sheet, limit: int = 60, setting: GraphSetting | None = None
) -> list[BtcPricePoint]:
    mapping = header_mapping(price_sheet, ["date", "price"], "BTC Price")
    entries: list[tuple[datetime | None, int, BtcPricePoint]] = []

    for index, row in enumerate(price_sheet.iter_rows(min_row=2, values_only=True)):
        if not row_enabled(row, mapping):
            continue
        if not row_matches_series_filter(row, mapping, setting, ["series", "asset"]):
            continue
        price = parse_number(row_value(row, mapping["price"]), default=-1)
        if price <= 0:
            continue
        date_raw = row_value(row, mapping["date"])
        date_value = parse_date_value(date_raw)
        entries.append(
            (
                date_value,
                index,
                BtcPricePoint(date_label=render_date_label(date_raw), price=price),
            )
        )

    if not entries:
        return []

    if any(item[0] is not None for item in entries):
        entries.sort(
            key=lambda item: (
                item[0] is None,
                item[0] if item[0] is not None else datetime.max,
                item[1],
            )
        )

    by_date_label: dict[str, BtcPricePoint] = {}
    for _, _, point in entries:
        by_date_label[point.date_label] = point
    points = list(by_date_label.values())
    if len(points) > limit:
        points = points[-limit:]
    return points


def read_btc_price_points_from_live_prices(
    live_prices_sheet, limit: int = 60
) -> list[BtcPricePoint]:
    mapping = header_mapping(live_prices_sheet, ["date", "asset"], "live_prices")
    price_index = mapping.get("price", mapping.get("close", -1))
    if price_index < 0:
        return []

    entries: list[tuple[datetime | None, int, BtcPricePoint]] = []
    for index, row in enumerate(live_prices_sheet.iter_rows(min_row=2, values_only=True)):
        if not is_btc_asset(row_value(row, mapping["asset"])):
            continue
        price = parse_number(row_value(row, price_index), default=-1)
        if price <= 0:
            continue
        date_raw = row_value(row, mapping["date"])
        entries.append(
            (
                parse_date_value(date_raw),
                index,
                BtcPricePoint(date_label=render_date_label(date_raw), price=price),
            )
        )

    if not entries:
        return []

    if any(item[0] is not None for item in entries):
        entries.sort(
            key=lambda item: (
                item[0] is None,
                item[0] if item[0] is not None else datetime.max,
                item[1],
            )
        )

    by_date_label: dict[str, BtcPricePoint] = {}
    for _, _, point in entries:
        by_date_label[point.date_label] = point
    points = list(by_date_label.values())
    if len(points) > limit:
        points = points[-limit:]
    return points


def read_treasury_bars(
    treasuries_sheet, limit: int = 6, setting: GraphSetting | None = None
) -> list[TreasuryBar]:
    mapping = header_mapping(treasuries_sheet, ["entity", "btc"], "Treasuries")
    row_type_index = mapping.get("row_type", -1)
    bars: list[TreasuryBar] = []

    for row in treasuries_sheet.iter_rows(min_row=2, values_only=True):
        if not row_enabled(row, mapping):
            continue
        if not row_matches_series_filter(
            row, mapping, setting, ["series", "holder_group", "group"]
        ):
            continue
        row_type = normalize_text(row_value(row, row_type_index)).lower()
        if row_type and row_type != "entity":
            continue

        entity = normalize_text(row_value(row, mapping["entity"]))
        btc = parse_number(row_value(row, mapping["btc"]), default=0)
        if not entity or btc <= 0:
            continue

        bars.append(TreasuryBar(entity=entity, btc=btc))

    bars.sort(key=lambda item: item.btc, reverse=True)
    return bars[:limit]


def read_circulating_metric(
    circulating_sheet, setting: GraphSetting | None = None
) -> CirculatingMetric | None:
    mapping = header_mapping(
        circulating_sheet,
        ["circulating_supply_btc", "max_supply_btc"],
        "Circulating BTC",
    )
    as_of_index = mapping.get("as_of_date", -1)
    note_index = mapping.get("note", -1)

    candidates: list[tuple[datetime | None, int, CirculatingMetric]] = []
    for index, row in enumerate(circulating_sheet.iter_rows(min_row=2, values_only=True)):
        if not row_enabled(row, mapping):
            continue
        if not row_matches_series_filter(row, mapping, setting, ["series"]):
            continue
        circulating = parse_number(row_value(row, mapping["circulating_supply_btc"]), default=0)
        max_supply = parse_number(row_value(row, mapping["max_supply_btc"]), default=21_000_000)
        if circulating <= 0 and max_supply <= 0:
            continue
        if max_supply <= 0:
            max_supply = 21_000_000

        as_of_raw = row_value(row, as_of_index)
        candidates.append(
            (
                parse_date_value(as_of_raw),
                index,
                CirculatingMetric(
                    as_of_date=render_date_label(as_of_raw),
                    circulating_supply_btc=circulating,
                    max_supply_btc=max_supply,
                    note=normalize_text(row_value(row, note_index)),
                ),
            )
        )

    if not candidates:
        return None

    if any(item[0] is not None for item in candidates):
        candidates.sort(
            key=lambda item: (
                item[0] is None,
                item[0] if item[0] is not None else datetime.max,
                item[1],
            )
        )
    return candidates[-1][2]


def read_liquidation_bars(
    liquidations_sheet, limit: int = 6, setting: GraphSetting | None = None
) -> list[LiquidationBar]:
    mapping = header_mapping(liquidations_sheet, ["label", "longs", "shorts"], "Liquidations")
    period_type_index = mapping.get("period_type", -1)
    total_index = mapping.get("total", -1)
    period_key_index = mapping.get("period_key", -1)

    def collect(monthly_only: bool) -> list[LiquidationBar]:
        bars: list[LiquidationBar] = []
        for row in liquidations_sheet.iter_rows(min_row=2, values_only=True):
            if not row_enabled(row, mapping):
                continue
            if not row_matches_series_filter(
                row, mapping, setting, ["series", "period_type", "asset"]
            ):
                continue
            period_type = normalize_text(row_value(row, period_type_index)).lower()
            if monthly_only and period_type and period_type != "monthly":
                continue

            label = normalize_text(row_value(row, mapping["label"])) or normalize_text(
                row_value(row, period_key_index)
            )
            longs = max(0.0, parse_number(row_value(row, mapping["longs"]), default=0))
            shorts = max(0.0, parse_number(row_value(row, mapping["shorts"]), default=0))
            total = parse_number(row_value(row, total_index), default=longs + shorts)
            if total <= 0:
                total = longs + shorts

            if not label or total <= 0:
                continue

            bars.append(LiquidationBar(label=label, longs=longs, shorts=shorts, total=total))
        return bars

    bars = collect(monthly_only=True)
    if not bars:
        bars = collect(monthly_only=False)

    if len(bars) > limit:
        bars = bars[-limit:]
    return bars


def read_ownership_segments(
    distribution_sheet, setting: GraphSetting | None = None
) -> list[OwnershipSegment]:
    mapping = header_mapping(
        distribution_sheet,
        ["category", "amount_btc"],
        "Distribution",
    )
    as_of_index = mapping.get("as_of_date", mapping.get("date", mapping.get("as_of", -1)))
    segments: list[OwnershipSegment] = []
    for row in distribution_sheet.iter_rows(min_row=2, values_only=True):
        if not row_enabled(row, mapping):
            continue
        if not row_matches_series_filter(
            row, mapping, setting, ["series", "tier", "group"]
        ):
            continue
        category = normalize_text(row_value(row, mapping["category"]))
        amount_btc = parse_number(row_value(row, mapping["amount_btc"]), default=0.0)
        color = (
            normalize_text(row_value(row, mapping["color"]))
            if "color" in mapping
            else "rgb(255, 66, 2)"
        ) or "rgb(255, 66, 2)"
        percent = (
            parse_number(row_value(row, mapping["percent"]), default=0.0)
            if "percent" in mapping
            else 0.0
        )
        if not category and amount_btc <= 0:
            continue
        if not category:
            raise ValueError("Distribution row is missing category.")
        if amount_btc < 0:
            raise ValueError(f"Distribution amount cannot be negative for category '{category}'.")
        segments.append(
            OwnershipSegment(
                category=category,
                amount_btc=amount_btc,
                percent=percent,
                color=color,
                as_of_date=normalize_text(row_value(row, as_of_index))
                if as_of_index >= 0
                else "",
            )
        )

    if not segments:
        return []

    total_percent = sum(max(0.0, segment.percent) for segment in segments)
    if total_percent <= 0:
        total_amount = sum(max(0.0, segment.amount_btc) for segment in segments)
        if total_amount > 0:
            for segment in segments:
                segment.percent = (max(0.0, segment.amount_btc) / total_amount) * 100.0
    else:
        for segment in segments:
            segment.percent = (max(0.0, segment.percent) / total_percent) * 100.0

    segments.sort(key=lambda item: item.amount_btc, reverse=True)
    return segments


def render_content_blocks(raw: str) -> str:
    lines = raw.replace("\r\n", "\n").split("\n")
    blocks: list[str] = []
    list_open = False

    def close_list() -> None:
        nonlocal list_open
        if list_open:
            blocks.append("</ul>")
            list_open = False

    for original_line in lines:
        line = original_line.strip()
        if not line:
            close_list()
            continue

        if line.startswith("- ") or line.startswith("* "):
            item = emphasize_lead_label_and_numbers(line[2:].strip())
            if not list_open:
                blocks.append("<ul>")
                list_open = True
            blocks.append(f"<li>{item}</li>")
        else:
            close_list()
            blocks.append(f"<p>{emphasize_lead_label_and_numbers(line)}</p>")

    close_list()
    return "\n".join(blocks)


def emphasize_lead_label_and_numbers(text: str) -> str:
    normalized = normalize_text(text)
    if not normalized:
        return ""
    if re.match(r"^https?://", normalized, flags=re.IGNORECASE):
        return emphasize_numbers(normalized)
    label_match = re.match(r"^([^:\n]{1,120}:)(\s*.*)?$", normalized)
    if not label_match:
        return emphasize_numbers(normalized)
    lead = f"<strong>{html.escape(label_match.group(1))}</strong>"
    rest = normalize_text(label_match.group(2) or "")
    if not rest:
        return lead
    return f"{lead} {emphasize_numbers(rest)}"


def emphasize_numbers(text: str) -> str:
    output: list[str] = []
    start = 0
    for match in NUMBER_PATTERN.finditer(text):
        output.append(html.escape(text[start : match.start()]))
        output.append(f"<strong>{html.escape(match.group(0))}</strong>")
        start = match.end()
    output.append(html.escape(text[start:]))
    return "".join(output)


def format_btc_integer(value: float) -> str:
    return f"{int(round(value)):,}"


def format_btc_compact(value: float) -> str:
    abs_value = abs(value)
    if abs_value >= 1_000_000:
        rendered = f"{value / 1_000_000:.2f}".rstrip("0").rstrip(".")
        return f"{rendered}M BTC"
    if abs_value >= 1_000:
        rendered = f"{int(round(value / 1_000)):,}"
        return f"{rendered}K BTC"
    return f"{format_btc_integer(value)} BTC"


def format_percent(value: float) -> str:
    return f"{value:.1f}".rstrip("0").rstrip(".") + "%"


def render_block_height(value: str) -> str:
    clean = normalize_text(value)
    if not clean:
        return "n/a"
    numeric = parse_number(clean, default=-1)
    if numeric >= 0:
        return f"{int(round(numeric)):,}"
    return clean


def format_usd(value: float) -> str:
    abs_value = abs(value)
    if abs_value >= 1_000:
        return f"${int(round(value)):,}"
    return f"${value:,.2f}".rstrip("0").rstrip(".")


def build_btc_price_chart_svg(points: list[BtcPricePoint]) -> str:
    if not points:
        return '<p class="market-empty">No BTC price points found.</p>'

    width = 620
    height = 230
    pad_left = 54
    pad_right = 14
    pad_top = 14
    pad_bottom = 34
    plot_width = width - pad_left - pad_right
    plot_height = height - pad_top - pad_bottom
    plot_bottom = pad_top + plot_height

    min_price = min(point.price for point in points)
    max_price = max(point.price for point in points)
    if max_price <= min_price:
        max_price = min_price + 1

    padding = max((max_price - min_price) * 0.08, max_price * 0.01)
    min_price = max(0.0, min_price - padding)
    max_price = max_price + padding

    span = max_price - min_price
    denominator = max(1, len(points) - 1)

    coords: list[tuple[float, float]] = []
    for index, point in enumerate(points):
        x = pad_left + (plot_width * index / denominator)
        y = pad_top + ((max_price - point.price) / span) * plot_height
        coords.append((x, y))

    line_path = "M " + " L ".join(f"{x:.2f} {y:.2f}" for x, y in coords)
    area_path = (
        f"M {coords[0][0]:.2f} {plot_bottom:.2f} "
        + " ".join(f"L {x:.2f} {y:.2f}" for x, y in coords)
        + f" L {coords[-1][0]:.2f} {plot_bottom:.2f} Z"
    )

    grid_lines: list[str] = []
    y_labels: list[str] = []
    for step in range(5):
        y = pad_top + (plot_height * step / 4)
        value = max_price - ((max_price - min_price) * step / 4)
        grid_lines.append(
            f'<line x1="{pad_left:.2f}" y1="{y:.2f}" x2="{(width - pad_right):.2f}" y2="{y:.2f}" stroke="#262626" stroke-width="1" />'
        )
        y_labels.append(
            f'<text x="{(pad_left - 8):.2f}" y="{(y + 4):.2f}" text-anchor="end" fill="#8b8b8b" font-size="10">{html.escape(format_usd(value))}</text>'
        )

    first_label = html.escape(points[0].date_label)
    mid_label = html.escape(points[len(points) // 2].date_label)
    last_label = html.escape(points[-1].date_label)
    x_labels = (
        f'<text x="{pad_left:.2f}" y="{(height - 8):.2f}" text-anchor="start" fill="#8b8b8b" font-size="10">{first_label}</text>'
        f'<text x="{(pad_left + plot_width / 2):.2f}" y="{(height - 8):.2f}" text-anchor="middle" fill="#8b8b8b" font-size="10">{mid_label}</text>'
        f'<text x="{(width - pad_right):.2f}" y="{(height - 8):.2f}" text-anchor="end" fill="#8b8b8b" font-size="10">{last_label}</text>'
    )

    last_x, last_y = coords[-1]
    return f"""
<svg class="market-price-svg" viewBox="0 0 {width} {height}" role="img" aria-label="Bitcoin price trend chart">
  <defs>
    <linearGradient id="priceAreaGradient" x1="0" x2="0" y1="0" y2="1">
      <stop offset="0%" stop-color="#ff4202" stop-opacity="0.42" />
      <stop offset="100%" stop-color="#ff4202" stop-opacity="0" />
    </linearGradient>
  </defs>
  {''.join(grid_lines)}
  {''.join(y_labels)}
  <path d="{area_path}" fill="url(#priceAreaGradient)" />
  <path d="{line_path}" fill="none" stroke="#ff4202" stroke-width="2.5" stroke-linejoin="round" stroke-linecap="round" />
  <circle cx="{last_x:.2f}" cy="{last_y:.2f}" r="4" fill="#ff4202" stroke="#ffffff" stroke-width="1" />
  {x_labels}
</svg>
"""


def render_treasury_bars(treasury_bars: list[TreasuryBar]) -> str:
    if not treasury_bars:
        return '<p class="market-empty">No treasury rows found.</p>'

    max_btc = max((bar.btc for bar in treasury_bars), default=0)
    if max_btc <= 0:
        max_btc = 1

    bars_html: list[str] = []
    for bar in treasury_bars:
        height = max(8.0, (bar.btc / max_btc) * 100.0)
        bars_html.append(
            '<div class="market-bar-item">'
            f'<div class="market-bar-track"><div class="market-bar-fill" style="height:{height:.6f}%;"></div></div>'
            f'<p class="market-bar-label">{html.escape(clean_entity_label(bar.entity))}</p>'
            f'<p class="market-bar-value">{html.escape(format_btc_compact(bar.btc))}</p>'
            "</div>"
        )

    return '<div class="market-bars">' + "".join(bars_html) + "</div>"


def render_liquidations_bars(liquidation_bars: list[LiquidationBar]) -> str:
    if not liquidation_bars:
        return '<p class="market-empty">No liquidation rows found.</p>'

    max_total = max((bar.total for bar in liquidation_bars), default=0)
    if max_total <= 0:
        max_total = 1

    bars_html: list[str] = []
    for bar in liquidation_bars:
        long_height = max(4.0, (bar.longs / max_total) * 100.0) if bar.longs > 0 else 0.0
        short_height = max(4.0, (bar.shorts / max_total) * 100.0) if bar.shorts > 0 else 0.0
        bars_html.append(
            '<div class="market-bar-item">'
            '<div class="market-bar-track market-liq-track">'
            f'<div class="market-liq-short" style="height:{short_height:.6f}%;"></div>'
            f'<div class="market-liq-long" style="height:{long_height:.6f}%;"></div>'
            "</div>"
            f'<p class="market-bar-label">{html.escape(bar.label)}</p>'
            f'<p class="market-bar-value">{html.escape(f"${bar.total:.2f}B".rstrip("0").rstrip("."))}</p>'
            "</div>"
        )

    return (
        '<div class="market-legend">'
        '<span><i class="dot dot-long"></i>Longs</span>'
        '<span><i class="dot dot-short"></i>Shorts</span>'
        "</div>"
        + '<div class="market-bars">'
        + "".join(bars_html)
        + "</div>"
    )


def render_circulating_card(circulating: CirculatingMetric | None) -> str:
    if circulating is None:
        return '<p class="market-empty">No circulating supply row found.</p>'

    max_supply = circulating.max_supply_btc if circulating.max_supply_btc > 0 else 21_000_000
    pct = (circulating.circulating_supply_btc / max_supply) * 100 if max_supply > 0 else 0
    pct = max(0.0, min(100.0, pct))
    note_html = (
        f'<p class="market-note">{html.escape(circulating.note)}</p>' if circulating.note else ""
    )
    as_of_html = (
        f'<p class="market-subnote">As of {html.escape(circulating.as_of_date)}</p>'
        if circulating.as_of_date
        else ""
    )

    return f"""
<div class="market-circ-value">{html.escape(format_btc_integer(circulating.circulating_supply_btc))} BTC</div>
<div class="market-progress-track"><div class="market-progress-fill" style="width:{pct:.6f}%;"></div></div>
<p class="market-subnote">{html.escape(format_percent(pct))} of {html.escape(format_btc_integer(max_supply))} BTC max supply</p>
{as_of_html}
{note_html}
"""


def render_ownership_card(
    ownership_segments: list[OwnershipSegment], max_items: int = 8
) -> str:
    if not ownership_segments:
        return '<p class="market-empty">No ownership rows found.</p>'

    total_supply = sum(max(0.0, segment.amount_btc) for segment in ownership_segments)
    as_of_raw = next(
        (
            normalize_text(segment.as_of_date)
            for segment in ownership_segments
            if normalize_text(segment.as_of_date)
        ),
        "",
    )
    as_of_html = (
        f'<p class="market-subnote">Ownership Breakdown (as of {html.escape(format_as_of_date(as_of_raw))})</p>'
        if as_of_raw
        else '<p class="market-subnote">Ownership Breakdown</p>'
    )

    bar_segments = "".join(
        (
            f'<div class="market-own-segment" style="width:{max(0.0, segment.percent):.6f}%;background:{html.escape(segment.color)};" '
            f'title="{html.escape(segment.category)}: {html.escape(format_btc_compact(segment.amount_btc))} ({html.escape(format_percent(segment.percent))})"></div>'
        )
        for segment in ownership_segments
    )

    legend_rows = "".join(
        (
            '<div class="market-own-item">'
            f'<span class="market-own-dot" style="background:{html.escape(segment.color)};"></span>'
            f'<span class="market-own-name">{html.escape(segment.category)}</span>'
            f'<span class="market-own-value">{html.escape(format_percent(segment.percent))}</span>'
            "</div>"
        )
        for segment in ownership_segments[:max_items]
    )

    return (
        f'<div class="market-own-total">{html.escape(format_btc_integer(total_supply))} BTC</div>'
        + as_of_html
        + '<div class="market-own-bar">'
        + bar_segments
        + "</div>"
        + '<div class="market-own-legend">'
        + legend_rows
        + "</div>"
    )


def render_market_card(
    setting: GraphSetting, body_html: str, extra_class: str = ""
) -> str:
    card_class = "market-card" + (f" {extra_class}" if extra_class else "")
    comment_above = ""
    comment_below = ""
    if setting.comment:
        rendered = f'<p class="market-graph-note">{html.escape(setting.comment)}</p>'
        if setting.comment_position == "above":
            comment_above = rendered
        else:
            comment_below = rendered

    return (
        f'<div class="{card_class}">'
        f"<h3>{html.escape(setting.title)}</h3>"
        f"{comment_above}"
        f"{body_html}"
        f"{comment_below}"
        "</div>"
    )


def render_market_section(
    meta: dict[str, str],
    btc_price_points: list[BtcPricePoint],
    treasury_bars: list[TreasuryBar],
    circulating: CirculatingMetric | None,
    liquidation_bars: list[LiquidationBar],
    ownership_segments: list[OwnershipSegment],
    graph_settings: dict[str, GraphSetting],
    live_btc: LiveBtcPrice | None,
) -> str:
    section_title = (
        normalize_text(meta.get("market_section_title", ""))
        or "Bitcoin Market Dashboard"
    )
    section_intro = (
        normalize_text(meta.get("market_section_intro", ""))
        or "Auto-rendered from live_prices/BTC Price, Liquidations, Treasuries, Circulating BTC, and Distribution tabs."
    )

    live_chip = ""
    if live_btc is not None:
        live_chip = (
            f'<p class="market-live">Live BTC: <strong>{html.escape(format_usd(live_btc.price))}</strong>'
            + (
                f' <span>({html.escape(live_btc.date_label)})</span>'
                if live_btc.date_label
                else ""
            )
            + "</p>"
        )

    cards: list[str] = []

    btc_setting = get_graph_setting(graph_settings, "btc_price")
    if btc_setting.show:
        cards.append(
            render_market_card(
                btc_setting,
                build_btc_price_chart_svg(btc_price_points),
                "market-price-card",
            )
        )

    liq_setting = get_graph_setting(graph_settings, "liquidations")
    if liq_setting.show:
        cards.append(render_market_card(liq_setting, render_liquidations_bars(liquidation_bars)))

    treasuries_setting = get_graph_setting(graph_settings, "treasuries")
    if treasuries_setting.show:
        cards.append(render_market_card(treasuries_setting, render_treasury_bars(treasury_bars)))

    circulating_setting = get_graph_setting(graph_settings, "circulating_btc")
    if circulating_setting.show:
        cards.append(
            render_market_card(
                circulating_setting,
                render_circulating_card(circulating),
                "market-circ-card",
            )
        )

    ownership_setting = get_graph_setting(graph_settings, "ownership")
    if ownership_setting.show:
        ownership_limit = ownership_setting.top_n or 8
        cards.append(
            render_market_card(
                ownership_setting,
                render_ownership_card(ownership_segments, max_items=ownership_limit),
                "market-own-card",
            )
        )

    if not cards:
        return ""

    return f"""
            <tr>
              <td class="section market">
                <h2>{html.escape(section_title)}</h2>
                <p class="market-intro">{html.escape(section_intro)}</p>
                {live_chip}
                <div class="market-grid">
                  {"".join(cards)}
                </div>
              </td>
            </tr>
"""


def looks_like_remote_image_source(path: str) -> bool:
    return path.startswith(("http://", "https://", "data:", "cid:"))


def to_mobile_variant_path(path: str) -> str:
    raw = normalize_text(path)
    if not raw:
        return ""
    if re.match(r"^(data:|cid:)", raw, flags=re.IGNORECASE):
        return ""

    if re.match(r"^https?://", raw, flags=re.IGNORECASE):
        try:
            split = urlsplit(raw)
            parts = split.path.split("/")
            filename = parts[-1] if parts else ""
            if not filename or filename.lower().startswith("mobile"):
                return ""
            parts[-1] = f"mobile{filename}"
            return urlunsplit((split.scheme, split.netloc, "/".join(parts), split.query, split.fragment))
        except Exception:
            return ""

    suffix_match = re.search(r"([?#].*)$", raw)
    suffix = suffix_match.group(1) if suffix_match else ""
    base = raw[: -len(suffix)] if suffix else raw
    dir_name, _, file_name = base.rpartition("/")
    if not file_name:
        return ""
    if file_name.lower().startswith("mobile"):
        return ""
    prefixed = f"mobile{file_name}"
    if dir_name:
        return f"{dir_name}/{prefixed}{suffix}"
    return f"{prefixed}{suffix}"


def build_mobile_variant_candidates(paths: list[str]) -> list[str]:
    out: list[str] = []
    seen: set[str] = set()
    for path in paths:
        candidate = normalize_text(to_mobile_variant_path(path))
        if not candidate:
            continue
        key = candidate.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(candidate)
    return out


def build_auto_image_name_candidates(base: str, extensions: list[str]) -> list[str]:
    clean_base = re.sub(r"\s+", "", normalize_text(base))
    if not clean_base:
        return []
    clean_exts = [normalize_text(ext).lstrip(".").lower() for ext in extensions if normalize_text(ext)]
    modes = ["full", "tight", "plain"]
    names: list[str] = []
    for mode in modes:
        stem = clean_base if mode == "plain" else f"{clean_base}_{mode}"
        for ext in clean_exts:
            names.append(f"{stem}.{ext}")
    return names


def detect_image_mode_from_path(path: str) -> str:
    raw = normalize_text(path).split("?")[0].split("#")[0].lower()
    if re.search(r"_full\.[a-z0-9]{2,5}$", raw):
        return "full"
    if re.search(r"_tight\.[a-z0-9]{2,5}$", raw):
        return "tight"
    return "plain"


def resolve_image_path(point: Point, meta: dict[str, str], output_dir: Path) -> str:
    image_path = normalize_text(point.image_path)
    if image_path:
        candidate = image_path
    elif parse_bool(meta.get("auto_image_by_order", "true")):
        image_dir = normalize_text(meta.get("image_dir", ".")) or "."
        filename_candidates = build_auto_image_name_candidates(
            str(point.order),
            ["png"],
        )
        candidate = ""
        for filename in filename_candidates:
            relative_candidate = (Path(image_dir) / filename).as_posix()
            if (output_dir / relative_candidate).exists():
                candidate = relative_candidate
                break
        if not candidate:
            return ""
    else:
        return ""

    if looks_like_remote_image_source(candidate):
        return candidate

    candidate_path = Path(candidate)
    if candidate_path.is_absolute():
        return candidate if candidate_path.exists() else ""

    return candidate if (output_dir / candidate).exists() else ""


def resolve_extra_image_paths(
    point: Point, meta: dict[str, str], output_dir: Path
) -> list[str]:
    if not parse_bool(meta.get("auto_image_by_order", "true")):
        return []

    image_dir = normalize_text(meta.get("image_dir", ".")) or "."
    max_extra_images = int(parse_number(meta.get("max_extra_images", "10"), default=10))
    max_extra_images = max(0, min(20, max_extra_images))
    extensions = ["png"]
    sources: list[str] = []

    for index in range(1, max_extra_images + 1):
        base = f"{point.order}.{index}"
        name_candidates = build_auto_image_name_candidates(base, extensions)
        for candidate_name in name_candidates:
            candidate = (Path(image_dir) / candidate_name).as_posix()
            if (output_dir / candidate).exists():
                sources.append(candidate)
                break
    return sources


def render_image_block(point: Point, image_src: str, source_text: str = "") -> str:
    if not image_src:
        return ""
    caption = point.image_caption or point.title
    mode = detect_image_mode_from_path(image_src)
    mobile_sources = build_mobile_variant_candidates([image_src])
    mobile_attr = (
        f' data-mobile-srcs="{html.escape("|".join(mobile_sources), quote=True)}"'
        if mobile_sources
        else ""
    )
    source_html = (
        f'  <p class="point-source image-source">{html.escape(source_text)}</p>\n'
        if normalize_text(source_text)
        else ""
    )
    return (
        f'<div class="image mode-{mode}">\n'
        f'  <img class="mode-{mode}" src="{html.escape(image_src)}" data-fallbacks=""{mobile_attr} alt="{html.escape(point.title)}" onerror="const list=(this.dataset.fallbacks||\'\').split(\'|\').filter(Boolean);if(list.length){{this.src=list.shift();this.dataset.fallbacks=list.join(\'|\');}}else{{this.closest(\'.image\').style.display=\'none\';}}">\n'
        f'  <div class="caption">{html.escape(caption)}</div>\n'
        f"{source_html}"
        "</div>"
    )


def render_extra_images_block(
    point: Point, image_sources: list[str], source_by_key: dict[str, str] | None = None
) -> str:
    if not image_sources:
        return ""
    source_by_key = source_by_key or {}
    rows: list[str] = []
    for index, src in enumerate(image_sources, start=1):
        mode = detect_image_mode_from_path(src)
        mobile_sources = build_mobile_variant_candidates([src])
        mobile_attr = (
            f' data-mobile-srcs="{html.escape("|".join(mobile_sources), quote=True)}"'
            if mobile_sources
            else ""
        )
        source_text = source_by_key.get(f"{point.order}.{index}", "")
        source_html = (
            f'    <p class="point-source image-source">{html.escape(source_text)}</p>\n'
            if source_text
            else ""
        )
        rows.append(
            f'  <div class="extra-image-item mode-{mode}">\n'
            f'    <img class="mode-{mode}" src="{html.escape(src)}" data-fallbacks=""{mobile_attr} alt="{html.escape(point.title)} - extra {index}" onerror="const list=(this.dataset.fallbacks||\'\').split(\'|\').filter(Boolean);if(list.length){{this.src=list.shift();this.dataset.fallbacks=list.join(\'|\');}}else{{this.closest(\'.extra-image-item\').style.display=\'none\';}}">\n'
            f"{source_html}"
            "  </div>"
        )
    image_tags = "\n".join(rows)
    return '<div class="extra-images">\n' + image_tags + "\n</div>"


def render_point(point: Point, meta: dict[str, str], output_dir: Path) -> str:
    source_data = parse_point_source_data(point.source)
    parts = [
        "            <tr>",
        '              <td class="section">',
        f"                <h2 class=\"point-title\">{point.order}. {html.escape(point.title)}</h2>",
    ]

    image_src = resolve_image_path(point, meta, output_dir)
    main_source = source_data["by_key"].get(str(point.order), "")
    image_block = render_image_block(point, image_src, main_source)
    if image_block:
        parts.append(indent_block(image_block, 16))

    extra_image_sources = resolve_extra_image_paths(point, meta, output_dir)
    extra_images_block = render_extra_images_block(
        point, extra_image_sources, source_data["by_key"]
    )

    parts.append(indent_block(render_content_blocks(point.content), 16))
    if source_data["unscoped_text"]:
        parts.append(
            f'                <p class="point-source">{html.escape(source_data["unscoped_text"])}</p>'
        )
    if extra_images_block:
        parts.append(indent_block(extra_images_block, 16))
    parts.extend(
        [
            "              </td>",
            "            </tr>",
        ]
    )
    return "\n".join(parts) + "\n"


def parse_point_source_data(raw_source: str) -> dict[str, object]:
    by_key: dict[str, str] = {}
    unscoped_lines: list[str] = []
    text = normalize_text(raw_source)
    if not text:
        return {"by_key": by_key, "unscoped_text": ""}

    for line in [normalize_text(item) for item in text.splitlines() if normalize_text(item)]:
        match = re.match(r"^(\d+(?:\.\d+)?)\s*:\s*(.+)$", line)
        if match:
            key = normalize_source_key(match.group(1))
            value = normalize_text(match.group(2))
            if key and value:
                by_key[key] = value
                continue
        unscoped_lines.append(line)

    return {"by_key": by_key, "unscoped_text": " ".join(unscoped_lines)}


def normalize_source_key(raw_key: str) -> str:
    key = normalize_text(raw_key)
    if not key:
        return ""
    if "." in key:
        return ".".join(str(int(part)) for part in key.split("."))
    return str(int(key))


def indent_block(text: str, spaces: int) -> str:
    if not text:
        return ""
    prefix = " " * spaces
    return "\n".join(prefix + line if line else "" for line in text.splitlines())


def render_html(
    meta: dict[str, str],
    points: list[Point],
    btc_price_points: list[BtcPricePoint],
    treasury_bars: list[TreasuryBar],
    circulating: CirculatingMetric | None,
    liquidation_bars: list[LiquidationBar],
    ownership_segments: list[OwnershipSegment],
    graph_settings: dict[str, GraphSetting],
    live_btc: LiveBtcPrice | None,
    output_dir: Path,
) -> str:
    title = html.escape(meta["main_title"])
    subtitle = html.escape(meta["subtitle"])
    eyebrow = html.escape(meta["eyebrow"])
    block_height = html.escape(render_block_height(meta["block_height"]))
    tldr_title = html.escape(meta["tldr_title"])
    tldr_content = indent_block(render_content_blocks(meta["tldr_content"]), 16)
    conclusion_title = html.escape(meta["conclusion_title"])
    conclusion_content = indent_block(render_content_blocks(meta["conclusion_content"]), 16)
    address_line = html.escape(meta["address_line"])
    footer_line = html.escape(meta["footer_line"])
    hero_image_url = html.escape(
        normalize_text(meta.get("hero_image_url", "public/hero.png")),
        quote=True,
    )
    footer_logo_url = html.escape(
        normalize_text(meta.get("footer_logo_url", "public/logotosite.png")),
        quote=True,
    )
    footer_x_icon = html.escape(
        normalize_text(meta.get("footer_x_icon", "public/x:twitter.png")),
        quote=True,
    )
    footer_linkedin_icon = html.escape(
        normalize_text(meta.get("footer_linkedin_icon", "public/linkedin.png")),
        quote=True,
    )

    points_html = "".join(render_point(point, meta, output_dir) for point in points)
    market_html = render_market_section(
        meta,
        btc_price_points,
        treasury_bars,
        circulating,
        liquidation_bars,
        ownership_segments,
        graph_settings,
        live_btc,
    )

    return f"""<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>{title} - Globalite Macro Brief</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;700&display=swap" rel="stylesheet">
    <style>
      body {{
        margin: 0;
        padding: 0;
        background: #f5f5f5;
        font-family: "Poppins", Arial, sans-serif;
        color: #1f1f1f;
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
      }}
      table {{ border-collapse: collapse; }}
      img {{ border: 0; display: block; max-width: 100%; height: auto; }}
      a {{ color: #ff4202; text-decoration: none; }}
      .toolbar {{ width: 100%; max-width: 512px; margin: 0 auto; display: flex; justify-content: flex-end; padding: 12px 0 8px; }}
      .download-pdf-btn {{ border: 1px solid #ff4202; border-radius: 999px; padding: 8px 14px; background: #ffffff; color: #ff4202; font: 600 12px/1 "Poppins", Arial, sans-serif; cursor: pointer; }}
      .download-pdf-btn:hover {{ background: #fff4ef; }}
      .wrapper {{ width: 100%; background: #f5f5f5; padding: 32px 0; }}
      .container {{ width: 512px; max-width: 512px; background: #ffffff; border: 1px solid #e6e6e6; border-radius: 16px; overflow: hidden; }}
      .divider {{ height: 4px; background: #ff4202; line-height: 4px; }}
      .hero {{ position: relative; overflow: hidden; background: #0a0a0a; min-height: 260px; display: flex; flex-direction: column; justify-content: flex-end; }}
      .hero-bg {{ position: absolute; inset: 0; width: 100%; height: 100%; object-fit: cover; opacity: 0.82; }}
      .hero-gradient {{ position: absolute; inset: 0; background: linear-gradient(to bottom, rgba(0,0,0,0.1) 0%, rgba(0,0,0,0.72) 100%); }}
      .hero-content {{ position: relative; z-index: 2; padding: 28px 32px; }}
      .hero-logo img {{ width: 150px; height: auto; display: block; margin: 0 0 14px; }}
      .hero-eyebrow {{ color: #ff4202; font-weight: 700; font-size: 11px; letter-spacing: 1.4px; text-transform: uppercase; margin: 0 0 6px; }}
      .hero-title {{ margin: 0 0 6px; font-size: 28px; font-weight: 700; color: #ffffff; }}
      .hero-subtitle {{ margin: 0 0 14px; color: rgba(255,255,255,0.6); font-size: 14px; }}
      .hero-badge {{ display: inline-block; font-size: 11px; color: #ffcfb8; background: rgba(255,66,2,0.25); border: 1px solid rgba(255,66,2,0.45); border-radius: 999px; padding: 5px 12px; }}
      .section {{ padding: 16px 32px; border-top: 1px solid #f0f0f0; }}
      .section h2 {{ margin: 0 0 20px; font-size: 18px; font-weight: 700; }}
      .section h2.point-title {{ font-size: 28px; line-height: 1.2; color: #ff4202; }}
      .section p {{ margin: 0; font-size: 14px; line-height: 1.6; }}
      .section p + p {{ margin-top: 12px; }}
      .section ul {{ margin: 20px 0 20px 18px; padding: 0; font-size: 14px; line-height: 1.6; }}
      .section li {{ margin-bottom: 8px; }}
      .section .point-source {{ margin-top: 14px; font-size: 11px; line-height: 1.5; color: #8a8a8a; }}
      .section .point-source.image-source {{ margin-top: 6px; font-size: 10px; line-height: 1.4; }}
      .image {{ margin: 20px 0; width: 100%; box-sizing: border-box; }}
      .image.mode-full {{ margin-left: -32px; margin-right: -32px; width: calc(100% + 64px); }}
      .image.mode-tight {{ text-align: center; }}
      .image img {{
        display: block;
        width: 100%;
        max-width: 100%;
        height: auto;
        margin: 0 auto;
        object-fit: contain;
        border-radius: 12px;
        border: 1px solid #e6e6e6;
      }}
      .image img.is-standard {{
        width: 100%;
        height: auto;
      }}
      .image img.is-portrait {{
        width: auto;
        max-width: 72%;
        height: auto;
      }}
      .image img.is-wide {{
        width: 100%;
        height: auto;
        max-height: none;
      }}
      .image img.mode-full {{
        width: 100% !important;
        max-width: none !important;
        max-height: none !important;
        border-radius: 0;
        border-left: 0;
        border-right: 0;
      }}
      .image img.mode-tight {{
        width: 100% !important;
        max-width: 380px !important;
      }}
      .caption {{ font-size: 12px; color: #7a7a7a; margin-top: 6px; }}
      .extra-images {{ margin: 14px 0 24px; display: grid; gap: 10px; }}
      .extra-image-item {{ display: block; }}
      .extra-image-item.mode-full {{ margin-left: -32px; margin-right: -32px; width: calc(100% + 64px); }}
      .extra-image-item.mode-tight {{ text-align: center; }}
      .extra-images img {{
        display: block;
        width: 100%;
        max-width: 100%;
        height: auto;
        margin: 0 auto;
        object-fit: contain;
        border-radius: 12px;
        border: 1px solid #e6e6e6;
      }}
      .extra-images img.is-standard {{
        width: 100%;
        height: auto;
      }}
      .extra-images img.is-portrait {{
        width: auto;
        max-width: 68%;
        height: auto;
      }}
      .extra-images img.is-wide {{
        width: 100%;
        height: auto;
      }}
      .extra-image-item img.mode-full {{
        width: 100% !important;
        max-width: none !important;
        max-height: none !important;
        border-radius: 0;
        border-left: 0;
        border-right: 0;
      }}
      .extra-image-item img.mode-tight {{
        width: 100% !important;
        max-width: 380px !important;
      }}
      .market {{ background: #070707; color: #f4f4f4; border-top: 1px solid #171717; }}
      .market h2 {{ color: #ffffff; margin-bottom: 8px; }}
      .market-intro {{ margin: 0 0 12px; color: #b8b8b8; font-size: 13px; }}
      .market-live {{ margin: 0 0 12px; display: inline-flex; gap: 6px; align-items: baseline; font-size: 12px; color: #ffcfb8; background: rgba(255,66,2,0.16); border: 1px solid rgba(255,66,2,0.35); border-radius: 999px; padding: 4px 10px; }}
      .market-live strong {{ color: #ffffff; }}
      .market-live span {{ color: #ffcfb8; }}
      .market-grid {{ display: grid; grid-template-columns: 1fr; gap: 14px; }}
      .market-card {{ border: 1px solid #1e1e1e; border-radius: 12px; padding: 16px; background: #101010; }}
      .market-card h3 {{ margin: 0 0 12px; color: #ffffff; font-size: 18px; font-weight: 700; }}
      .market-graph-note {{ margin: 0 0 10px; font-size: 12px; color: #b8b8b8; line-height: 1.5; }}
      .market-price-card {{ grid-column: 1 / -1; }}
      .market-price-svg {{ width: 100%; height: auto; display: block; }}
      .market-bars {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(84px, 1fr)); gap: 12px; align-items: end; min-height: 220px; }}
      .market-bar-item {{ display: flex; flex-direction: column; align-items: center; gap: 6px; }}
      .market-bar-track {{ width: 100%; max-width: 92px; height: 170px; border-radius: 8px; border: 1px solid #2a2a2a; background: linear-gradient(180deg, #141414 0%, #0b0b0b 100%); overflow: hidden; display: flex; flex-direction: column; justify-content: flex-end; }}
      .market-bar-fill {{ width: 100%; background: linear-gradient(180deg, #ff8b61 0%, #ff4202 100%); }}
      .market-liq-track {{ justify-content: flex-end; }}
      .market-liq-short {{ width: 100%; background: #8f3a1d; }}
      .market-liq-long {{ width: 100%; background: #ff4202; }}
      .market-bar-label {{ margin: 0; font-size: 11px; color: #d0d0d0; text-align: center; line-height: 1.3; }}
      .market-bar-value {{ margin: 0; font-size: 11px; color: #ffcfb8; }}
      .market-legend {{ display: flex; gap: 14px; margin: 0 0 8px; font-size: 11px; color: #bcbcbc; }}
      .market-legend .dot {{ display: inline-block; width: 8px; height: 8px; border-radius: 50%; margin-right: 6px; }}
      .market-legend .dot-long {{ background: #ff4202; }}
      .market-legend .dot-short {{ background: #8f3a1d; }}
      .market-circ-card .market-circ-value {{ margin: 0 0 8px; font-size: 22px; font-weight: 700; color: #ffffff; }}
      .market-progress-track {{ width: 100%; height: 14px; border: 1px solid #2a2a2a; border-radius: 999px; background: #0c0c0c; overflow: hidden; }}
      .market-progress-fill {{ height: 100%; background: linear-gradient(90deg, #ff4202 0%, #ff8b61 100%); }}
      .market-own-card {{ grid-column: 1 / -1; }}
      .market-own-total {{ margin: 0; color: #ff4202; font-size: 22px; font-weight: 700; line-height: 1.15; }}
      .market-own-bar {{ width: 100%; height: 18px; border: 1px solid #2a2a2a; border-radius: 999px; overflow: hidden; display: flex; margin: 10px 0; }}
      .market-own-segment {{ height: 100%; min-width: 2px; }}
      .market-own-legend {{ display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 6px 10px; }}
      .market-own-item {{ display: grid; grid-template-columns: 10px 1fr auto; align-items: center; gap: 6px; }}
      .market-own-dot {{ width: 8px; height: 8px; border-radius: 999px; }}
      .market-own-name {{ font-size: 11px; color: #d0d0d0; }}
      .market-own-value {{ font-size: 11px; color: #ffcfb8; }}
      .market-subnote {{ margin: 8px 0 0; font-size: 12px; color: #b9b9b9; }}
      .market-note {{ margin: 8px 0 0; font-size: 11px; color: #8e8e8e; }}
      .market-empty {{ margin: 0; color: #9a9a9a; font-size: 12px; }}
      .tldr {{ background: #fff8ec; border-top: 2px solid #ff4202; }}
      .conclusion {{ background: #fff7f3; border-top: 2px solid #ff4202; }}
      .footer {{ padding: 0; font-size: 12px; color: #7a7a7a; }}
      .footer-legal {{ padding: 16px 32px 12px; }}
      .footer-dark {{ background: #0f0f0f; padding: 22px 32px; }}
      .footer-bar {{ display: flex; align-items: center; justify-content: space-between; }}
      .footer-site-link {{ display: inline-flex; align-items: center; gap: 10px; text-decoration: none; }}
      .footer-site-link img {{ width: 36px; height: 36px; border-radius: 10px; object-fit: contain; }}
      .footer-site-name {{ display: block; color: #ffffff; font-size: 13px; font-weight: 700; }}
      .footer-site-url {{ display: block; color: #ff4202; font-size: 11px; }}
      .footer-socials {{ display: flex; gap: 8px; }}
      .footer-social-btn {{ display: inline-flex; align-items: center; justify-content: center; width: 36px; height: 36px; background: rgba(255,255,255,0.08); border-radius: 10px; }}
      .footer-social-btn img {{ width: 18px; height: 18px; filter: brightness(0) invert(1); }}
      .footer-copy {{ margin: 14px 0 0; padding-top: 14px; border-top: 1px solid rgba(255,255,255,0.08); font-size: 11px; color: rgba(255,255,255,0.3); text-align: center; }}
      @media (max-width: 720px) {{
        .toolbar {{ padding: 10px 16px 6px; box-sizing: border-box; }}
        .wrapper {{ padding: 16px 0; }}
        .container {{ width: 100%; max-width: 100%; border-radius: 0; }}
        .image.mode-full {{ margin-left: -20px; margin-right: -20px; width: calc(100% + 40px); }}
        .extra-image-item.mode-full {{ margin-left: -20px; margin-right: -20px; width: calc(100% + 40px); }}
        .image img.mode-tight, .extra-image-item img.mode-tight {{ max-width: 260px !important; }}
        .image img, .extra-images img {{ max-width: 100%; margin: 0 auto; }}
        .section {{ padding: 18px 20px; }}
        .section h2.point-title {{ font-size: 24px; color: #ff4202; }}
        .hero-content {{ padding: 20px; }}
        .hero-title {{ font-size: 22px; }}
        .market-grid {{ grid-template-columns: 1fr; }}
        .market-price-card {{ grid-column: auto; }}
        .market-own-card {{ grid-column: auto; }}
        .market-own-legend {{ grid-template-columns: 1fr; }}
        .footer-legal {{ padding: 14px 20px 10px; }}
        .footer-dark {{ padding: 18px 20px; }}
        .footer-bar {{ flex-direction: column; align-items: flex-start; gap: 14px; }}
      }}
      @media (max-width: 920px) and (orientation: landscape) {{
        .container {{ width: 100%; max-width: 100%; border-radius: 0; }}
        .image img {{ width: 100%; height: auto; max-height: none; }}
        .extra-images img {{ width: 100%; height: auto; max-height: none; }}
        .image.mode-full {{ margin-left: -18px; margin-right: -18px; width: calc(100% + 36px); }}
        .extra-image-item.mode-full {{ margin-left: -18px; margin-right: -18px; width: calc(100% + 36px); }}
        .image.mode-full img, .extra-image-item.mode-full img {{ width: 100% !important; max-width: none !important; max-height: none !important; }}
        .image.mode-tight img, .extra-image-item.mode-tight img {{ width: 100% !important; max-width: 260px !important; max-height: 32vh !important; }}
        .section {{ padding: 14px 18px; }}
        .hero {{ min-height: 180px; }}
      }}
      @media print {{
        .no-print {{ display: none !important; }}
        @page {{ margin: 0; size: auto; }}
        html, body {{ margin: 0 !important; padding: 0 !important; background: #ffffff !important; width: 100% !important; }}
        .wrapper {{ background: #ffffff !important; padding: 0 !important; width: 100% !important; }}
        .container {{ width: 100% !important; max-width: 100% !important; border: 0 !important; border-radius: 0 !important; overflow: visible !important; }}
        table, tr, td, th, div, section, article, p, h1, h2, h3, h4, h5, h6, img, figure, blockquote, ul, ol, li {{ break-inside: avoid !important; page-break-inside: avoid !important; }}
        .section, .snapshot-card, .market-card, .snapshot-treas-table-wrap {{ break-inside: auto !important; page-break-inside: auto !important; }}
        .section h2, .section h3, .snapshot-card h3 {{ break-after: avoid-page !important; page-break-after: avoid !important; }}
        .section h2 + *, .section h3 + *, .snapshot-card h3 + * {{ break-before: avoid-page !important; page-break-before: avoid !important; }}
        p, li {{ orphans: 3; widows: 3; }}
        img {{ max-width: 100% !important; height: auto !important; display: block !important; }}
        canvas {{ break-inside: avoid !important; page-break-inside: avoid !important; max-width: 100% !important; }}
        a {{ color: #ff4202 !important; text-decoration: none !important; }}
      }}
    </style>
  </head>
  <body>
    <div class="toolbar no-print">
      <button class="download-pdf-btn" type="button" onclick="copyHtmlForEmail()">Copy HTML for Email</button>
      <button class="download-pdf-btn" type="button" onclick="openMailchimpCampaign()">Send via Mailchimp</button>
      <button class="download-pdf-btn" type="button" onclick="window.print()">Download PDF</button>
    </div>
    <table class="wrapper" role="presentation" width="100%" cellpadding="0" cellspacing="0">
      <tr>
        <td align="center">
          <table class="container" role="presentation" cellpadding="0" cellspacing="0">
            <tr>
              <td class="divider">&nbsp;</td>
            </tr>
            <tr>
              <td class="hero">
                <img class="hero-bg" src="{hero_image_url}" alt="">
                <div class="hero-gradient"></div>
                <div class="hero-content">
                  <div class="hero-logo"><img src="brand_orange_bg_transparent@2xSite.svg" alt="Globalite"></div>
                  <p class="hero-eyebrow">{eyebrow}</p>
                  <h1 class="hero-title">{title}</h1>
                  <p class="hero-subtitle">{subtitle}</p>
                  <span class="hero-badge">Written at block height: <strong>{block_height}</strong></span>
                </div>
              </td>
            </tr>
{points_html}
            <tr>
              <td class="section tldr">
                <h2>{tldr_title}</h2>
{tldr_content}
              </td>
            </tr>
            <tr>
              <td class="section conclusion">
                <h2>{conclusion_title}</h2>
{conclusion_content}
              </td>
            </tr>
{market_html}
            <tr>
              <td class="footer">
                <div class="footer-legal">
                  <p>{footer_line}</p>
                  <p>{address_line}</p>
                </div>
                <div class="footer-dark">
                  <div class="footer-bar">
                    <a class="footer-site-link" href="https://globalite.co" target="_blank">
                      <img src="{footer_logo_url}" alt="Globalite">
                      <span>
                        <span class="footer-site-name">Globalite</span>
                        <span class="footer-site-url">globalite.co</span>
                      </span>
                    </a>
                    <div class="footer-socials">
                      <a class="footer-social-btn" href="https://x.com/globalite_sa"><img src="{footer_x_icon}" alt="X"></a>
                      <a class="footer-social-btn" href="https://www.linkedin.com/company/globalite-sa"><img src="{footer_linkedin_icon}" alt="LinkedIn"></a>
                    </div>
                  </div>
                  <p class="footer-copy">© 2026 Globalite SA. All rights reserved.</p>
                </div>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
    <script>
      (function () {{
        var params = new URLSearchParams(window.location.search);
        var refreshSeconds = Number(params.get("refresh"));
        if (!Number.isFinite(refreshSeconds) || refreshSeconds < 5) {{
          return;
        }}
        window.setInterval(function () {{
          window.location.reload();
        }}, refreshSeconds * 1000);
      }})();
    </script>
    <script>
      function openMailchimpCampaign() {{
        window.open('https://admin.mailchimp.com/campaigns/#/create-campaign/', '_blank', 'noopener,noreferrer');
      }}

      function copyHtmlForEmail() {{
        var container = document.querySelector('.container');
        if (!container) {{
          return;
        }}
        var html = container.outerHTML;
        if (navigator.clipboard && window.isSecureContext) {{
          navigator.clipboard.writeText(html).then(function () {{
            alert('Newsletter HTML copied! Paste it into Mailchimp > Email > Code your own.');
          }}).catch(function () {{
            fallbackCopyText(html);
          }});
          return;
        }}
        fallbackCopyText(html);
      }}

      function fallbackCopyText(text) {{
        var textarea = document.createElement('textarea');
        textarea.value = text;
        textarea.setAttribute('readonly', '');
        textarea.style.position = 'fixed';
        textarea.style.left = '-9999px';
        document.body.appendChild(textarea);
        textarea.select();
        try {{
          document.execCommand('copy');
          alert('Newsletter HTML copied! Paste it into Mailchimp > Email > Code your own.');
        }} finally {{
          document.body.removeChild(textarea);
        }}
      }}
    </script>
    <script>
      (function () {{
        function detectModeFromPath(path) {{
          var raw = String(path || '').split('?')[0].split('#')[0].toLowerCase();
          if (/_full\.[a-z0-9]{2,5}$/.test(raw)) return 'full';
          if (/_tight\.[a-z0-9]{2,5}$/.test(raw)) return 'tight';
          return 'plain';
        }}

        function applyDisplayMode(img) {{
          var mode = detectModeFromPath(img.currentSrc || img.getAttribute('src') || '');
          img.classList.remove('mode-full', 'mode-tight', 'mode-plain');
          img.classList.add('mode-' + mode);
          var wrap = img.closest('.image, .extra-image-item');
          if (wrap) {{
            wrap.classList.remove('mode-full', 'mode-tight', 'mode-plain');
            wrap.classList.add('mode-' + mode);
          }}
        }}

        function applyMobileSpecificSources() {{
          if (!window.matchMedia("(max-width: 720px)").matches) {{
            return;
          }}
          document.querySelectorAll('.image img, .extra-images img').forEach(function (img) {{
            if (img.dataset.mobileApplied === '1') {{
              return;
            }}
            var mobileList = (img.dataset.mobileSrcs || '').split('|').filter(Boolean);
            if (!mobileList.length) {{
              return;
            }}
            var current = img.getAttribute('src') || '';
            var existingFallbacks = (img.dataset.fallbacks || '').split('|').filter(Boolean);
            var mergedFallbacks = mobileList.slice(1);
            if (current) {{
              mergedFallbacks.push(current);
            }}
            mergedFallbacks = mergedFallbacks.concat(existingFallbacks);
            img.dataset.fallbacks = mergedFallbacks.join('|');
            img.dataset.mobileApplied = '1';
            img.src = mobileList[0];
          }});
        }}

        function tagOrientation(img) {{
          if (!(img.naturalWidth > 0) || !(img.naturalHeight > 0)) {{
            return;
          }}
          img.classList.remove('is-portrait', 'is-standard', 'is-wide');
          var ratio = img.naturalWidth / img.naturalHeight;
          if (ratio < 0.85) {{
            img.classList.add('is-portrait');
          }} else if (ratio > 1.85) {{
            img.classList.add('is-wide');
          }} else {{
            img.classList.add('is-standard');
          }}
        }}
        applyMobileSpecificSources();
        document.querySelectorAll('.image img, .extra-images img').forEach(function (img) {{
          if (img.complete && img.naturalWidth > 0) {{
            applyDisplayMode(img);
            tagOrientation(img);
          }} else {{
            img.addEventListener('load', function () {{ applyDisplayMode(img); tagOrientation(img); }});
          }}
        }});
      }})();
    </script>
  </body>
</html>
"""


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate newsletter HTML from Excel or Google Sheets."
    )
    parser.add_argument(
        "--xlsx",
        default="newsletter_data.xlsx",
        help="Input workbook path (default: newsletter_data.xlsx). Ignored when --google-sheet is set.",
    )
    parser.add_argument(
        "--google-sheet",
        default="",
        help="Google Sheet URL or ID. When set, this is used as the source instead of --xlsx.",
    )
    parser.add_argument(
        "--google-meta-tab",
        default="meta",
        help="Tab name for meta values in Google Sheets (default: meta).",
    )
    parser.add_argument(
        "--google-points-tab",
        default="points",
        help="Tab name for points in Google Sheets (default: points).",
    )
    parser.add_argument(
        "--google-live-prices-tab",
        default="live_prices",
        help="Tab name for live prices in Google Sheets (default: live_prices).",
    )
    parser.add_argument(
        "--google-btc-price-tab",
        default="BTC Price",
        help='Tab name for BTC price chart data (default: "BTC Price").',
    )
    parser.add_argument(
        "--google-treasuries-tab",
        default="Treasuries",
        help='Tab name for treasuries chart data (default: "Treasuries").',
    )
    parser.add_argument(
        "--google-circulating-tab",
        default="Circulating BTC",
        help='Tab name for circulating supply data (default: "Circulating BTC").',
    )
    parser.add_argument(
        "--google-liquidations-tab",
        default="Liquidations",
        help='Tab name for liquidations chart data (default: "Liquidations").',
    )
    parser.add_argument(
        "--google-ownership-tab",
        default="Distribution",
        help='Tab name for ownership distribution data (default: "Distribution").',
    )
    parser.add_argument(
        "--google-graph-settings-tab",
        default="graph_settings",
        help='Tab name for graph controls (default: "graph_settings").',
    )
    parser.add_argument(
        "--out",
        default="newsletter.html",
        help="Output HTML file path (default: newsletter.html)",
    )
    parser.add_argument(
        "--init-template",
        action="store_true",
        help="Create a starter workbook template and exit.",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Overwrite files when used with --init-template.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    base_dir = Path(__file__).resolve().parent
    xlsx_path = Path(args.xlsx)
    if not xlsx_path.is_absolute():
        xlsx_path = base_dir / xlsx_path
    out_path = Path(args.out)
    if not out_path.is_absolute():
        out_path = base_dir / out_path
    google_sheet_ref = normalize_text(args.google_sheet)

    if args.init_template:
        if google_sheet_ref:
            raise ValueError("--init-template only works with local Excel files.")
        xlsx_path.parent.mkdir(parents=True, exist_ok=True)
        create_template_workbook(xlsx_path, force=args.force)
        print(f"Template created: {xlsx_path}")
        return 0

    meta: dict[str, str]
    points: list[Point]
    live_btc: LiveBtcPrice | None
    btc_price_points: list[BtcPricePoint]
    treasury_bars: list[TreasuryBar]
    circulating: CirculatingMetric | None
    liquidation_bars: list[LiquidationBar]
    ownership_segments: list[OwnershipSegment]
    graph_settings: dict[str, GraphSetting]
    backup_path: Path

    def parse_optional(label: str, parser, fallback):
        try:
            return parser()
        except ValueError as error:
            print(f"Warning: ignoring {label} tab: {error}", file=sys.stderr)
            return fallback

    if google_sheet_ref:
        sheet_id = extract_google_sheet_id(google_sheet_ref)
        meta_rows = fetch_google_sheet_rows(sheet_id, args.google_meta_tab, required=True)
        points_rows = fetch_google_sheet_rows(
            sheet_id, args.google_points_tab, required=True
        )
        live_prices_rows = fetch_google_sheet_rows(
            sheet_id, args.google_live_prices_tab, required=False
        )
        btc_price_rows = fetch_google_sheet_rows(
            sheet_id, args.google_btc_price_tab, required=False
        )
        treasuries_rows = fetch_google_sheet_rows(
            sheet_id, args.google_treasuries_tab, required=False
        )
        circulating_rows = fetch_google_sheet_rows(
            sheet_id, args.google_circulating_tab, required=False
        )
        liquidations_rows = fetch_google_sheet_rows(
            sheet_id, args.google_liquidations_tab, required=False
        )
        ownership_rows = fetch_google_sheet_rows(
            sheet_id, args.google_ownership_tab, required=False
        )
        graph_settings_rows = fetch_google_sheet_rows(
            sheet_id, args.google_graph_settings_tab, required=False
        )

        history_dir = base_dir / "history"
        history_dir.mkdir(exist_ok=True)
        backup_path = next_history_workbook_path(history_dir)
        write_google_snapshot_workbook(
            backup_path,
            meta_rows,
            points_rows,
            live_prices_rows,
            btc_price_rows,
            treasuries_rows,
            circulating_rows,
            liquidations_rows,
            ownership_rows,
            graph_settings_rows,
        )

        meta = read_meta(TabularSheet(meta_rows))
        points = read_points(TabularSheet(points_rows))
        graph_settings = (
            parse_optional(
                args.google_graph_settings_tab,
                lambda: read_graph_settings(TabularSheet(graph_settings_rows)),
                default_graph_settings(),
            )
            if graph_settings_rows
            else default_graph_settings()
        )
        btc_setting = get_graph_setting(graph_settings, "btc_price")
        treasuries_setting = get_graph_setting(graph_settings, "treasuries")
        circulating_setting = get_graph_setting(graph_settings, "circulating_btc")
        liquidations_setting = get_graph_setting(graph_settings, "liquidations")
        ownership_setting = get_graph_setting(graph_settings, "ownership")
        live_btc = (
            parse_optional(
                args.google_live_prices_tab,
                lambda: read_live_btc_price(TabularSheet(live_prices_rows)),
                None,
            )
            if live_prices_rows
            else None
        )
        btc_price_points = (
            parse_optional(
                args.google_btc_price_tab,
                lambda: read_btc_price_points(
                    TabularSheet(btc_price_rows),
                    limit=btc_setting.top_n or 60,
                    setting=btc_setting,
                ),
                [],
            )
            if btc_price_rows
            else []
        )
        if not btc_price_points and live_prices_rows:
            btc_price_points = parse_optional(
                args.google_live_prices_tab,
                lambda: read_btc_price_points_from_live_prices(
                    TabularSheet(live_prices_rows),
                    limit=btc_setting.top_n or 60,
                ),
                [],
            )
        treasury_bars = (
            parse_optional(
                args.google_treasuries_tab,
                lambda: read_treasury_bars(
                    TabularSheet(treasuries_rows),
                    limit=treasuries_setting.top_n or 6,
                    setting=treasuries_setting,
                ),
                [],
            )
            if treasuries_rows
            else []
        )
        circulating = (
            parse_optional(
                args.google_circulating_tab,
                lambda: read_circulating_metric(
                    TabularSheet(circulating_rows),
                    setting=circulating_setting,
                ),
                None,
            )
            if circulating_rows
            else None
        )
        liquidation_bars = (
            parse_optional(
                args.google_liquidations_tab,
                lambda: read_liquidation_bars(
                    TabularSheet(liquidations_rows),
                    limit=liquidations_setting.top_n or 6,
                    setting=liquidations_setting,
                ),
                [],
            )
            if liquidations_rows
            else []
        )
        ownership_segments = (
            parse_optional(
                args.google_ownership_tab,
                lambda: read_ownership_segments(
                    TabularSheet(ownership_rows),
                    setting=ownership_setting,
                ),
                [],
            )
            if ownership_rows
            else []
        )
    else:
        if not xlsx_path.exists():
            raise FileNotFoundError(
                f"Workbook not found: {xlsx_path}. "
                "Run with --init-template first to create it."
            )

        history_dir = xlsx_path.parent / "history"
        history_dir.mkdir(exist_ok=True)
        backup_path = next_history_workbook_path(history_dir)
        shutil.copy2(xlsx_path, backup_path)

        workbook = load_workbook(xlsx_path, data_only=True)
        if "meta" not in workbook.sheetnames:
            raise ValueError("Workbook is missing required sheet: meta")
        if "points" not in workbook.sheetnames:
            raise ValueError("Workbook is missing required sheet: points")

        meta = read_meta(workbook["meta"])
        points = read_points(workbook["points"])
        graph_settings_sheet_name = ""
        for candidate in ["graph_settings", "Graph Settings", "graphs", "Graphs"]:
            if candidate in workbook.sheetnames:
                graph_settings_sheet_name = candidate
                break
        graph_settings = (
            parse_optional(
                graph_settings_sheet_name,
                lambda: read_graph_settings(workbook[graph_settings_sheet_name]),
                default_graph_settings(),
            )
            if graph_settings_sheet_name
            else default_graph_settings()
        )
        btc_setting = get_graph_setting(graph_settings, "btc_price")
        treasuries_setting = get_graph_setting(graph_settings, "treasuries")
        circulating_setting = get_graph_setting(graph_settings, "circulating_btc")
        liquidations_setting = get_graph_setting(graph_settings, "liquidations")
        ownership_setting = get_graph_setting(graph_settings, "ownership")
        live_btc = (
            parse_optional("live_prices", lambda: read_live_btc_price(workbook["live_prices"]), None)
            if "live_prices" in workbook.sheetnames
            else None
        )
        btc_price_points = (
            parse_optional(
                "BTC Price",
                lambda: read_btc_price_points(
                    workbook["BTC Price"],
                    limit=btc_setting.top_n or 60,
                    setting=btc_setting,
                ),
                [],
            )
            if "BTC Price" in workbook.sheetnames
            else []
        )
        if not btc_price_points and "live_prices" in workbook.sheetnames:
            btc_price_points = parse_optional(
                "live_prices",
                lambda: read_btc_price_points_from_live_prices(
                    workbook["live_prices"],
                    limit=btc_setting.top_n or 60,
                ),
                [],
            )
        treasury_bars = (
            parse_optional(
                "Treasuries",
                lambda: read_treasury_bars(
                    workbook["Treasuries"],
                    limit=treasuries_setting.top_n or 6,
                    setting=treasuries_setting,
                ),
                [],
            )
            if "Treasuries" in workbook.sheetnames
            else []
        )
        circulating = (
            parse_optional(
                "Circulating BTC",
                lambda: read_circulating_metric(
                    workbook["Circulating BTC"],
                    setting=circulating_setting,
                ),
                None,
            )
            if "Circulating BTC" in workbook.sheetnames
            else None
        )
        liquidation_bars = (
            parse_optional(
                "Liquidations",
                lambda: read_liquidation_bars(
                    workbook["Liquidations"],
                    limit=liquidations_setting.top_n or 6,
                    setting=liquidations_setting,
                ),
                [],
            )
            if "Liquidations" in workbook.sheetnames
            else []
        )
        ownership_sheet_name = (
            "Distribution"
            if "Distribution" in workbook.sheetnames
            else ("distribution" if "distribution" in workbook.sheetnames else "")
        )
        ownership_segments = (
            parse_optional(
                ownership_sheet_name,
                lambda: read_ownership_segments(
                    workbook[ownership_sheet_name],
                    setting=ownership_setting,
                ),
                [],
            )
            if ownership_sheet_name
            else []
        )

    if circulating is None:
        max_supply = parse_number(meta.get("max_supply_btc", "21000000"), default=21_000_000)
        if max_supply <= 0:
            max_supply = 21_000_000
        circulating_value = parse_number(
            meta.get("circulating_supply_btc", "0"), default=0
        )
        if circulating_value > 0:
            circulating = CirculatingMetric(
                as_of_date="",
                circulating_supply_btc=circulating_value,
                max_supply_btc=max_supply,
                note="",
            )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    html_output = render_html(
        meta,
        points,
        btc_price_points,
        treasury_bars,
        circulating,
        liquidation_bars,
        ownership_segments,
        graph_settings,
        live_btc,
        out_path.parent,
    )

    out_path.write_text(html_output, encoding="utf-8")
    print(
        f"Generated {out_path} with {len(points)} points "
        f"(max allowed: {MAX_POINTS})."
    )
    email_ready_output: str | None = None
    if premailer_transform is not None:
        email_out_path = out_path.with_name(f"{out_path.stem}_email{out_path.suffix}")
        email_ready_output = premailer_transform(html_output)
        email_out_path.write_text(email_ready_output, encoding="utf-8")
        print(f"Generated email-safe HTML: {email_out_path}")
    else:
        print(
            "Warning: premailer not installed; skipped newsletter_email.html generation. "
            "Install with: pip install premailer",
            file=sys.stderr,
        )
    public_dir = base_dir / "public"
    if public_dir.exists():
        public_newsletter_path = public_dir / "newsletter.html"
        public_newsletter_path.write_text(html_output, encoding="utf-8")
        print(f"Published static asset: {public_newsletter_path}")
        if email_ready_output is not None:
            public_email_path = public_dir / "newsletter_email.html"
            public_email_path.write_text(email_ready_output, encoding="utf-8")
            print(f"Published static email asset: {public_email_path}")
    print(f"Source snapshot saved: {backup_path}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:  # pragma: no cover
        print(f"Error: {error}", file=sys.stderr)
        raise SystemExit(1)
