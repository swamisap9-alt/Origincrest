"""Small shared helpers: dates, text cleanup, logging."""

from __future__ import annotations

import logging
import logging.handlers
import re
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

from dateutil import parser as date_parser

try:
    from zoneinfo import ZoneInfo
except ImportError:  # pragma: no cover - Python < 3.9
    ZoneInfo = None  # type: ignore

log = logging.getLogger("origincrest")


def today_in(tz_name: str) -> date:
    if ZoneInfo is None:
        return date.today()
    return datetime.now(ZoneInfo(tz_name)).date()


def to_date(value: Any) -> date | None:
    """Best-effort conversion of a spreadsheet cell to a date (Australian day-first)."""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        # Excel serial date stored as a plain number.
        if 20000 < value < 80000:
            return date(1899, 12, 30) + timedelta(days=int(value))
        return None
    text = str(value).strip()
    if not text or not re.search(r"\d", text):
        return None
    try:
        return date_parser.parse(text, dayfirst=True, fuzzy=True).date()
    except (ValueError, OverflowError):
        return None


def to_int(value: Any, default: int = 0) -> int:
    if value is None or value == "":
        return default
    try:
        return int(float(str(value).strip()))
    except ValueError:
        return default


def clean(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def fmt_date(d: date | None) -> str:
    return d.strftime("%d/%m/%Y") if d else ""


def setup_logging(log_dir: Path, verbose: bool = False) -> None:
    log_dir.mkdir(parents=True, exist_ok=True)
    root = logging.getLogger("origincrest")
    if root.handlers:
        return
    root.setLevel(logging.DEBUG)
    fmt = logging.Formatter("%(asctime)s %(levelname)-7s %(message)s", "%Y-%m-%d %H:%M:%S")
    file_handler = logging.handlers.RotatingFileHandler(
        log_dir / "agent.log", maxBytes=2_000_000, backupCount=5, encoding="utf-8"
    )
    file_handler.setFormatter(fmt)
    file_handler.setLevel(logging.DEBUG)
    root.addHandler(file_handler)
    console = logging.StreamHandler()
    console.setFormatter(logging.Formatter("%(levelname)-7s %(message)s"))
    console.setLevel(logging.DEBUG if verbose else logging.INFO)
    root.addHandler(console)
