"""OriginCrest's current coffee range, read from your website and/or catalogue files.

List them under offer.range_sources in config.yaml: web pages (your shop / range pages) and files such
as an Excel or CSV price list, a Word document or a text file. The agent reads them at most every
`range_refresh_hours`, keeps a copy, and gives the text to Claude for fit scores, intro emails and
follow-ups, so they always reflect what you sell now. If a source can't be read, the last good copy is
used and the run report says so.
"""

from __future__ import annotations

import csv
import hashlib
import html
import json
import re
import zipfile
from datetime import datetime, timedelta
from pathlib import Path

import requests

from .config import expand_path
from .utils import log

USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) OriginCrestAgent/1.0"
_BLOCK_TAGS = re.compile(r"<\s*(br|/p|/div|/li|/tr|/h[1-6]|/section|/article|/table|/ul|/ol)\b[^>]*>", re.I)
_CELL_TAGS = re.compile(r"<\s*/t[dh]\s*>", re.I)
_DROP = re.compile(r"<(script|style|noscript|svg|head|nav|footer)\b[^>]*>.*?</\1\s*>", re.I | re.S)
_TAGS = re.compile(r"<[^>]+>")


def html_text(raw: str) -> str:
    """Readable text of a web page, one block per line (table cells separated by ' | ')."""
    text = _DROP.sub(" ", raw)
    text = _CELL_TAGS.sub(" | ", text)
    text = _BLOCK_TAGS.sub("\n", text)
    text = html.unescape(_TAGS.sub(" ", text))
    return _tidy(text)


def _tidy(text: str) -> str:
    lines = (re.sub(r"[ \t ]+", " ", line).strip(" |") for line in text.splitlines())
    out: list[str] = []
    for line in lines:
        if line and (not out or line != out[-1]):
            out.append(line)
    return "\n".join(out)


def _xlsx_text(path: Path) -> str:
    import openpyxl

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    parts = []
    try:
        for ws in wb.worksheets:
            rows = []
            for row in ws.iter_rows(values_only=True):
                cells = [str(v).strip() for v in row if v not in (None, "")]
                if cells:
                    rows.append(" | ".join(cells))
            if rows:
                parts.append(f"[{ws.title}]\n" + "\n".join(rows))
    finally:
        wb.close()
    return "\n\n".join(parts)


def _docx_text(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8", "replace")
    xml = re.sub(r"</w:p>|<w:br/>", "\n", xml)
    xml = re.sub(r"</w:tc>", " | ", xml)
    return _tidy(html.unescape(_TAGS.sub("", xml)))


def read_file(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xlsm"):
        return _xlsx_text(path)
    if suffix == ".csv":
        with open(path, newline="", encoding="utf-8-sig", errors="replace") as fh:
            return "\n".join(" | ".join(c.strip() for c in row if c.strip()) for row in csv.reader(fh) if any(row))
    if suffix == ".docx":
        return _docx_text(path)
    if suffix in (".htm", ".html"):
        return html_text(path.read_text(encoding="utf-8", errors="replace"))
    if suffix in (".txt", ".md", ""):
        return _tidy(path.read_text(encoding="utf-8", errors="replace"))
    raise ValueError(f"{path.name}: can't read {suffix} files - save it as Excel, CSV, Word or text")


def fetch_url(url: str, timeout: float = 20) -> str:
    resp = requests.get(url, timeout=timeout, headers={"User-Agent": USER_AGENT})
    resp.raise_for_status()
    return html_text(resp.text)


class OfferRange:
    def __init__(self, text: str = "", loaded: list[str] | None = None, problems: list[str] | None = None,
                 fetched_at: str = "", from_cache: bool = False, sizes: list | None = None):
        self.text = text
        self.loaded = loaded or []
        self.problems = problems or []
        self.fetched_at = fetched_at
        self.from_cache = from_cache
        self.sizes = [tuple(s) for s in sizes or []]  # (source, characters kept, characters read)

    def describe(self) -> str:
        if not self.text:
            return "no range loaded"
        when = self.fetched_at[:16].replace("T", " ")
        how = "last good copy from " + when if self.from_cache else "read " + when
        return f"{len(self.text):,} characters from {len(self.loaded)} source(s) ({how})"


def _shares(lengths: list[int], limit: int) -> list[int]:
    """Characters each source may keep: an equal share, with what short sources don't need passed on."""
    shares = [0] * len(lengths)
    remaining = limit
    order = sorted(range(len(lengths)), key=lengths.__getitem__)
    for done, i in enumerate(order):
        shares[i] = min(lengths[i], remaining // (len(order) - done))
        remaining -= shares[i]
    return shares


def load_range(cfg, *, refresh: bool = False, fetch=fetch_url, now: datetime | None = None) -> OfferRange:
    """The range text from offer.range_sources (cached), or an empty OfferRange if none are set."""
    offer = cfg["offer"]
    sources = [str(s).strip() for s in offer.get("range_sources") or [] if str(s).strip()]
    if not sources:
        return OfferRange()
    now = now or datetime.now()
    cache_path = cfg.data_dir / "offer_range.json"
    limit = int(offer.get("range_max_chars", 24000))
    key = hashlib.sha256("\n".join([*sources, str(limit)]).encode()).hexdigest()[:16]
    cached = None
    try:
        cached = json.loads(cache_path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        pass
    if cached and cached.get("key") == key and not refresh:
        age = now - datetime.fromisoformat(cached["fetched_at"])
        if age < timedelta(hours=float(offer.get("range_refresh_hours", 12))):
            return OfferRange(cached["text"], cached["loaded"], [], cached["fetched_at"],
                              sizes=cached.get("sizes"))

    found, problems = [], []
    for source in sources:
        try:
            if source.lower().startswith(("http://", "https://")):
                text = fetch(source)
            else:
                path = expand_path(source)
                if not path.exists():
                    raise FileNotFoundError(f"{path} not found")
                text = read_file(path)
        except (OSError, ValueError, KeyError, zipfile.BadZipFile, requests.RequestException) as exc:
            problems.append(f"{source}: {exc}")
            log.warning("Couldn't read range source %s: %s", source, exc)
            continue
        if text.strip():
            found.append((source, text.strip()))
        else:
            problems.append(f"{source}: no text found")

    if not found:
        if cached and cached.get("text"):
            problems.append("using the last good copy")
            return OfferRange(cached["text"], cached["loaded"], problems, cached["fetched_at"], from_cache=True,
                              sizes=cached.get("sizes"))
        return OfferRange(problems=problems)
    # Each source gets a fair share of the size limit, so a long website can't crowd out your price list.
    shares = _shares([len(text) for _, text in found], limit)
    parts, sizes = [], []
    for (source, text), share in zip(found, shares):
        kept = text if len(text) <= share else text[:share].rsplit("\n", 1)[0] + "\n(... trimmed to fit)"
        parts.append(f"--- From {source} ---\n{kept}")
        sizes.append((source, min(len(text), share), len(text)))
    text = "\n\n".join(parts)
    loaded = [source for source, _ in found]
    stamp = now.isoformat(timespec="seconds")
    try:
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        cache_path.write_text(json.dumps({"key": key, "fetched_at": stamp, "loaded": loaded, "text": text,
                                          "sizes": sizes}), encoding="utf-8")
    except OSError as exc:
        log.debug("Couldn't cache the range: %s", exc)
    return OfferRange(text, loaded, problems, stamp, sizes=sizes)
