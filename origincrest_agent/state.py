"""Local SQLite state: what the agent has found, researched and drafted.

This makes every workflow safe to re-run: research is cached, a draft is never
created twice for the same step, and rows found while a workbook was open in
Excel are kept here until they can be written.
"""

from __future__ import annotations

import json
import sqlite3
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

SCHEMA = """
CREATE TABLE IF NOT EXISTS discovered (
    id INTEGER PRIMARY KEY,
    company TEXT NOT NULL,
    kind TEXT NOT NULL,
    keys TEXT NOT NULL,
    source_url TEXT,
    found_on TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS pending_rows (
    id INTEGER PRIMARY KEY,
    workbook TEXT NOT NULL,
    tab TEXT NOT NULL,
    kind TEXT NOT NULL,
    values_json TEXT NOT NULL,
    created_at TEXT NOT NULL,
    written_at TEXT
);
CREATE TABLE IF NOT EXISTS research (
    company_key TEXT PRIMARY KEY,
    data_json TEXT NOT NULL,
    researched_on TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS drafts (
    id INTEGER PRIMARY KEY,
    company_key TEXT NOT NULL,
    workflow TEXT NOT NULL,
    step INTEGER NOT NULL,
    to_addr TEXT NOT NULL,
    subject TEXT NOT NULL,
    message_id TEXT NOT NULL,
    row_ref TEXT,
    created_on TEXT NOT NULL,
    UNIQUE (company_key, workflow, step)
);
CREATE TABLE IF NOT EXISTS suppression (
    address TEXT PRIMARY KEY,
    reason TEXT,
    added_on TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS runs (
    id INTEGER PRIMARY KEY,
    workflow TEXT NOT NULL,
    started_at TEXT NOT NULL,
    finished_at TEXT,
    summary_json TEXT
);
"""


class State:
    def __init__(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)
        self.path = path
        self.db = sqlite3.connect(path)
        self.db.row_factory = sqlite3.Row
        self.db.executescript(SCHEMA)
        self.db.commit()

    def close(self) -> None:
        self.db.close()

    # discovery ---------------------------------------------------------
    def discovered_keys(self) -> set[str]:
        keys: set[str] = set()
        for row in self.db.execute("SELECT keys FROM discovered"):
            keys.update(json.loads(row["keys"]))
        return keys

    def add_discovered(self, company: str, kind: str, keys: set[str], source_url: str, found_on: date) -> None:
        self.db.execute(
            "INSERT INTO discovered (company, kind, keys, source_url, found_on) VALUES (?,?,?,?,?)",
            (company, kind, json.dumps(sorted(keys)), source_url, found_on.isoformat()),
        )
        self.db.commit()

    def queue_row(self, workbook: str, tab: str, kind: str, values: dict[str, Any]) -> int:
        cur = self.db.execute(
            "INSERT INTO pending_rows (workbook, tab, kind, values_json, created_at) VALUES (?,?,?,?,?)",
            (workbook, tab, kind, json.dumps(values, default=_json_default), datetime.now().isoformat()),
        )
        self.db.commit()
        return cur.lastrowid

    def pending_rows(self, workbook: str) -> list[sqlite3.Row]:
        return list(self.db.execute(
            "SELECT * FROM pending_rows WHERE workbook=? AND written_at IS NULL ORDER BY id", (workbook,)))

    def mark_rows_written(self, ids: list[int]) -> None:
        now = datetime.now().isoformat()
        self.db.executemany("UPDATE pending_rows SET written_at=? WHERE id=?", [(now, i) for i in ids])
        self.db.commit()

    # research ----------------------------------------------------------
    def research(self, company_key: str, max_age_days: int) -> dict | None:
        row = self.db.execute("SELECT * FROM research WHERE company_key=?", (company_key,)).fetchone()
        if not row:
            return None
        if date.fromisoformat(row["researched_on"]) < date.today() - timedelta(days=max_age_days):
            return None
        return json.loads(row["data_json"])

    def save_research(self, company_key: str, data: dict, on: date) -> None:
        self.db.execute(
            "INSERT OR REPLACE INTO research (company_key, data_json, researched_on) VALUES (?,?,?)",
            (company_key, json.dumps(data), on.isoformat()),
        )
        self.db.commit()

    # drafts ------------------------------------------------------------
    def draft(self, company_key: str, workflow: str, step: int) -> sqlite3.Row | None:
        return self.db.execute(
            "SELECT * FROM drafts WHERE company_key=? AND workflow=? AND step=?", (company_key, workflow, step)
        ).fetchone()

    def latest_draft(self, company_key: str) -> sqlite3.Row | None:
        return self.db.execute(
            "SELECT * FROM drafts WHERE company_key=? ORDER BY created_on DESC, id DESC LIMIT 1", (company_key,)
        ).fetchone()

    def record_draft(self, company_key: str, workflow: str, step: int, to_addr: str, subject: str,
                     message_id: str, row_ref: str, on: date) -> None:
        self.db.execute(
            "INSERT OR REPLACE INTO drafts (company_key, workflow, step, to_addr, subject, message_id, row_ref,"
            " created_on) VALUES (?,?,?,?,?,?,?,?)",
            (company_key, workflow, step, to_addr, subject, message_id, row_ref, on.isoformat()),
        )
        self.db.commit()

    # suppression -------------------------------------------------------
    def suppress(self, address: str, reason: str) -> None:
        address = address.strip().lower()
        if not address:
            return
        self.db.execute(
            "INSERT OR IGNORE INTO suppression (address, reason, added_on) VALUES (?,?,?)",
            (address, reason, date.today().isoformat()),
        )
        self.db.commit()

    def is_suppressed(self, email: str) -> bool:
        email = email.strip().lower()
        if not email:
            return False
        domain = email.rsplit("@", 1)[-1]
        row = self.db.execute(
            "SELECT 1 FROM suppression WHERE address=? OR address=?", (email, "@" + domain)
        ).fetchone()
        return row is not None

    def suppression_list(self) -> list[sqlite3.Row]:
        return list(self.db.execute("SELECT * FROM suppression ORDER BY added_on"))

    # runs --------------------------------------------------------------
    def start_run(self, workflow: str) -> int:
        cur = self.db.execute(
            "INSERT INTO runs (workflow, started_at) VALUES (?,?)", (workflow, datetime.now().isoformat()))
        self.db.commit()
        return cur.lastrowid

    def finish_run(self, run_id: int, summary: dict) -> None:
        self.db.execute(
            "UPDATE runs SET finished_at=?, summary_json=? WHERE id=?",
            (datetime.now().isoformat(), json.dumps(summary, default=_json_default), run_id),
        )
        self.db.commit()


def _json_default(value: Any) -> Any:
    if isinstance(value, (date, datetime)):
        return {"__date__": value.isoformat()[:10]}
    raise TypeError(f"Not JSON serialisable: {type(value)}")


def decode_values(values_json: str) -> dict[str, Any]:
    def hook(obj: dict) -> Any:
        if set(obj) == {"__date__"}:
            return date.fromisoformat(obj["__date__"])
        return obj

    return json.loads(values_json, object_hook=hook)
