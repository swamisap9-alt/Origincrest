"""Run context shared by the workflows, plus the safe "commit to Excel" step.

Workflows read a snapshot of the workbooks, then describe their changes in a
ChangeSet. `commit()` waits for Excel to release each file, reloads it fresh
(so edits you saved meanwhile are kept), applies the changes and saves.
"""

from __future__ import annotations

import os
import time
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable

from .config import Config
from .dedupe import core_name, norm_email
from .state import State, decode_values
from .utils import fmt_date, log
from .workbook import Row, WorkbookLocked, WorkbookStore, build_dedupe_index, wait_until_unlocked


class Report:
    def __init__(self, workflow: str, today: date):
        self.workflow = workflow
        self.today = today
        self.sections: dict[str, list[str]] = {}
        self.counts: dict[str, int] = {}

    def add(self, section: str, line: str, quiet: bool = False) -> None:
        self.sections.setdefault(section, []).append(line)
        (log.debug if quiet else log.info)("[%s] %s", section, line)

    def count(self, key: str, n: int = 1) -> None:
        self.counts[key] = self.counts.get(key, 0) + n

    def write(self, directory: Path, usage: dict | None = None, history: bool = True) -> Path:
        """Write latest_<workflow>.md, plus a timestamped copy unless history=False."""
        directory.mkdir(parents=True, exist_ok=True)
        path = directory / f"{datetime.now():%Y-%m-%d_%H%M}_{self.workflow}.md"
        lines = [f"# OriginCrest agent - {self.workflow} - {fmt_date(self.today)}", ""]
        if self.counts:
            lines += ["## Summary", ""] + [f"- {k}: {v}" for k, v in self.counts.items()] + [""]
        for section, items in self.sections.items():
            lines += [f"## {section}", ""] + [f"- {item}" for item in items] + [""]
        if usage:
            lines += ["## Claude usage", ""] + [f"- {k}: {v}" for k, v in usage.items()] + [""]
        latest = directory / f"latest_{self.workflow}.md"
        latest.write_text("\n".join(lines), encoding="utf-8")
        if not history:
            return latest
        path.write_text("\n".join(lines), encoding="utf-8")
        return path


@dataclass
class RowUpdate:
    workbook: str
    tab: str
    index: int
    company: str
    email: str
    values: dict[str, Any]
    notes: list[str] = field(default_factory=list)


class ChangeSet:
    def __init__(self) -> None:
        self.updates: list[RowUpdate] = []
        self.touched: set[str] = set()

    def update(self, row: Row, values: dict[str, Any] | None = None, note: str = "") -> None:
        if not values and not note:
            return
        upd = RowUpdate(row.sheet.store.name, row.sheet.title, row.index, row.text("company"),
                        norm_email(row.text("email")), dict(values or {}), [note] if note else [])
        self.updates.append(upd)
        self.touched.add(upd.workbook)

    def touch(self, workbook: str) -> None:
        self.touched.add(workbook)

    def __bool__(self) -> bool:
        return bool(self.updates)


@dataclass
class RunContext:
    cfg: Config
    state: State
    today: date
    report: Report
    dry_run: bool = False
    agent_factory: Callable[[], Any] | None = None
    mailbox_factory: Callable[[], Any] | None = None
    verifier: Any = None
    _agent: Any = None

    @property
    def agent(self):
        if self._agent is None:
            if self.agent_factory is None:
                raise RuntimeError("No Claude agent configured")
            self._agent = self.agent_factory()
        return self._agent

    @property
    def backup_dir(self) -> Path:
        return self.cfg.data_dir / "backups"

    @property
    def preview_dir(self) -> Path:
        return self.cfg.data_dir / "previews"


def add_note(existing: Any, note: str, today: date, limit: int = 1500) -> str:
    """Prepend a dated note unless the same note is already there."""
    current = str(existing or "").strip()
    if note in current:
        return current
    entry = f"{today:%d/%m}: {note}"
    return (entry + (" | " + current if current else ""))[:limit]


def _locate(store: WorkbookStore, upd: RowUpdate) -> Row | None:
    view = store.view(upd.tab)
    if view is None:
        return None

    def same(row: Row) -> bool:
        if upd.email and norm_email(row.text("email")) == upd.email:
            return True
        return bool(upd.company) and core_name(row.text("company")) == core_name(upd.company)

    if upd.index <= view.last_data_row():
        candidate = Row(view, upd.index)
        if same(candidate):
            return candidate
    # The row moved (rows inserted/deleted/sorted since we read it): find it again.
    return next((row for row in view.rows() if same(row)), None)


def commit(ctx: RunContext, changes: ChangeSet) -> dict[str, str]:
    """Write row updates and queued new rows. Returns {workbook: outcome}."""
    cfg = ctx.cfg
    outcomes: dict[str, str] = {}
    names = set(changes.touched) | {n for n in cfg["workbooks"] if ctx.state.pending_rows(n)}
    for name in sorted(names):
        path = cfg.workbook_path(name)
        if not path.exists():
            outcomes[name] = "missing"
            ctx.report.add("Not written", f"{path} not found")
            continue
        updates = [u for u in changes.updates if u.workbook == name]
        pending = ctx.state.pending_rows(name)
        if not updates and not pending:
            continue
        if ctx.dry_run:
            outcomes[name] = "dry-run"
            ctx.report.add("Dry run", f"Would write {len(updates)} row update(s) and {len(pending)} new row(s) "
                                      f"to {path.name}")
            continue
        locked = wait_until_unlocked([path], cfg["excel"]["lock_wait_minutes"])
        if locked:
            outcomes[name] = "locked"
            ctx.report.add("Not written", f"{path.name}: {locked[path]}. Close it in Excel - the changes are "
                                          "kept and will be written on the next run.")
            continue
        store = WorkbookStore(name, path, cfg).load()
        missing = 0
        for upd in updates:
            row = _locate(store, upd)
            if row is None:
                missing += 1
                ctx.report.add("Not written", f"{upd.company} ({upd.tab}) - row not found in {path.name}")
                continue
            for key, value in upd.values.items():
                row.set(key, value)
            for note in upd.notes:
                row.set("agent_notes", add_note(row.get("agent_notes"), note, ctx.today))
        written_ids = []
        if pending:
            index = build_dedupe_index({name: store})
            for item in pending:
                values = decode_values(item["values_json"])
                dup = index.match(values.get("company", ""), values.get("website", ""), values.get("email", ""),
                                  values.get("phone", ""))
                if dup:
                    ctx.report.add("Skipped duplicates", f"{values.get('company')} - already in {dup}")
                else:
                    store.get_or_create(item["tab"], item["kind"]).append(values)
                    index.add(f"{item['tab']} (new)", values.get("company", ""), values.get("website", ""),
                              values.get("email", ""), values.get("phone", ""))
                written_ids.append(item["id"])
        try:
            store.save(ctx.backup_dir, cfg["excel"]["backup_keep"])
        except WorkbookLocked as exc:
            outcomes[name] = "locked"
            ctx.report.add("Not written", f"{exc}. The changes will be written on the next run.")
            continue
        ctx.state.mark_rows_written(written_ids)
        outcomes[name] = "saved"
        for warning in store.load_warnings:
            ctx.report.add("Workbook warnings", f"{path.name}: {warning} (a backup was saved before writing)")
        ctx.report.count(f"rows updated in {path.name}", len(updates) - missing)
        if pending:
            ctx.report.count(f"new rows written to {path.name}", len(pending))
    return outcomes


class AgentBusy(RuntimeError):
    pass


class AgentLock:
    """One workflow at a time (the scheduled tasks can overlap)."""

    def __init__(self, path: Path, wait_minutes: float = 45, stale_hours: float = 4):
        self.path = path
        self.wait_minutes = wait_minutes
        self.stale_hours = stale_hours

    def __enter__(self) -> "AgentLock":
        self.path.parent.mkdir(parents=True, exist_ok=True)
        deadline = time.monotonic() + self.wait_minutes * 60
        while True:
            try:
                fd = os.open(self.path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                os.write(fd, f"{os.getpid()} {datetime.now().isoformat()}".encode())
                os.close(fd)
                return self
            except FileExistsError:
                try:
                    age_h = (time.time() - self.path.stat().st_mtime) / 3600
                except FileNotFoundError:
                    continue
                if age_h > self.stale_hours:
                    log.warning("Removing stale agent lock %s", self.path)
                    self.path.unlink(missing_ok=True)
                    continue
                if time.monotonic() > deadline:
                    raise AgentBusy("another agent run is still in progress")
                time.sleep(15)

    def __exit__(self, *exc) -> None:
        self.path.unlink(missing_ok=True)
