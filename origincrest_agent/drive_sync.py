"""Daily one-way backup: copy new or changed files from the OneDrive Origincrest folder to Google Drive.

The Google Drive side is Google Drive for desktop: files copied into its folder on this PC are
uploaded by Google, and overwriting a file there adds a new version in Drive (history is kept).
Nothing is ever deleted, in Drive or on the laptop; files that disappear locally are only reported.
"""

from __future__ import annotations

import hashlib
import json
import os
import shutil
from fnmatch import fnmatch
from pathlib import Path

from .config import expand_path
from .runtime import RunContext
from .utils import log


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _skipped(rel: Path, patterns: list[str]) -> bool:
    return any(fnmatch(part.lower(), pat.lower()) for part in rel.parts for pat in patterns)


def _load(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def _save(path: Path, manifest: dict) -> None:
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(manifest, indent=1, sort_keys=True), encoding="utf-8")
    os.replace(tmp, path)


def run_drive_sync(ctx: RunContext) -> int:
    cfg, report = ctx.cfg["drive_sync"], ctx.report
    source, dest = expand_path(cfg["source"]), expand_path(cfg["destination"])
    if not source.is_dir():
        raise RuntimeError(f"OneDrive folder not found: {source} (drive_sync.source in config.yaml)")
    if not dest.is_dir():
        raise RuntimeError(
            f"Google Drive folder not found: {dest}. Is Google Drive for desktop installed, running and "
            "signed in? Set drive_sync.destination in config.yaml to the folder as it appears in File Explorer.")
    manifest_path = ctx.cfg.data_dir / "drive_sync_manifest.json"
    manifest = _load(manifest_path)
    max_bytes = int(cfg["max_file_mb"]) * 1024 * 1024
    seen: set[str] = set()
    copied = 0

    for path in sorted(source.rglob("*")):
        rel = path.relative_to(source)
        if not path.is_file() or _skipped(rel, cfg["skip"]):
            continue
        key = rel.as_posix()
        seen.add(key)
        target = dest / rel
        try:
            st = path.stat()
            entry = manifest.get(key)
            if entry and entry["size"] == st.st_size and entry["mtime"] == st.st_mtime_ns and target.exists():
                continue  # unchanged since the last sync
            if st.st_size > max_bytes:
                report.add("Skipped (too large)", f"{key} ({st.st_size / 1e6:.0f} MB) - copy it into Drive yourself")
                continue
            digest = _sha256(path)
            if entry and entry["sha256"] == digest and target.exists():
                entry["mtime"] = st.st_mtime_ns  # touched but not changed
                continue
            if not entry and target.exists() and target.stat().st_size == st.st_size:
                # Already in Drive (e.g. uploaded by the earlier sync): adopt it without copying.
                manifest[key] = {"sha256": digest, "size": st.st_size, "mtime": st.st_mtime_ns}
                continue
            action = "Updated in Drive (new version)" if target.exists() else "Added to Drive"
            if ctx.dry_run:
                report.add("Would copy", f"{key} -> {target}")
                continue
            if not target.parent.exists():
                target.parent.mkdir(parents=True)
                report.add("New Drive folders", str(target.parent.relative_to(dest)))
            shutil.copyfile(path, target)
            try:
                shutil.copystat(path, target)
            except OSError:
                pass  # timestamps are a nicety; some virtual drives refuse them
            manifest[key] = {"sha256": digest, "size": st.st_size, "mtime": st.st_mtime_ns}
            report.add(action, key)
            copied += 1
        except OSError as exc:
            report.add("Errors", f"{key}: {exc} - will retry next run")

    for key in sorted(set(manifest) - seen):
        report.add("Missing locally (kept in Drive)", key)
        if not ctx.dry_run:
            del manifest[key]  # reported once
    if not ctx.dry_run:
        _save(manifest_path, manifest)
    report.count("files copied to Google Drive", copied)
    if not copied and "Would copy" not in report.sections:
        report.add("Drive sync", "No changes", quiet=True)
    (log.info if copied else log.debug)("Drive sync: %d file(s) copied to %s", copied, dest)
    return copied
