"""Configuration: built-in defaults, overridden by config.yaml, secrets from .env."""

from __future__ import annotations

import copy
import os
from pathlib import Path
from typing import Any

import yaml
from dotenv import load_dotenv

PROJECT_DIR = Path(__file__).resolve().parent.parent

_ORIGINCREST = r"C:\Users\swami\OneDrive\Documents\Origincrest"

DEFAULTS: dict[str, Any] = {
    "timezone": "Australia/Adelaide",
    # Local folder for the state database, logs, workbook backups and run reports.
    # Kept outside OneDrive so the SQLite file is never synced mid-write.
    "data_dir": "%LOCALAPPDATA%/OriginCrestAgent" if os.name == "nt" else "~/.origincrest_agent",
    "workbooks": {
        "rfq": {
            "path": _ORIGINCREST
            + r"\Origincrest 3PL and Freight forwarder contacts RFQ\OriginCrest_Freight_3PL_Contacts_RFQ.xlsx",
            # Tabs not named under `tabs` are treated as this kind of row.
            "default_kind": "freight",
        },
        "master": {
            "path": _ORIGINCREST
            + r"\OriginCrest enquiries\Australian_Coffee_3PL_Master_Contacts_Oct2026.xlsx",
            "default_kind": "freight",
        },
    },
    "tabs": {
        "roasters_australia": "Roasters - Australia",
        "roasters_adelaide": "Roasters - Adelaide",
        "freight_australia": "Freight-3PL - Australia",
        "freight_adelaide": "Freight-3PL - Adelaide",
        # Other existing tabs that hold buyers/roasters (lead rows).
        "extra_roaster_tabs": [],
        # Tabs the agent must never read or write.
        "ignore_tabs": [],
    },
    "statuses": {
        "new": "New - to review",
        "intro_drafted": "Intro drafted - to review",
        "low_fit": "Researched - low fit",
        "no_email": "Researched - no public email",
        "open": "Open",
        "replied": "Replied - review",
        "unsubscribed": "Unsubscribed - do not contact",
        "bounced": "Bounced - check email",
        # Rows with any of these statuses are never researched or emailed.
        "skip": [
            "rejected",
            "do not contact",
            "unsubscribed",
            "unsubscribed - do not contact",
            "duplicate",
            "closed",
            "lost",
            "not a fit",
            "bounced - check email",
        ],
    },
    "claude": {
        "model": "claude-opus-5-5",
        "api_key_env": "ANTHROPIC_API_KEY",
        # Server-side refusal fallback (re-runs a declined request on another model).
        "use_fallbacks": True,
        "web_search_tool": "web_search_20260209",
        "web_fetch_tool": "web_fetch_20260209",
        "max_tokens": 32000,
        "timeout_seconds": 900,
        "user_location": {
            "type": "approximate",
            "city": "Adelaide",
            "region": "South Australia",
            "country": "AU",
            "timezone": "Australia/Adelaide",
        },
    },
    "discovery": {
        "workbooks": ["rfq", "master"],
        "segments": ["roasters_adelaide", "roasters_australia", "freight_adelaide", "freight_australia"],
        "max_new_rows_per_run": 30,
        "max_new_rows_per_segment": 10,
        "candidates_per_segment": 12,
        "web_searches_per_segment": 8,
        "web_fetches_per_segment": 8,
        "effort": "medium",
        # Fetch each source page and keep only contact details that appear on it.
        "verify_contacts": True,
        "run_leads_after": True,
    },
    "leads": {
        "workbooks": ["rfq", "master"],
        "max_research_per_run": 10,
        "max_drafts_per_run": 8,
        "min_fit_score_for_draft": 55,
        # true: draft intros for "New - to review" rows too (drafts still need your approval).
        # false: only draft once you've changed the row's status away from "New - to review".
        "draft_for_unreviewed": True,
        # Only rows with one of these statuses are researched / emailed (blank = a row you added).
        # Anything else (e.g. "Customer", "In discussion") is left alone.
        "eligible_statuses": ["", "new", "new - to review", "reviewed", "approved", "to contact", "lead",
                              "prospect"],
        # Your override: these statuses get an intro draft even if the fit score is low.
        "approved_statuses": ["approved"],
        "research_cache_days": 120,
        "web_searches_per_company": 5,
        "web_fetches_per_company": 5,
        "research_effort": "high",
        "email_effort": "medium",
    },
    "followup": {
        "workbooks": ["rfq"],
        "open_statuses": ["open"],
        "quote_days": 5,
        "lead_days": 10,
        "max_drafts_per_run": 15,
        "max_followups_per_contact": 3,
        "effort": "medium",
        "sent_lookback_days": 180,
    },
    "email": {
        "imap_host": "",
        "imap_port": 993,
        # ssl = implicit TLS (usually port 993), starttls = upgrade on port 143, none = local testing only.
        "security": "ssl",
        "username": "",
        "password_env": "IMAP_PASSWORD",
        # Leave blank to auto-detect from the server's special-use folder flags.
        "drafts_folder": "",
        "sent_folder": "",
        "reply_folders": [],
        "max_fetch_bytes": 5_000_000,
    },
    "sender": {
        "name": "<<EDIT ME: your name>>",
        "title": "",
        "business_name": "OriginCrest",
        "abn": "",
        "address": "<<EDIT ME: business postal address>>",
        "phone": "",
        "email": "<<EDIT ME: sending email address>>",
        "website": "",
        "unsubscribe_email": "",
    },
    "offer": {
        "summary": "<<EDIT ME: one or two sentences on what OriginCrest supplies>>",
        "origins": [
            "Sumatra",
            "Sulawesi",
            "Java",
            "Flores",
        ],
        "facts": [],
        "call_to_action": "Offer to send green samples and a current offer sheet.",
    },
    "compliance": {
        # Record types whose emails get the Spam Act unsubscribe footer.
        "unsubscribe_footer_for": ["lead"],
    },
    "excel": {
        "lock_wait_minutes": 5,
        "backup_keep": 40,
        "date_format": "dd/mm/yyyy",
    },
}


class ConfigError(Exception):
    pass


def _deep_merge(base: dict, override: dict) -> dict:
    out = copy.deepcopy(base)
    for key, value in (override or {}).items():
        if isinstance(value, dict) and isinstance(out.get(key), dict):
            out[key] = _deep_merge(out[key], value)
        else:
            out[key] = value
    return out


def expand_path(value: str) -> Path:
    return Path(os.path.expandvars(os.path.expanduser(value)))


class Config:
    def __init__(self, data: dict[str, Any], source: Path | None = None):
        self.data = data
        self.source = source

    def __getitem__(self, key: str) -> Any:
        return self.data[key]

    @property
    def data_dir(self) -> Path:
        return expand_path(self.data["data_dir"])

    @property
    def tz(self) -> str:
        return self.data["timezone"]

    def workbook_path(self, name: str) -> Path:
        try:
            return expand_path(self.data["workbooks"][name]["path"])
        except KeyError as exc:
            raise ConfigError(f"Unknown workbook '{name}' (check the workbooks: section)") from exc

    def secret(self, env_name: str) -> str:
        return os.environ.get(env_name, "").strip()

    def status(self, key: str) -> str:
        return self.data["statuses"][key]

    @property
    def skip_statuses(self) -> set[str]:
        return {s.strip().lower() for s in self.data["statuses"]["skip"]}


def load_config(path: str | os.PathLike | None = None) -> Config:
    load_dotenv(PROJECT_DIR / ".env")
    cfg_path = Path(path) if path else Path(os.environ.get("ORIGINCREST_CONFIG", PROJECT_DIR / "config.yaml"))
    override: dict[str, Any] = {}
    if cfg_path.exists():
        with open(cfg_path, encoding="utf-8") as fh:
            override = yaml.safe_load(fh) or {}
        if not isinstance(override, dict):
            raise ConfigError(f"{cfg_path} must contain a YAML mapping")
    elif path:
        raise ConfigError(f"Config file not found: {cfg_path}")
    return Config(_deep_merge(DEFAULTS, override), cfg_path if cfg_path.exists() else None)
