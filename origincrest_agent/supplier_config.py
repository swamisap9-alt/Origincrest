"""Settings for the supplier agent, kept apart from the outreach agent's config.yaml.

The supplier agent has its own file, config.suppliers.yaml, for everything about the supplier search:
the workbook, the vendor to compare with, the demand to describe, the search rotation and the follow-up
timing. Editing it can't change or break the outreach agent, and editing config.yaml can't change the
supplier search.

The two agents share one mailbox, so the supplier agent *reads* (never writes) four sections of
config.yaml - `email`, `sender`, `claude` and `timezone` - and the secrets in .env. That way drafts go to
the same Drafts folder, with the same From address and signature, and the IMAP details are entered once.
Any of those sections can be overridden in config.suppliers.yaml.

The supplier agent also keeps its own data folder (state database, logs, reports, workbook backups,
run lock): by default a `suppliers` folder inside the outreach agent's data folder.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any

import yaml
from dotenv import load_dotenv

from .config import _ORIGINCREST, DEFAULTS, PROJECT_DIR, Config, ConfigError, _deep_merge

SUPPLIER_CONFIG_NAME = "config.suppliers.yaml"
WORKBOOK = "suppliers"  # internal name of the supplier workbook
# Read from config.yaml (unless config.suppliers.yaml sets them): the shared mailbox and Claude account.
SHARED_SECTIONS = ("timezone", "claude", "email", "sender")
# Sections that belong in config.suppliers.yaml; if they turn up in config.yaml they are ignored.
SUPPLIER_SECTIONS = ("suppliers", "sourcing")
KNOWN_KEYS = {"suppliers", "sourcing", "followup", "data_dir", "shared_config", "excel", "compliance",
              *SHARED_SECTIONS}

SUPPLIER_DEFAULTS: dict[str, Any] = {
    "suppliers": {
        "workbook_path": _ORIGINCREST + r"\OriginCrest suppliers\Indonesian_Coffee_Suppliers.xlsx",
        "tab": "New Suppliers - Indonesia",
        # The vendor every find is compared with. `profile` is what you know about them (origins, grades,
        # prices, minimums, terms, how responsive they are); the comparison is only as good as this.
        # Never used in an email.
        "baseline_vendor": {"name": "", "profile": ""},
        "every_days": 3,                 # how often the scheduled search runs; also rotates the search focus
        "focus_per_run": 2,              # how many regions / supplier types each run searches
        "candidates_per_focus": 6,
        "max_new_rows_per_run": 8,
        "max_drafts_per_run": 8,
        "min_fit_score_for_draft": 50,
        "web_searches_per_focus": 8,
        "web_fetches_per_focus": 8,
        "effort": "high",
        "email_effort": "medium",
        "extra_focus": [],               # more regions / supplier types for the rotation
        # Searches for what you need now (e.g. the coffees in a customer's quote): every run takes the next
        # one of these, plus focus_per_run - 1 from the rotation.
        "priority_focus": [],
    },
    # What the enquiries say about OriginCrest as a buyer. Blank values become [placeholders] in the draft;
    # the agent never invents volumes, prices or dates.
    "sourcing": {
        "intro": "",
        "experience": "",   # OriginCrest's experience, for the pitch (facts only; used word for word)
        "demand": "",
        "logistics": ("Sea freight, customs and biosecurity clearance and warehousing in Australia are "
                      "already arranged."),
        "wanted": ["all Indonesian origins: arabica, robusta and liberica, specialty to commercial grades"],
        "ask": ["your quotation: FOB price per kg at your loading port for the coffees in our current order (or "
                "your closest equivalents), for the trial (LCL) quantity and for a 20 ft FCL",
                "your current offer list with FOB prices for your other coffees (grade, process, moisture, crop)",
                "whether your prices include export documents and charges (PEB, phytosanitary certificate, "
                "fumigation, ICO certificate / certificate of origin)",
                "your minimum order, and whether you can prepare a part-container (LCL) shipment",
                "preparation time after deposit, and your loading port",
                "green samples for approval before shipment, and payment terms"],
    },
    # The supplier agent's own follow-up settings (config.yaml's followup section is not used).
    "followup": {
        "supplier_days": 3,
        "max_drafts_per_run": 15,
        "max_followups_per_contact": 3,
    },
    "statuses": {"supplier_drafted": "Enquiry drafted - to review"},
    # Enquiries and follow-ups carry the same one-line opt-out as RFQs ("reply 'unsubscribe' and we won't
    # contact you again") plus a List-Unsubscribe header; replies saying unsubscribe are honoured.
    "compliance": {"unsubscribe_footer_for": ["lead", "quote", "supplier"]},
}


class SupplierConfig(Config):
    """A Config for the supplier agent, plus where each part of it came from (shown by `check`)."""

    def __init__(self, data: dict[str, Any], source: Path | None, *, shared_source: Path | None,
                 shared_used: list[str], ignored_in_shared: list[str], unknown_keys: list[str],
                 warnings: list[str]):
        super().__init__(data, source)
        self.shared_source = shared_source
        self.shared_used = shared_used
        self.ignored_in_shared = ignored_in_shared
        self.unknown_keys = unknown_keys
        self.warnings = warnings


def _read_yaml(path: Path) -> dict[str, Any]:
    with open(path, encoding="utf-8") as fh:
        try:
            data = yaml.safe_load(fh) or {}
        except yaml.YAMLError as exc:
            raise ConfigError(f"{path} has a layout mistake (check the indentation around the line shown; "
                              f"use spaces, not tabs): {exc}") from exc
    if not isinstance(data, dict):
        raise ConfigError(f"{path} must contain a YAML mapping")
    return data


def supplier_config_path(path: str | os.PathLike | None = None) -> Path:
    return Path(path) if path else Path(os.environ.get("ORIGINCREST_SUPPLIER_CONFIG",
                                                       PROJECT_DIR / SUPPLIER_CONFIG_NAME))


def load_supplier_config(path: str | os.PathLike | None = None,
                         shared_path: str | os.PathLike | None = None) -> SupplierConfig:
    load_dotenv(PROJECT_DIR / ".env")
    own_path = supplier_config_path(path)
    own: dict[str, Any] = {}
    if own_path.exists():
        own = _read_yaml(own_path)
    elif path:
        raise ConfigError(f"Supplier config file not found: {own_path}")

    shared_file = Path(shared_path or own.get("shared_config")
                       or os.environ.get("ORIGINCREST_CONFIG", PROJECT_DIR / "config.yaml"))
    shared_all: dict[str, Any] = {}
    warnings: list[str] = []
    if shared_file.exists():
        try:
            shared_all = _read_yaml(shared_file)
        except ConfigError as exc:
            if "email" not in own:
                raise ConfigError(f"Can't read the mailbox settings: {exc}") from exc
            warnings.append(f"{shared_file.name} could not be read ({exc}); using the email settings in "
                            f"{own_path.name}")
    shared = {k: shared_all[k] for k in SHARED_SECTIONS if k in shared_all}

    data = _deep_merge(_deep_merge(_deep_merge(DEFAULTS, SUPPLIER_DEFAULTS), shared),
                       {k: v for k, v in own.items() if k != "shared_config"})
    if "data_dir" not in own:
        base = str(shared_all.get("data_dir") or DEFAULTS["data_dir"]).rstrip("\\/")
        data["data_dir"] = base + ("\\suppliers" if "\\" in base else "/suppliers")
    sc = data["suppliers"]
    # Only the supplier workbook exists for this agent; the outreach workbooks and tabs are out of reach.
    data["workbooks"] = {WORKBOOK: {"path": sc["workbook_path"], "default_kind": "supplier"}}
    data["discovery"]["profiles"] = {}
    data["leads"]["extra_tabs"] = [sc["tab"]]
    data["followup"]["workbooks"] = [WORKBOOK]
    return SupplierConfig(
        data, own_path if own_path.exists() else None,
        shared_source=shared_file if shared_file.exists() else None,
        shared_used=[k for k in SHARED_SECTIONS if k in shared and k not in own],
        ignored_in_shared=[k for k in SUPPLIER_SECTIONS if k in shared_all],
        unknown_keys=sorted(k for k in own if k not in KNOWN_KEYS),
        warnings=warnings,
    )
