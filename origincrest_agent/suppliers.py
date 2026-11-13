"""The supplier agent's workflows: Indonesian green-coffee exporters, cooperatives and mills.

* search (every 3 days): a couple of growing regions or supplier types per run (rotating, so repeated
  runs keep finding new ones, plus any `priority_focus`). Every find is compared with the current
  vendor and added to the supplier workbook with a fit score, the comparison and a suggested approach.
* drafts: a sourcing enquiry for each new supplier, saved to the shared mailbox's Drafts folder.
* follow-up (weekdays): a threaded follow-up for enquiries with no reply after `supplier_days`.

Runs with its own settings (config.suppliers.yaml, see supplier_config.py), its own workbook and its own
data folder, so nothing here touches the outreach agent's workbooks, state or settings.
"""

from __future__ import annotations

from pathlib import Path

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

import re

from .dedupe import DedupeIndex, entity_keys, norm_phone
from .followup import run_followup
from .leads import LeadRun, _candidates, _first_email
from .llm import AgentError
from .prompts import EMAIL_SCHEMA, _filled
from .runtime import ChangeSet, RunContext, commit
from .schema import FIELD_BY_KEY, NEW_TAB_KEYS, SUPPLIER_PRIVATE
from .state import decode_values
from .supplier_config import WORKBOOK
from .supplier_prompts import (RECORD_SUPPLIERS, SUPPLIER_INTRO_SYSTEM, SUPPLIER_SYSTEM, supplier_discovery_prompt,
                               supplier_intro_prompt)
from .utils import clean, log, to_int
from .workbook import HEADER_FILL, Row, WorkbookStore

AGENT_NAME = "OriginCrest Agent"

# What each run concentrates on. Two per run by default, so the whole list comes round about every month.
FOCUS = (
    "Aceh Gayo arabica: cooperatives and exporters in Takengon, Bener Meriah and Gayo Lues",
    "North Sumatra arabica: Mandheling, Lintong, Sidikalang and Simalungun exporters shipping from Medan / Belawan",
    "Sulawesi arabica: Toraja, Enrekang and Mamasa exporters and mills shipping from Makassar",
    "Java arabica: West Java (Preanger, Garut, Pangalengan) and East Java (Ijen, Bondowoso) exporters shipping "
    "from Jakarta or Surabaya",
    "Flores (Bajawa, Manggarai) and Bali Kintamani cooperatives and exporters",
    "Producer cooperatives with their own export licence and Organic, Fairtrade or Rainforest Alliance certification",
    "Established Indonesian export houses and members of the exporter associations AEKI and GAEKI",
    "Indonesian exhibitors of green coffee at trade shows: Trade Expo Indonesia, World of Coffee Jakarta, "
    "MICE (Melbourne International Coffee Expo), SCAJ",
    "Kerinci (Jambi), Solok (West Sumatra) and Bengkulu specialty arabica exporters",
    "Indonesian exporters that already ship green coffee to Australia or New Zealand, and those that offer "
    "part-container (LCL) or mixed-lot shipments for small importers",
    "Fine robusta and specialty-grade robusta exporters: Lampung, South Sumatra, Temanggung and Flores",
    "Papua (Wamena, Baliem), Sumbawa and other less-known origins with licensed exporters",
)


def run_focus(cfg, today) -> list[str]:
    """This run's search themes (they move on every `every_days` days).

    With `priority_focus` set, each run searches the next priority theme first, then fills the rest of
    `focus_per_run` from the general rotation, so urgent needs are covered without stopping discovery."""
    sc = cfg["suppliers"]
    pool = list(FOCUS) + [f for f in (_filled(x) for x in sc.get("extra_focus") or []) if f]
    priority = [f for f in (_filled(x) for x in sc.get("priority_focus") or []) if f]
    per_run = max(1, int(sc["focus_per_run"]))
    run_number = today.toordinal() // max(1, int(sc["every_days"]))
    picked = [priority[run_number % len(priority)]] if priority else []
    rest = min(per_run - len(picked), len(pool))
    start = (run_number * max(rest, 1)) % len(pool)
    return picked + [pool[(start + i) % len(pool)] for i in range(rest)]


# Indonesian company forms, ignored when checking for duplicates ("PT. Kopi Nusantara" = "Kopi Nusantara").
_LEGAL_PREFIX = re.compile(r"^\s*(?:pt|cv|ud|fa)\b\.?\s+", re.I)
_LEGAL_SUFFIX = re.compile(r"[\s,]+(?:tbk|persero)\.?\s*$", re.I)


def dedupe_name(company: str) -> str:
    name = _LEGAL_SUFFIX.sub("", _LEGAL_PREFIX.sub("", clean(company)))
    return name or clean(company)


def supplier_index(store: WorkbookStore | None) -> DedupeIndex:
    """Every supplier already in the workbook, by name (without PT / CV), website, email and phone."""
    index = DedupeIndex()
    for view in store.views() if store else []:
        for row in view.rows():
            company = row.text("company")
            label = f"{store.path.name} > {view.title} row {row.index}" + (f" ({company})" if company else "")
            index.add(label, company=dedupe_name(company), website=row.text("website"), email=row.text("email"),
                      phone=row.text("phone"), extra_text=row.all_text())
    return index


def ensure_workbook(cfg) -> Path:
    """Create the supplier workbook (one tab, headers only) if it doesn't exist yet."""
    path = cfg.workbook_path(WORKBOOK)
    if path.exists():
        return path
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = cfg["suppliers"]["tab"]
    keys = NEW_TAB_KEYS["supplier"]
    for c, key in enumerate(keys, start=1):
        field = FIELD_BY_KEY[key]
        cell = ws.cell(1, c, field.header)
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        ws.column_dimensions[get_column_letter(c)].width = field.width
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(keys))}1"
    wb.save(path)
    log.info("Created the supplier workbook: %s", path)
    return path


def _phones(phone: str, whatsapp: str) -> str:
    if whatsapp and norm_phone(whatsapp) == norm_phone(phone):
        return f"{phone} (also WhatsApp)"
    return " / ".join(x for x in (phone, f"WhatsApp {whatsapp}" if whatsapp else "") if x)


def run_suppliers(ctx: RunContext, max_rows: int | None = None, focus: list[str] | None = None) -> int:
    """Search for new suppliers and add them to the supplier workbook. Returns how many were added."""
    cfg, report = ctx.cfg, ctx.report
    sc = cfg["suppliers"]
    name = WORKBOOK
    path = cfg.workbook_path(name)
    if not path.exists() and not ctx.dry_run:
        ensure_workbook(cfg)
        report.add("Supplier workbook", f"Created {path}")

    vendor = _filled(sc["baseline_vendor"].get("name"))
    index = supplier_index(WorkbookStore(name, path, cfg).load() if path.exists() else None)
    for keys in ctx.state.discovered_entities(name):
        index.add_keys("an earlier supplier search", keys)
    for item in ctx.state.pending_rows(name):
        v = decode_values(item["values_json"])
        index.add("rows waiting to be written", dedupe_name(v.get("company", "")), v.get("website", ""),
                  v.get("email", ""), v.get("phone", ""))
    if vendor:
        index.add("your current vendor", dedupe_name(vendor))
    if not _filled(sc["baseline_vendor"].get("profile")):
        report.add("Comparison", f"suppliers.baseline_vendor.profile is blank, so comparisons with {vendor or 'your vendor'} "
                                 "only say what each new supplier shows and what to ask - add what you know "
                                 "(origins, grades, prices, minimums, terms) to config.yaml for a real comparison")
    if not _filled(cfg["sourcing"].get("demand")):
        report.add("Demand", "sourcing.demand is blank, so enquiry drafts carry [placeholders] for the customer and "
                             "the LCL / FCL quantities - fill it in in config.yaml")

    cap = sc["max_new_rows_per_run"] if max_rows is None else max_rows
    themes = focus or run_focus(cfg, ctx.today)
    tab = sc["tab"]
    added = 0
    for theme in themes:
        if added >= cap:
            report.add("Caps", f"Cap of {cap} new suppliers reached; not searched: {theme}")
            continue
        log.info("Searching Indonesian suppliers: %s", theme)
        try:
            result = ctx.agent.research(
                system=SUPPLIER_SYSTEM,
                prompt=supplier_discovery_prompt(focus=[theme], n=sc["candidates_per_focus"],
                                                 known=index.known_names(250), cfg=cfg),
                record_tool=RECORD_SUPPLIERS, effort=sc["effort"],
                web_searches=sc["web_searches_per_focus"], web_fetches=sc["web_fetches_per_focus"])
        except AgentError as exc:
            report.add("Errors", f"{theme}: {exc}")
            continue
        candidates = sorted(result.data.get("suppliers", []), key=lambda c: -int(c.get("fit_score") or 0))
        for i, cand in enumerate(candidates):
            company = clean(cand.get("company_name"))
            if not company:
                continue
            if added >= cap:
                report.add("Caps", f"Cap of {cap} new suppliers reached; {len(candidates) - i} left out ({theme})")
                break
            website = clean(cand.get("website"))
            if website and not website.lower().startswith(("http://", "https://")):
                website = "https://" + website
            email, phone, whatsapp = (clean(cand.get(k)) for k in ("email", "phone", "whatsapp"))
            if dup := index.match(dedupe_name(company), website, email, phone or whatsapp):
                report.add("Skipped duplicates", f"{company} - already in {dup}")
                report.count("duplicates skipped")
                continue
            source_url = clean(cand.get("source_url"))
            check = ctx.verifier.verify(email=email, phone=phone, contact_name=clean(cand.get("contact_name")),
                                        source_url=source_url, website=website, evidence_pages=result.pages)
            notes = list(check.notes)
            if whatsapp:
                wa = ctx.verifier.verify(email="", phone=whatsapp, contact_name="", source_url=source_url,
                                         website=website, evidence_pages=result.pages)
                whatsapp = wa.phone
                notes += [n.replace("phone", "WhatsApp number") for n in wa.notes]
            if clean(cand.get("notes")):
                notes.append(clean(cand["notes"]))
            score = max(0, min(100, int(cand.get("fit_score") or 0)))
            values = {
                "company": company,
                "category": clean(cand.get("business_type")) or "Exporter",
                "contact_name": check.contact_name,
                "contact_role": clean(cand.get("contact_role")) if check.contact_name else "",
                "email": check.email,
                "phone": _phones(check.phone, whatsapp),
                "website": website,
                "city": clean(cand.get("city")),
                "state": clean(cand.get("province")),
                "origins": clean(cand.get("origins")),
                "certifications": clean(cand.get("certifications")),
                "export_terms": clean(cand.get("export_capability")),
                "fit_score": score,
                "comparison": clean(cand.get("comparison_with_current_vendor")),
                "fit_rationale": clean(cand.get("fit_rationale")),
                "approach": clean(cand.get("suggested_approach")),
                "source_url": source_url,
                "status": cfg.status("new"),
                "date_added": ctx.today,
                "added_by": AGENT_NAME,
                "record_type": "Supplier",
                "agent_notes": "; ".join(notes),
            }
            index.add(f"this run ({tab})", dedupe_name(company), website, check.email, check.phone or whatsapp)
            line = (f"{company} [{values['category']}, {values['city'] or '?'}] fit {score}/100 - "
                    f"{check.email or 'no verified email'}{' - ' + values['phone'] if values['phone'] else ''}")
            if ctx.dry_run:
                report.add("Would add", line)
                if values["comparison"]:
                    report.add(f"Comparison with {vendor or 'current vendor'}", f"{company}: {values['comparison']}")
            else:
                ctx.state.queue_row(name, tab, "supplier", values)
                ctx.state.add_discovered(name, company, "supplier",
                                         entity_keys(dedupe_name(company), website, check.email,
                                                     check.phone or whatsapp),
                                         source_url, ctx.today)
                report.add("Added (status: " + cfg.status("new") + ")", line)
                report.add(f"Comparison with {vendor or 'current vendor'}", f"{company}: {values['comparison']}")
                report.add("Suggested approach", f"{company}: {values['approach']}")
            report.count("new suppliers")
            added += 1

    commit(ctx, ChangeSet())
    return added


class SupplierRun(LeadRun):
    """Drafts sourcing enquiries to suppliers, through the same checks and Drafts folder as intros and RFQs."""

    def __init__(self, ctx: RunContext, stores: dict[str, WorkbookStore], max_drafts: int | None,
                 redraft: bool = False):
        super().__init__(ctx, stores, max_drafts, redraft=redraft)
        if max_drafts is None:
            self.draft_budget = ctx.cfg["suppliers"]["max_drafts_per_run"]

    def supplier(self, key: str, rows: list[Row]) -> None:
        cfg, row = self.cfg, rows[0]
        sc = cfg["suppliers"]
        email = next((e for r in rows if (e := _first_email(r.text("email")))), "")
        contact = row.text("contact_name")
        score_text = row.text("fit_score")
        score = to_int(row.get("fit_score"))
        where = ", ".join(x for x in (row.text("city"), row.text("state")) if x) or row.sheet.title
        line = f"{row.text('company')} ({where})" + (f": fit {score}/100" if score_text else "")
        if (email and score_text and score < sc["min_fit_score_for_draft"] and not self.redraft
                and row.text("status").lower() not in self.approved):
            self.report.add("Low fit (no draft)", line)
            self._finish(rows, {}, [], f"Fit {score} is below {sc['min_fit_score_for_draft']} - no draft. "
                                       "Set status to Approved to draft anyway.", cfg.status("low_fit"))
            return
        # The email writer sees the supplier's row without your private assessment (comparison, approach).
        self._draft(key, rows, {"record_type": "Supplier"}, [], email=email, contact=contact, kind="supplier",
                    line=line, compose=lambda: self.ctx.agent.structured(
                        system=SUPPLIER_INTRO_SYSTEM,
                        prompt=supplier_intro_prompt(row.email_context(skip=SUPPLIER_PRIVATE), contact, cfg),
                        schema=EMAIL_SCHEMA, effort=sc["email_effort"]))


def _store(ctx: RunContext) -> WorkbookStore | None:
    path = ctx.cfg.workbook_path(WORKBOOK)
    return WorkbookStore(WORKBOOK, path, ctx.cfg).load() if path.exists() else None


def run_supplier_drafts(ctx: RunContext, max_drafts: int | None = None, redraft: list[str] | None = None) -> int:
    """Enquiry drafts for suppliers in the workbook that don't have one yet (or, with `redraft`, for those
    companies again). Returns how many drafts were saved."""
    store = _store(ctx)
    if store is None:
        ctx.report.add("Nothing to draft", f"No supplier workbook yet ({ctx.cfg.workbook_path(WORKBOOK)})")
        return 0
    stores = {WORKBOOK: store}
    groups = _candidates(ctx, stores, redraft, only_kind="supplier")
    if redraft is not None:
        found = [dedupe_name(r.text("company")).lower() for rows in groups.values() for r in rows]
        for name in redraft:
            if not any(dedupe_name(name).lower() in f for f in found):
                ctx.report.add("Not found", f"No row for '{name}' in the supplier workbook")
    log.info("%d supplier(s) to draft", len(groups))
    run = SupplierRun(ctx, stores, max_drafts, redraft=redraft is not None)
    run.changes.touch(WORKBOOK)  # also writes any rows still queued from the search
    try:
        for key, rows in groups.items():
            run.supplier(key, rows)
    finally:
        if run.mailbox is not None:
            run.mailbox.close()
    commit(ctx, run.changes)
    return run.drafted


def run_supplier_followup(ctx: RunContext, max_drafts: int | None = None) -> int:
    """Follow-ups for enquiries with no reply after `followup.supplier_days`. Returns drafts saved."""
    if not ctx.cfg.workbook_path(WORKBOOK).exists():
        ctx.report.add("Nothing to follow up", "No supplier workbook yet - run a search first")
        return 0
    return run_followup(ctx, max_drafts=max_drafts)
