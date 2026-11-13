"""Supplier discovery: Indonesian green-coffee exporters, cooperatives and mills (every 3 days).

Each run searches a couple of growing regions or supplier types (rotating, so repeated runs keep
finding new ones), compares every find with the current vendor, and adds it to the supplier
workbook with a fit score, the comparison and a suggested way to make contact. A sourcing-enquiry
draft is then saved for each new supplier (see `LeadRun.supplier`), and the follow-up workflow
chases unanswered enquiries after 3 days.
"""

from __future__ import annotations

from pathlib import Path

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from .dedupe import DedupeIndex, entity_keys, norm_phone
from .llm import AgentError
from .prompts import RECORD_SUPPLIERS, SUPPLIER_SYSTEM, _filled, supplier_discovery_prompt
from .runtime import ChangeSet, RunContext, commit
from .schema import FIELD_BY_KEY, NEW_TAB_KEYS
from .state import decode_values
from .utils import clean, log
from .workbook import HEADER_FILL, WorkbookStore, build_dedupe_index

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


def ensure_workbook(cfg) -> Path:
    """Create the supplier workbook (one tab, headers only) if it doesn't exist yet."""
    name = cfg.supplier_workbook
    path = cfg.workbook_path(name)
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
    cfg, report = ctx.cfg, ctx.report
    sc = cfg["suppliers"]
    name = cfg.supplier_workbook
    if not name:
        report.add("Errors", f"suppliers.workbook '{sc.get('workbook')}' is not listed under workbooks: in config.yaml")
        return 0
    path = cfg.workbook_path(name)
    if not path.exists() and not ctx.dry_run:
        ensure_workbook(cfg)
        report.add("Supplier workbook", f"Created {path}")

    vendor = _filled(sc["baseline_vendor"].get("name"))
    index = DedupeIndex()
    if path.exists():
        index = build_dedupe_index({name: WorkbookStore(name, path, cfg).load()})
    for keys in ctx.state.discovered_entities(name):
        index.add_keys("an earlier supplier search", keys)
    for item in ctx.state.pending_rows(name):
        v = decode_values(item["values_json"])
        index.add("rows waiting to be written", v.get("company", ""), v.get("website", ""), v.get("email", ""),
                  v.get("phone", ""))
    if vendor:
        index.add("your current vendor", vendor)
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
            if dup := index.match(company, website, email, phone or whatsapp):
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
            index.add(f"this run ({tab})", company, website, check.email, check.phone or whatsapp)
            line = (f"{company} [{values['category']}, {values['city'] or '?'}] fit {score}/100 - "
                    f"{check.email or 'no verified email'}{' - ' + values['phone'] if values['phone'] else ''}")
            if ctx.dry_run:
                report.add("Would add", line)
                if values["comparison"]:
                    report.add(f"Comparison with {vendor or 'current vendor'}", f"{company}: {values['comparison']}")
            else:
                ctx.state.queue_row(name, tab, "supplier", values)
                ctx.state.add_discovered(name, company, "supplier",
                                         entity_keys(company, website, check.email, check.phone or whatsapp),
                                         source_url, ctx.today)
                report.add("Added (status: " + cfg.status("new") + ")", line)
                report.add(f"Comparison with {vendor or 'current vendor'}", f"{company}: {values['comparison']}")
                report.add("Suggested approach", f"{company}: {values['approach']}")
            report.count("new suppliers")
            added += 1

    commit(ctx, ChangeSet())
    return added
