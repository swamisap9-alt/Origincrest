"""Workflow 1: find new leads for each workbook (weekly).

* RFQ workbook: freight forwarders, customs brokers and 3PLs.
* Master workbook: the categories listed in its own tabs (e.g. "New Prospects Oct-26", and the
  importer / roaster / wholesaler / trader categories in "All contacts").

New leads go to two tabs per workbook: one for Adelaide / South Australia, one for all other states.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from .dedupe import entity_keys
from .llm import AgentError
from .prompts import DISCOVERY_SYSTEM, RECORD_COMPANIES, discovery_prompt
from .runtime import ChangeSet, RunContext, commit
from .schema import category_kind
from .state import decode_values
from .utils import clean, log
from .workbook import SheetView, WorkbookStore, build_dedupe_index

AGENT_NAME = "OriginCrest Agent"
REGIONS = ("adelaide", "other")

# What each weekly search concentrates on, rotating so repeated runs keep finding new businesses.
_TOPICAL_FREIGHT = {
    "adelaide": (
        "sea-freight import forwarders at Port Adelaide",
        "customs brokers with DAFF biosecurity and import-permit experience",
        "food-grade and bonded warehouses in Adelaide",
        "container unpacking and pallet storage around Port Adelaide, Wingfield and Gillman",
        "3PL fulfilment for food and beverage e-commerce brands in Adelaide",
        "interstate pallet and LTL carriers servicing Adelaide",
    ),
    "other": (
        "forwarders with Indonesia-Australia FCL/LCL services",
        "green-coffee import specialists and coffee logistics providers",
        "customs brokers handling coffee and food imports and biosecurity",
        "food-grade 3PL warehousing in Melbourne, Sydney and Brisbane",
        "fumigation and treatment providers for imported agricultural goods",
        "3PL pick-pack fulfilment for roasted-coffee e-commerce",
        "national pallet freight networks used by coffee businesses",
    ),
}
_PLACES = {
    "adelaide": (
        "Adelaide CBD and inner suburbs",
        "Adelaide Hills",
        "Barossa, Clare Valley and the Fleurieu Peninsula",
        "northern and western suburbs, including Port Adelaide",
        "southern suburbs and McLaren Vale",
        "regional SA (Mount Gambier, Port Lincoln, Riverland, Victor Harbor)",
    ),
    "other": (
        "Melbourne",
        "Sydney",
        "Brisbane",
        "Perth and Fremantle",
        "Hobart and Launceston",
        "Canberra",
        "Gold Coast and Sunshine Coast",
        "Newcastle and the Hunter",
        "Geelong and regional Victoria",
        "regional New South Wales",
        "regional Queensland",
        "Darwin",
    ),
}


@dataclass(frozen=True)
class Category:
    label: str
    kind: str  # "buyer" | "freight"
    examples: tuple[str, ...] = ()
    source: str = ""  # "config" or the tab it was read from

    @property
    def topical(self) -> bool:
        """The RFQ workbook's built-in freight category rotates through service themes, not places."""
        return self.kind == "freight" and self.source == "config"


@dataclass(frozen=True)
class Segment:
    workbook: str
    category: Category
    region: str

    def focus(self, week: int) -> list[str]:
        pool = (_TOPICAL_FREIGHT if self.category.topical else _PLACES)[self.region]
        per_week = 3 if (self.region == "other" and not self.category.topical) else 2
        start = (week * per_week) % len(pool)
        return [pool[(start + i) % len(pool)] for i in range(min(per_week, len(pool)))]


def _norm(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(text).lower())


def find_tab(store: WorkbookStore, name: str) -> SheetView | None:
    """Tab by name, ignoring case, spaces and punctuation ("New Prospects Oct-26" ~ "New prospects oct 26")."""
    target = _norm(name)
    titles = store.wb.sheetnames
    for title in titles:
        if _norm(title) == target:
            return store.view(title)
    for title in titles:
        if target and (target in _norm(title) or _norm(title) in target):
            return store.view(title)
    return None


def _split(label: str) -> list[str]:
    """'Importers, Roasters, Wholesalers & Traders' -> four searchable categories."""
    if "," not in label:
        return [label]
    return [p.strip() for p in re.split(r",|&|\band\b", label) if p.strip()]


def profile_categories(cfg, store: WorkbookStore, notes: list[str]) -> list[Category]:
    profile = cfg.profile(store.name) or {}
    found: dict[str, Category] = {}

    def add(label: str, source: str, examples: list[str] | None = None) -> None:
        label = clean(label)
        if label and label.lower() not in found:
            found[label.lower()] = Category(label, category_kind(label), tuple(examples or [])[:5], source)

    for label in profile.get("categories") or []:
        add(label, "config")
    for spec in profile.get("category_tabs") or []:
        spec = {"tab": spec} if isinstance(spec, str) else spec
        view = find_tab(store, spec["tab"])
        labels: dict[str, list[str]] = {}
        if view is None:
            notes.append(f"{store.path.name}: tab '{spec['tab']}' not found")
        elif not view.has("category"):
            notes.append(f"{store.path.name} > {view.title}: no Category column found")
        else:
            for row in view.rows():
                if label := row.text("category"):
                    labels.setdefault(label, []).append(row.text("company"))
        match = [w.lower() for w in spec.get("match") or []]
        before = len(found)
        for label, companies in labels.items():
            if match and not any(w in label.lower() for w in match):
                continue
            for part in _split(label):
                add(part, view.title, [c for c in companies if c])
        if len(found) == before and spec.get("fallback"):
            for part in _split(spec["fallback"]):
                add(part, "config fallback")
    exclude = {e.strip().lower() for e in profile.get("exclude_categories") or []}
    return [c for c in found.values() if c.label.lower() not in exclude]


def plan_segments(cfg, stores: dict[str, WorkbookStore], week: int, notes: list[str],
                  categories: list[str] | None = None, regions: list[str] | None = None) -> list[Segment]:
    per_run = cfg["discovery"]["categories_per_run"]
    wanted = {c.strip().lower() for c in categories or []}
    segments: list[Segment] = []
    for name, store in stores.items():
        cats = profile_categories(cfg, store, notes)
        notes.append(f"{store.path.name}: categories " + (
            "; ".join(f"{c.label} ({'RFQ' if c.kind == 'freight' else 'intro'})" for c in cats) or "(none found)"))
        if wanted:
            cats = [c for c in cats if c.label.lower() in wanted]
        elif len(cats) > per_run:
            start = (week * per_run) % len(cats)
            cats = [cats[(start + i) % len(cats)] for i in range(per_run)]
        for cat in cats:
            for region in regions or REGIONS:
                segments.append(Segment(name, cat, region))
    return segments


def _website(url: str) -> str:
    url = clean(url)
    if url and not url.lower().startswith(("http://", "https://")):
        url = "https://" + url
    return url


def run_discover(ctx: RunContext, workbooks: list[str] | None = None, categories: list[str] | None = None,
                 regions: list[str] | None = None, max_rows: int | None = None) -> int:
    cfg, report = ctx.cfg, ctx.report
    d = cfg["discovery"]
    names = workbooks or [n for n in d["profiles"] if n in cfg["workbooks"]]
    stores = {n: WorkbookStore(n, cfg.workbook_path(n), cfg).load() for n in names}
    week = ctx.today.isocalendar()[1]
    notes: list[str] = []
    segments = plan_segments(cfg, stores, week, notes, categories, regions)
    for note in notes:
        report.add("Categories", note)

    indexes = {}
    for name, store in stores.items():
        index = build_dedupe_index({name: store})
        for keys in ctx.state.discovered_entities(name):
            index.add_keys("an earlier discovery run", keys)
        for item in ctx.state.pending_rows(name):
            v = decode_values(item["values_json"])
            index.add("rows waiting to be written", v.get("company", ""), v.get("website", ""), v.get("email", ""),
                      v.get("phone", ""))
        indexes[name] = index
        log.info("%s: duplicate check covers %d existing companies", store.path.name, len(index))

    cap = max_rows if max_rows is not None else d["max_new_rows_per_workbook"]
    added = {name: 0 for name in stores}
    for seg in segments:
        store, index = stores[seg.workbook], indexes[seg.workbook]
        profile = cfg.profile(seg.workbook)
        if added[seg.workbook] >= cap:
            report.add("Caps", f"{store.path.name}: cap of {cap} new rows reached; skipped "
                               f"{seg.category.label} ({seg.region})")
            continue
        seg_cap = min(d["max_new_rows_per_segment"], cap - added[seg.workbook])
        focus = seg.focus(week)
        log.info("Searching %s: %s, %s (focus: %s)", store.path.name, seg.category.label, seg.region, "; ".join(focus))
        try:
            result = ctx.agent.research(
                system=DISCOVERY_SYSTEM,
                prompt=discovery_prompt(kind=seg.category.kind, category=seg.category.label, region=seg.region,
                                        focus=focus, n=d["candidates_per_segment"], known=index.known_names(250),
                                        examples=list(seg.category.examples)),
                record_tool=RECORD_COMPANIES,
                effort=d["effort"],
                web_searches=d["web_searches_per_segment"],
                web_fetches=d["web_fetches_per_segment"],
            )
        except AgentError as exc:
            report.add("Errors", f"{seg.category.label} ({seg.region}): {exc}")
            continue

        candidates = result.data.get("companies", [])
        seg_added = 0
        for i, cand in enumerate(candidates):
            if seg_added >= seg_cap:
                report.add("Caps", f"{seg.category.label} ({seg.region}): cap of {seg_cap} reached; "
                                   f"{len(candidates) - i} candidate(s) left out")
                break
            company = clean(cand.get("company_name"))
            if not company:
                continue
            website = _website(cand.get("website", ""))
            email, phone = clean(cand.get("email")), clean(cand.get("phone"))
            if dup := index.match(company, website, email, phone):
                report.add("Skipped duplicates", f"{company} - already in {dup}")
                report.count("duplicates skipped")
                continue

            source_url = clean(cand.get("source_url"))
            check = ctx.verifier.verify(email=email, phone=phone, contact_name=clean(cand.get("contact_name")),
                                        source_url=source_url, website=website, evidence_pages=result.pages)
            state_code = clean(cand.get("state"))
            adelaide = bool(cand.get("adelaide_presence")) or state_code == "SA"
            tab = profile["adelaide_tab"] if adelaide else profile["other_tab"]
            kind = seg.category.kind
            business_type = clean(cand.get("business_type"))
            category = (business_type or seg.category.label) if seg.category.topical else seg.category.label
            services = " - ".join(x for x in (clean(cand.get("services")), clean(cand.get("coffee_evidence"))) if x)
            if business_type and not seg.category.topical:
                services = f"{business_type}: {services}" if services else business_type
            notes_out = check.notes + ([clean(cand["notes"])] if clean(cand.get("notes")) else [])
            values = {
                "company": company,
                "category": category,
                "contact_name": check.contact_name,
                "contact_role": clean(cand.get("contact_role")) if check.contact_name else "",
                "email": check.email,
                "phone": check.phone,
                "website": website,
                "city": clean(cand.get("city")),
                "state": state_code,
                "services": services,
                "source_url": source_url,
                "status": cfg.status("new"),
                "date_added": ctx.today,
                "added_by": AGENT_NAME,
                "record_type": "Lead" if kind == "buyer" else "Quote",
                "agent_notes": "; ".join(notes_out),
            }
            index.add(f"this run ({tab})", company, website, check.email, check.phone)
            line = (f"{company} [{category}] -> {store.path.name} > {tab} ({values['city'] or '?'} {state_code}) "
                    f"{check.email or 'no verified email'}")
            if ctx.dry_run:
                report.add("Would add", line)
            else:
                ctx.state.queue_row(seg.workbook, tab, store.default_kind, values)
                ctx.state.add_discovered(seg.workbook, company, kind, entity_keys(company, website, check.email,
                                                                                  check.phone),
                                         source_url, ctx.today)
                report.add("Added (status: " + cfg.status("new") + ")", line)
            report.count(f"new rows - {store.path.name} > {tab}")
            seg_added += 1
            added[seg.workbook] += 1

    commit(ctx, ChangeSet())
    return sum(added.values())
