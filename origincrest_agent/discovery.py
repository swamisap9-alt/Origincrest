"""Workflow 1: discover new roasters and freight forwarders / 3PLs (weekly)."""

from __future__ import annotations

from dataclasses import dataclass

from .dedupe import entity_keys
from .llm import AgentError
from .prompts import DISCOVERY_SYSTEM, RECORD_COMPANIES, discovery_prompt
from .runtime import ChangeSet, RunContext, commit
from .utils import clean, log
from .workbook import WorkbookStore, build_dedupe_index

AGENT_NAME = "OriginCrest Agent"


@dataclass(frozen=True)
class Segment:
    kind: str  # "roaster" | "freight"
    region: str  # "adelaide" | "australia"
    focus_pool: tuple[str, ...]
    focus_per_week: int = 2

    def focus_for(self, week: int) -> list[str]:
        n = len(self.focus_pool)
        start = (week * self.focus_per_week) % n
        return [self.focus_pool[(start + i) % n] for i in range(min(self.focus_per_week, n))]


SEGMENTS: dict[str, Segment] = {
    "roasters_adelaide": Segment("roaster", "adelaide", (
        "Adelaide CBD and inner suburbs",
        "Adelaide Hills",
        "Barossa, Clare Valley and the Fleurieu Peninsula",
        "northern and western suburbs, including Port Adelaide",
        "southern suburbs and McLaren Vale",
        "regional SA (Mount Gambier, Port Lincoln, Riverland, Victor Harbor)",
        "wholesale roasters supplying South Australian cafés",
        "small-batch and micro roasters selling online",
    )),
    "roasters_australia": Segment("roaster", "australia", (
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
        "regional Queensland (Toowoomba, Cairns, Townsville)",
        "Darwin and the Northern Territory",
        "Byron Bay and the Northern Rivers",
        "roasters with recent Australian coffee-award results",
        "roasters that feature Asia-Pacific single origins",
    ), focus_per_week=3),
    "freight_adelaide": Segment("freight", "adelaide", (
        "sea-freight import forwarders at Port Adelaide",
        "customs brokers with DAFF biosecurity and import-permit experience",
        "food-grade and bonded warehouses in Adelaide",
        "container unpacking and pallet storage around Port Adelaide, Wingfield and Gillman",
        "3PL fulfilment for food and beverage e-commerce brands in Adelaide",
        "interstate pallet and LTL carriers servicing Adelaide",
    )),
    "freight_australia": Segment("freight", "australia", (
        "forwarders with Indonesia-Australia FCL/LCL services",
        "green-coffee import specialists and coffee logistics providers",
        "customs brokers handling coffee and food imports and biosecurity",
        "food-grade 3PL warehousing in Melbourne, Sydney and Brisbane",
        "fumigation and treatment providers for imported agricultural goods",
        "3PL pick-pack fulfilment for roasted-coffee e-commerce",
        "national pallet freight networks used by coffee roasters",
    )),
}


def _tab_for(cfg, kind: str, adelaide: bool) -> str:
    prefix = "roasters" if kind == "roaster" else "freight"
    return cfg["tabs"][f"{prefix}_{'adelaide' if adelaide else 'australia'}"]


def _website(url: str) -> str:
    url = clean(url)
    if url and not url.lower().startswith(("http://", "https://")):
        url = "https://" + url
    return url


def run_discover(ctx: RunContext, segments: list[str] | None = None, max_rows: int | None = None) -> int:
    cfg, report = ctx.cfg, ctx.report
    d = cfg["discovery"]
    targets = list(d["workbooks"])
    stores: dict[str, WorkbookStore] = {}
    for name in cfg["workbooks"]:
        try:
            stores[name] = WorkbookStore(name, cfg.workbook_path(name), cfg).load()
        except FileNotFoundError as exc:
            if name in targets:
                raise
            log.warning("%s (only used for duplicate checks)", exc)

    index = build_dedupe_index(stores)
    index.add_keys("an earlier discovery run", ctx.state.discovered_keys())
    log.info("Duplicate check covers %d existing companies", len(index))

    total_cap = max_rows if max_rows is not None else d["max_new_rows_per_run"]
    week = ctx.today.isocalendar()[1]
    added = 0
    for seg_id in segments or d["segments"]:
        if added >= total_cap:
            report.add("Caps", f"Per-run cap of {total_cap} new rows reached; skipped remaining segments")
            break
        seg = SEGMENTS.get(seg_id)
        if seg is None:
            report.add("Errors", f"Unknown discovery segment '{seg_id}'")
            continue
        seg_cap = min(d["max_new_rows_per_segment"], total_cap - added)
        focus = seg.focus_for(week)
        log.info("Searching: %s (focus: %s)", seg_id, "; ".join(focus))
        try:
            result = ctx.agent.research(
                system=DISCOVERY_SYSTEM,
                prompt=discovery_prompt(kind=seg.kind, region=seg.region, focus=focus,
                                        n=d["candidates_per_segment"], known=index.known_names(250)),
                record_tool=RECORD_COMPANIES,
                effort=d["effort"],
                web_searches=d["web_searches_per_segment"],
                web_fetches=d["web_fetches_per_segment"],
            )
        except AgentError as exc:
            report.add("Errors", f"{seg_id}: {exc}")
            continue

        candidates = result.data.get("companies", [])
        seg_added = 0
        for i, cand in enumerate(candidates):
            if seg_added >= seg_cap:
                report.add("Caps", f"{seg_id}: cap of {seg_cap} reached; {len(candidates) - i} candidate(s) left out")
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

            kind = "roaster" if cand.get("category") == "Specialty roaster" else "freight"
            source_url = clean(cand.get("source_url"))
            check = ctx.verifier.verify(email=email, phone=phone, contact_name=clean(cand.get("contact_name")),
                                        source_url=source_url, website=website, evidence_pages=result.pages)
            state_code = clean(cand.get("state"))
            adelaide = bool(cand.get("adelaide_presence")) or state_code == "SA"
            tab = _tab_for(cfg, kind, adelaide)
            services = " - ".join(x for x in (clean(cand.get("services")), clean(cand.get("coffee_evidence"))) if x)
            notes = check.notes + ([clean(cand["notes"])] if clean(cand.get("notes")) else [])
            values = {
                "company": company,
                "category": cand.get("category", ""),
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
                "record_type": "Lead" if kind == "roaster" else "Quote",
                "agent_notes": "; ".join(notes),
            }
            keys = entity_keys(company, website, check.email, check.phone)
            index.add(f"this run ({tab})", company, website, check.email, check.phone)
            line = (f"{company} -> {tab} ({values['city'] or '?'} {state_code}) "
                    f"{check.email or 'no verified email'}")
            if ctx.dry_run:
                report.add("Would add", line)
            else:
                for name in targets:
                    ctx.state.queue_row(name, tab, kind, values)
                ctx.state.add_discovered(company, kind, keys, source_url, ctx.today)
                report.add("Added (status: " + cfg.status("new") + ")", line)
            report.count(f"new rows - {tab}")
            seg_added += 1
            added += 1

    commit(ctx, ChangeSet())
    return added
