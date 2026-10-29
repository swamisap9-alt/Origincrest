"""Workflow 2: turn new leads into email drafts tailored to each row.

* Coffee businesses (roasters, importers, wholesalers, traders...): web research, a fit score, and a
  personalised intro email.
* Freight forwarders, customs brokers and 3PLs: a request for quotation built from your requirements,
  the rows in your RFQ workbook and the last RFQ you sent.

Works on each workbook's new-lead tabs (plus leads.extra_tabs). Rows elsewhere are never emailed.
"""

from __future__ import annotations

import re
from datetime import date, timedelta
from typing import Callable

from . import compliance
from .dedupe import EMAIL_RE, core_name, norm_email, primary_key
from .llm import AgentError
from .mailer import build_message
from .prompts import (EMAIL_SCHEMA, INTRO_SYSTEM, RECORD_RESEARCH, RESEARCH_SYSTEM, RFQ_SYSTEM, _filled,
                      intro_prompt, research_prompt, rfq_prompt)
from .runtime import ChangeSet, RunContext, commit
from .utils import clean, fmt_date, log, to_date
from .workbook import Row, WorkbookStore

AGENT_NAME = "OriginCrest Agent"
_PLACEHOLDER = re.compile(r"\[[^\]]{3,80}\]")


def _first_email(text: str) -> str:
    match = EMAIL_RE.search(text or "")
    return norm_email(match.group(0)) if match else ""


def _candidates(ctx: RunContext, stores: dict[str, WorkbookStore],
                redraft: list[str] | None = None) -> dict[str, list[Row]]:
    """Rows waiting for a draft - or, with `redraft`, the rows of those companies whatever their status."""
    cfg = ctx.cfg
    eligible = {s.strip().lower() for s in cfg["leads"]["eligible_statuses"]}
    never = cfg.skip_statuses | {cfg.status("unsubscribed").lower()}
    wanted = [core_name(name) for name in redraft or []]
    groups: dict[str, list[Row]] = {}
    for store in stores.values():
        tabs = cfg.lead_tabs(store.name)
        for view in store.views():
            if view.title not in tabs or not view.has("company"):
                continue
            for row in view.rows():
                company = row.text("company")
                if not company:
                    continue
                status = row.text("status").lower()
                if redraft is not None:
                    if status in never or not any(w and w in core_name(company) for w in wanted):
                        continue
                elif row.text("draft_created") or status not in eligible:
                    continue
                key = primary_key(company, row.text("website"), _first_email(row.text("email")))
                groups.setdefault(key, []).append(row)

    def priority(item: tuple[str, list[Row]]):
        row = item[1][0]
        profile = cfg.profile(row.sheet.store.name) or {}
        return (row.text("added_by") == AGENT_NAME, row.sheet.title != profile.get("adelaide_tab"), row.index)

    return dict(sorted(groups.items(), key=priority))


def _research_values(data: dict) -> dict:
    rationale = clean(data.get("fit_rationale"))
    if clean(data.get("red_flags")):
        rationale += f" Red flags: {clean(data['red_flags'])}"
    size = clean(data.get("size_estimate"))
    if clean(data.get("size_evidence")):
        size += f" ({clean(data['size_evidence'])})"
    return {
        "roast_profile": clean(data.get("what_they_roast")),
        "size": size,
        "sourcing": clean(data.get("sourcing")),
        "indonesian_now": clean(data.get("current_indonesian_coffees")) or "None found",
        "fit_score": data.get("fit_score", 0),
        "fit_origins": ", ".join(data.get("best_fit_origins") or []),
        "fit_rationale": rationale,
        "record_type": "Lead",
    }


def _email_context(data: dict) -> dict:
    keep = ("company_name", "city", "state", "what_they_roast", "sourcing", "current_indonesian_coffees",
            "wholesale", "best_fit_origins", "fit_rationale", "personalisation_hooks")
    return {k: data.get(k) for k in keep}


class LeadRun:
    def __init__(self, ctx: RunContext, stores: dict[str, WorkbookStore], max_drafts: int | None,
                 redraft: bool = False):
        self.redraft = redraft  # draft again for companies whose earlier draft was deleted
        self.ctx, self.cfg, self.report = ctx, ctx.cfg, ctx.report
        self.lc = ctx.cfg["leads"]
        self.stores = stores
        self.changes = ChangeSet()
        self.research_budget = self.lc["max_research_per_run"]
        self.draft_budget = self.lc["max_drafts_per_run"] if max_drafts is None else max_drafts
        self.approved = {s.lower() for s in self.lc["approved_statuses"]}
        self._sender: compliance.Sender | None = None
        self.mailbox = None
        self.mailbox_failed = False
        self.drafted = 0
        self._rfq_context: tuple[str, list[dict], str] | None = None

    # helpers -------------------------------------------------------------
    def _mailbox(self):
        if self.mailbox is None and not self.mailbox_failed and self.ctx.mailbox_factory:
            try:
                self.mailbox = self.ctx.mailbox_factory()
                self.mailbox.connect()
            except Exception as exc:  # noqa: BLE001 - any IMAP/auth/network failure
                self.mailbox_failed = True
                self.mailbox = None
                self.report.add("Errors", f"IMAP connection failed, no drafts saved: {exc}")
        return self.mailbox

    def _sender_details(self) -> compliance.Sender:
        """From name/address and signature - learned from your Sent folder unless config.yaml says otherwise."""
        if self._sender is None:
            self._sender = compliance.resolve_sender(self.cfg, self._mailbox())
            compliance.describe_sender(self._sender, self.cfg, self.report)
        return self._sender

    def _finish(self, rows: list[Row], values: dict, notes: list[str], note: str, status: str | None = None) -> None:
        if status:
            values["status"] = status
        for r in rows:
            self.changes.update(r, values, note="; ".join(notes + [note]))

    def _rfq_examples(self) -> tuple[str, list[dict], str]:
        """What OriginCrest asks forwarders to quote: config text, RFQ tracker rows, last RFQ sent."""
        if self._rfq_context is not None:
            return self._rfq_context
        rc = self.cfg["rfq"]
        requirements = _filled(rc.get("requirements"))
        examples: list[tuple[date, dict, str]] = []
        name = rc.get("example_workbook")
        if name in self.cfg["workbooks"]:
            store = self.stores.get(name)
            if store is None:
                try:
                    store = WorkbookStore(name, self.cfg.workbook_path(name), self.cfg).load()
                except FileNotFoundError:
                    store = None
            for view in store.views() if store else []:
                if view.title in self.cfg.lead_tabs(name):
                    continue
                for row in view.rows():
                    sent_on = to_date(row.get("last_contacted"))
                    status = row.text("status").lower()
                    if row.kind != "freight" or not (sent_on or status in ("open", "quoted", "replied - review")):
                        continue
                    context = row.email_context(skip=("email", "phone", "contact_name", "contact_role", "website"))
                    examples.append((sent_on or date.min, context, _first_email(row.text("email"))))
        examples.sort(key=lambda e: e[0], reverse=True)
        earlier = ""
        mailbox = self._mailbox()
        for _, _, address in examples[:5]:
            if not mailbox or not address:
                continue
            sent = mailbox.messages_to(address, self.ctx.today - timedelta(days=365))
            if sent and sent[0].text:
                earlier = f"Subject: {sent[0].subject}\n\n{sent[0].text[:1500]}"
                break
        self._rfq_context = (requirements, [e[1] for e in examples[:5]], earlier)
        if not (requirements or examples or earlier):
            self.report.add("RFQ details", "No RFQ requirements, tracker rows or earlier RFQ found - drafts will "
                                           "contain [placeholders] for the shipment details")
        return self._rfq_context

    # one lead --------------------------------------------------------------
    def _draft(self, key: str, rows: list[Row], values: dict, notes: list[str], *, email: str, contact: str,
               kind: str, line: str, compose: Callable[[], dict]) -> None:
        """Shared checks, then compose and save the draft (intro for buyers, RFQ for freight)."""
        cfg, report, row = self.cfg, self.report, rows[0]
        workflow = "intro" if kind == "buyer" else "rfq"
        status = row.text("status").lower()
        if not email:
            report.add("No public email", line)
            self._finish(rows, values, notes, "No published business email found - add one and set the status "
                                              "to Approved to get a draft", cfg.status("no_email"))
            return
        if self.ctx.state.is_suppressed(email):
            report.add("Suppressed", f"{line} - {email} has unsubscribed")
            self._finish(rows, values, notes, f"{email} is on the unsubscribe list", cfg.status("unsubscribed"))
            return
        if status == cfg.status("new").lower() and not self.lc["draft_for_unreviewed"]:
            values.pop("researched_on", None)
            report.add("Waiting for your review", line)
            self._finish(rows, values, notes, "Change status to Reviewed or Approved to get a draft")
            return
        existing = None if self.redraft else self.ctx.state.draft(key, workflow, 1)
        if existing is not None:
            subject, created = existing["subject"], date.fromisoformat(existing["created_on"])
        else:
            if self.draft_budget <= 0:
                values.pop("researched_on", None)
                report.add("Deferred (draft cap)", line)
                self._finish(rows, values, notes, "Draft deferred (per-run cap)")
                return
            mailbox = self._mailbox()
            if mailbox is not None and self.redraft and mailbox.draft_waiting(to_addr=email):
                report.add("Still in Drafts", f"{line} - a draft to {email} is already in your Drafts folder")
                return
            if mailbox is not None:
                earlier = mailbox.messages_to(email, self.ctx.today - timedelta(days=365))
                if earlier and earlier[0].when:
                    sent_on = earlier[0].when.astimezone().date()
                    days = cfg["followup"]["lead_days" if kind == "buyer" else "quote_days"]
                    report.add("Already contacted", f"{line} - you emailed {email} on {fmt_date(sent_on)}")
                    values.update(last_contacted=sent_on, next_action_date=sent_on + timedelta(days=days))
                    self._finish(rows, values, notes, f"Found your email to {email} from {fmt_date(sent_on)} in "
                                                      "Sent - set to Open", cfg.status("open"))
                    return
            elif not self.ctx.dry_run:
                values.pop("researched_on", None)
                self._finish(rows, values, notes, "Draft waiting (email connection failed)")
                return
            sender = self._sender_details()
            if sender.problems:
                values.pop("researched_on", None)
                self._finish(rows, values, notes, "Draft waiting for sender details: " + ", ".join(sender.problems))
                return
            try:
                draft = compose()
            except AgentError as exc:
                report.add("Errors", f"{row.text('company')}: drafting failed - {exc}")
                return
            record_type = "lead" if kind == "buyer" else "quote"
            body = sender.finish_body(cfg, draft["body"], record_type=record_type, source_url=row.text("source_url"))
            fallback = "Indonesian single-origin green coffee" if kind == "buyer" else "Request for quotation"
            msg = build_message(
                from_name=sender.name, from_addr=sender.email, to_addr=email, to_name=contact,
                subject=clean(draft["subject"]) or fallback, body=body,
                unsubscribe_addr=sender.unsubscribe_email, agent_tag=f"{workflow}; {row.ref}")
            self.draft_budget -= 1
            if _PLACEHOLDER.search(draft["body"]):
                notes.append("Fill in the [bracketed] details before sending")
            if self.ctx.dry_run:
                self.ctx.preview_dir.mkdir(parents=True, exist_ok=True)
                path = self.ctx.preview_dir / f"{workflow}_{key.split(':', 1)[-1].replace('/', '_')}.eml"
                path.write_bytes(msg.as_bytes())
                report.add("Would draft", f"{line} -> {email} (preview: {path})")
                return
            mailbox.append_draft(msg)
            self.ctx.state.record_draft(key, workflow, 1, email, msg["Subject"], msg["Message-ID"], row.ref,
                                        self.ctx.today)
            self.drafted += 1
            subject, created = msg["Subject"], self.ctx.today
        values.update(draft_created=created, draft_subject=subject)
        label = "Intro" if kind == "buyer" else "RFQ"
        report.add(f"{label} drafted (in your Drafts folder)", f"{line} -> {email}: \"{subject}\"")
        report.count(f"{label.lower()} drafts")
        self._finish(rows, values, notes, f"{label} draft saved {fmt_date(created)} - review and send it from Drafts",
                     cfg.status("intro_drafted" if kind == "buyer" else "rfq_drafted"))

    def buyer(self, key: str, rows: list[Row]) -> None:
        cfg, lc, row = self.cfg, self.lc, rows[0]
        company = row.text("company")
        data = self.ctx.state.research(key, lc["research_cache_days"])
        if data is None:
            if self.research_budget <= 0:
                self.report.add("Deferred (research cap)", company)
                return
            self.research_budget -= 1
            try:
                result = self.ctx.agent.research(
                    system=RESEARCH_SYSTEM, prompt=research_prompt(row.context(), cfg), record_tool=RECORD_RESEARCH,
                    effort=lc["research_effort"], web_searches=lc["web_searches_per_company"],
                    web_fetches=lc["web_fetches_per_company"])
            except AgentError as exc:
                self.report.add("Errors", f"{company}: research failed - {exc}")
                return
            data = result.data
            data["fit_score"] = max(0, min(100, int(data.get("fit_score") or 0)))
            check = self.ctx.verifier.verify(
                email=clean(data.get("public_email")), phone="", contact_name=clean(data.get("contact_name")),
                source_url=clean(data.get("public_email_source_url")),
                website=clean(data.get("website")) or row.text("website"), evidence_pages=result.pages)
            data.update(verified_email=check.email, verified_contact_name=check.contact_name,
                        verification_notes=check.notes)
            self.ctx.state.save_research(key, data, self.ctx.today)

        values = _research_values(data)
        values["researched_on"] = self.ctx.today
        notes = list(data.get("verification_notes") or [])
        email = next((e for r in rows if (e := _first_email(r.text("email")))), "") or data.get("verified_email", "")
        if email and not row.text("email"):
            values["email"] = email
        contact = row.text("contact_name") or data.get("verified_contact_name", "")
        if contact and not row.text("contact_name"):
            values["contact_name"] = contact
            values["contact_role"] = clean(data.get("contact_role"))
        if data.get("website") and not row.text("website"):
            values["website"] = data["website"]
        score = data["fit_score"]
        category = row.text("category")
        line = (f"{company}{f' [{category}]' if category else ''}: fit {score}/100 "
                f"({values['fit_origins'] or 'no origin match'})")
        if (email and score < lc["min_fit_score_for_draft"] and not self.redraft
                and row.text("status").lower() not in self.approved):
            self.report.add("Low fit (no draft)", line)
            self._finish(rows, values, notes, f"Fit {score} is below {lc['min_fit_score_for_draft']} - no draft. "
                                              "Set status to Approved to draft anyway.", cfg.status("low_fit"))
            return
        self._draft(key, rows, values, notes, email=email, contact=contact, kind="buyer", line=line,
                    compose=lambda: self.ctx.agent.structured(
                        system=INTRO_SYSTEM, prompt=intro_prompt(_email_context(data), row.email_context(), contact, cfg),
                        schema=EMAIL_SCHEMA, effort=lc["email_effort"]))

    def freight(self, key: str, rows: list[Row]) -> None:
        cfg, row = self.cfg, rows[0]
        email = next((e for r in rows if (e := _first_email(r.text("email")))), "")
        contact = row.text("contact_name")
        category = row.text("category")
        line = f"{row.text('company')}{f' [{category}]' if category else ''} ({row.text('city') or row.sheet.title})"

        def compose() -> dict:
            requirements, examples, earlier = self._rfq_examples()
            return self.ctx.agent.structured(
                system=RFQ_SYSTEM,
                prompt=rfq_prompt(row.email_context(), requirements, examples, earlier, contact, cfg),
                schema=EMAIL_SCHEMA, effort=self.lc["email_effort"])

        self._draft(key, rows, {"record_type": "Quote"}, [], email=email, contact=contact, kind="freight",
                    line=line, compose=compose)


def run_leads(ctx: RunContext, max_drafts: int | None = None, redraft: list[str] | None = None) -> int:
    """Draft for new leads; with `redraft` (company names), draft again for just those companies."""
    cfg = ctx.cfg
    names = [n for n in cfg["workbooks"] if cfg.lead_tabs(n)]
    stores = {}
    for name in names:
        try:
            stores[name] = WorkbookStore(name, cfg.workbook_path(name), cfg).load()
        except FileNotFoundError as exc:
            ctx.report.add("Errors", str(exc))
    groups = _candidates(ctx, stores, redraft)
    if redraft is not None:
        found = [r.text("company") for rows in groups.values() for r in rows]
        for name in redraft:
            if not any(core_name(name) in core_name(c) for c in found):
                ctx.report.add("Not found", f"No row for '{name}' in the new-lead tabs")
    log.info("%d lead(s) to research or draft", len(groups))
    run = LeadRun(ctx, stores, max_drafts, redraft=redraft is not None)
    for name in stores:
        run.changes.touch(name)  # also writes any rows still queued from discovery
    try:
        for key, rows in groups.items():
            if rows[0].kind == "freight":
                run.freight(key, rows)
            else:
                run.buyer(key, rows)
    finally:
        if run.mailbox is not None:
            run.mailbox.close()
    commit(ctx, run.changes)
    return run.drafted
