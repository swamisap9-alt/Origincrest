"""Workflow 2: research new roaster rows, score fit, draft a personalised intro."""

from __future__ import annotations

from datetime import date, timedelta

from . import compliance
from .dedupe import EMAIL_RE, norm_email, primary_key
from .llm import AgentError
from .mailer import build_message
from .prompts import EMAIL_SCHEMA, INTRO_SYSTEM, RECORD_RESEARCH, RESEARCH_SYSTEM, intro_prompt, research_prompt
from .runtime import ChangeSet, RunContext, commit
from .utils import clean, fmt_date, log
from .workbook import Row, WorkbookStore


def _first_email(text: str) -> str:
    match = EMAIL_RE.search(text or "")
    return norm_email(match.group(0)) if match else ""


def _candidates(ctx: RunContext, stores: dict[str, WorkbookStore]) -> dict[str, list[Row]]:
    cfg = ctx.cfg
    eligible = {s.strip().lower() for s in cfg["leads"]["eligible_statuses"]}
    groups: dict[str, list[Row]] = {}
    for store in stores.values():
        for view in store.views():
            if view.kind != "roaster" or not view.has("company"):
                continue
            for row in view.rows():
                company = row.text("company")
                if not company or row.text("draft_created"):
                    continue
                if row.text("status").lower() not in eligible:
                    continue
                key = primary_key(company, row.text("website"), _first_email(row.text("email")))
                groups.setdefault(key, []).append(row)

    def priority(item: tuple[str, list[Row]]):
        row = item[1][0]
        added_by_agent = row.text("added_by") == "OriginCrest Agent"
        adelaide = row.sheet.title == cfg["tabs"]["roasters_adelaide"]
        return (added_by_agent, not adelaide, row.index)

    return dict(sorted(groups.items(), key=priority))


def _research(ctx: RunContext, key: str, row: Row) -> dict | None:
    cfg = ctx.cfg
    cached = ctx.state.research(key, cfg["leads"]["research_cache_days"])
    if cached is not None:
        return cached
    lc = cfg["leads"]
    result = ctx.agent.research(
        system=RESEARCH_SYSTEM,
        prompt=research_prompt(row.context(), cfg),
        record_tool=RECORD_RESEARCH,
        effort=lc["research_effort"],
        web_searches=lc["web_searches_per_company"],
        web_fetches=lc["web_fetches_per_company"],
    )
    data = result.data
    data["fit_score"] = max(0, min(100, int(data.get("fit_score") or 0)))
    check = ctx.verifier.verify(
        email=clean(data.get("public_email")), phone="", contact_name=clean(data.get("contact_name")),
        source_url=clean(data.get("public_email_source_url")),
        website=clean(data.get("website")) or row.text("website"), evidence_pages=result.pages)
    data["verified_email"] = check.email
    data["verified_contact_name"] = check.contact_name
    data["verification_notes"] = check.notes
    ctx.state.save_research(key, data, ctx.today)
    return data


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


def run_leads(ctx: RunContext, max_drafts: int | None = None) -> int:
    cfg, report, lc = ctx.cfg, ctx.report, ctx.cfg["leads"]
    stores = {name: WorkbookStore(name, cfg.workbook_path(name), cfg).load() for name in lc["workbooks"]}
    groups = _candidates(ctx, stores)
    log.info("%d roaster(s) waiting for research or an intro draft", len(groups))
    changes = ChangeSet()
    for name in stores:
        changes.touch(name)  # also flushes any rows still queued from discovery
    if not groups:
        commit(ctx, changes)
        return 0

    sender_problems = compliance.sender_problems(cfg)
    if sender_problems:
        report.add("Errors", "Intro drafts are off until these are filled in config.yaml: "
                   + ", ".join(sender_problems))
    research_budget = lc["max_research_per_run"]
    draft_budget = lc["max_drafts_per_run"] if max_drafts is None else max_drafts
    approved = {s.lower() for s in lc["approved_statuses"]}
    new_status = cfg.status("new").lower()
    mailbox = None
    mailbox_failed = False
    drafted = 0

    try:
        for key, rows in groups.items():
            row = rows[0]
            company = row.text("company")
            if ctx.state.research(key, lc["research_cache_days"]) is None:
                if research_budget <= 0:
                    report.add("Deferred (research cap)", company)
                    continue
                research_budget -= 1
            try:
                data = _research(ctx, key, row)
            except AgentError as exc:
                report.add("Errors", f"{company}: research failed - {exc}")
                continue
            values = _research_values(data)
            values["researched_on"] = ctx.today
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
            status = row.text("status").lower()
            line = f"{company}: fit {score}/100 ({values['fit_origins'] or 'no origin match'})"

            def finish(note: str, status_value: str | None = None) -> None:
                if status_value:
                    values["status"] = status_value
                for r in rows:
                    changes.update(r, values, note="; ".join(notes + [note]))

            if not email:
                report.add("No public email", line)
                finish("Researched - no published business email found", cfg.status("no_email"))
                continue
            if ctx.state.is_suppressed(email):
                report.add("Suppressed", f"{line} - {email} has unsubscribed")
                finish(f"{email} is on the unsubscribe list", cfg.status("unsubscribed"))
                continue
            if score < lc["min_fit_score_for_draft"] and status not in approved:
                report.add("Low fit (no draft)", line)
                finish(f"Fit {score} is below {lc['min_fit_score_for_draft']} - no draft. "
                       "Set status to Approved to draft anyway.", cfg.status("low_fit"))
                continue
            if status == new_status and not lc["draft_for_unreviewed"]:
                values.pop("researched_on")
                report.add("Waiting for your review", line)
                finish("Researched - change status to Reviewed or Approved to get an intro draft")
                continue
            if sender_problems:
                values.pop("researched_on")
                finish("Researched - intro draft waiting for sender details in config.yaml")
                continue

            existing = ctx.state.draft(key, "intro", 1)
            if existing is None:
                if draft_budget <= 0:
                    values.pop("researched_on")
                    report.add("Deferred (draft cap)", line)
                    finish("Researched - intro draft deferred (per-run cap)")
                    continue
                if mailbox is None and not mailbox_failed and ctx.mailbox_factory:
                    try:
                        mailbox = ctx.mailbox_factory()
                        mailbox.connect()
                    except Exception as exc:  # noqa: BLE001 - any IMAP/auth/network failure
                        mailbox_failed = True
                        mailbox = None
                        report.add("Errors", f"IMAP connection failed, no drafts saved: {exc}")
                if mailbox is not None:
                    earlier = mailbox.messages_to(email, ctx.today - timedelta(days=365))
                    if earlier and earlier[0].when:
                        sent_on = earlier[0].when.astimezone().date()
                        report.add("Already contacted", f"{line} - you emailed {email} on {fmt_date(sent_on)}")
                        values.update(last_contacted=sent_on,
                                      next_action_date=sent_on + timedelta(days=cfg["followup"]["lead_days"]))
                        finish(f"Found your email to {email} from {fmt_date(sent_on)} in Sent - set to Open",
                               cfg.status("open"))
                        continue
                if mailbox is None and not ctx.dry_run:
                    values.pop("researched_on")
                    finish("Researched - intro draft waiting (email connection failed)")
                    continue
                try:
                    draft = ctx.agent.structured(system=INTRO_SYSTEM,
                                                 prompt=intro_prompt(_email_context(data), contact, cfg),
                                                 schema=EMAIL_SCHEMA, effort=lc["email_effort"])
                except AgentError as exc:
                    report.add("Errors", f"{company}: drafting failed - {exc}")
                    continue
                source = row.text("source_url") or clean(data.get("public_email_source_url"))
                body = compliance.finish_body(cfg, draft["body"], record_type="lead", source_url=source)
                msg = build_message(
                    from_name=cfg["sender"]["name"], from_addr=cfg["sender"]["email"], to_addr=email,
                    to_name=contact, subject=clean(draft["subject"]) or "Indonesian single-origin green coffee",
                    body=body,
                    unsubscribe_addr=compliance.unsubscribe_address(cfg), agent_tag=f"intro; {row.ref}")
                draft_budget -= 1
                if ctx.dry_run:
                    ctx.preview_dir.mkdir(parents=True, exist_ok=True)
                    path = ctx.preview_dir / f"intro_{key.split(':', 1)[-1].replace('/', '_')}.eml"
                    path.write_bytes(msg.as_bytes())
                    report.add("Would draft", f"{line} -> {email} (preview: {path})")
                    continue
                mailbox.append_draft(msg)
                ctx.state.record_draft(key, "intro", 1, email, msg["Subject"], msg["Message-ID"], row.ref, ctx.today)
                drafted += 1
                subject, created = msg["Subject"], ctx.today
            else:
                subject, created = existing["subject"], date.fromisoformat(existing["created_on"])
            values.update(draft_created=created, draft_subject=subject)
            report.add("Intro drafted (in your Drafts folder)", f"{line} -> {email}: \"{subject}\"")
            report.count("intro drafts")
            finish(f"Intro draft saved {fmt_date(created)} - review and send it from Drafts",
                   cfg.status("intro_drafted"))
    finally:
        if mailbox is not None:
            mailbox.close()

    commit(ctx, changes)
    return drafted
