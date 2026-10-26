"""Workflow 3: follow up on open RFQs and leads (weekdays).

For each "Open" row: sync "Last contacted" from your Sent folder, look for a
reply (or unsubscribe / bounce), and if nothing came back within the window
(5 days for quotes, 10 for leads by default) draft a threaded follow-up.
"""

from __future__ import annotations

from datetime import date, datetime, timedelta, timezone

from . import compliance
from .dedupe import EMAIL_RE, company_domain, email_domain, is_free_mail, norm_email, primary_key
from .llm import AgentError
from .mailer import MailInfo, build_message
from .prompts import EMAIL_SCHEMA, FOLLOWUP_SYSTEM, followup_prompt
from .runtime import ChangeSet, RunContext, commit
from .utils import clean, fmt_date, log, to_date, to_int
from .workbook import Row, WorkbookStore


def _utc(dt: datetime) -> datetime:
    return dt.replace(tzinfo=timezone.utc) if dt.tzinfo is None else dt


def _record_type(row: Row) -> str:
    return "lead" if row.kind == "buyer" else "quote"


def _local_date(info: MailInfo, fallback: date) -> date:
    return info.when.astimezone().date() if info.when else fallback


def run_followup(ctx: RunContext, max_drafts: int | None = None) -> int:
    cfg, report, fc = ctx.cfg, ctx.report, ctx.cfg["followup"]
    names = [n for n in cfg["workbooks"] if n in fc["workbooks"] or cfg.lead_tabs(n)]
    stores = {}
    for name in names:
        try:
            stores[name] = WorkbookStore(name, cfg.workbook_path(name), cfg).load()
        except FileNotFoundError as exc:
            if name in fc["workbooks"]:
                raise
            report.add("Errors", str(exc))
    open_statuses = {s.strip().lower() for s in fc["open_statuses"]}
    awaiting_send = {cfg.status("intro_drafted").lower(), cfg.status("rfq_drafted").lower()}
    groups: dict[str, list[Row]] = {}
    for store in stores.values():
        for view in store.views():
            if not view.has("status"):
                continue
            if store.name not in fc["workbooks"] and view.title not in cfg.lead_tabs(store.name):
                continue
            for row in view.rows():
                status = row.text("status").lower()
                if status in open_statuses or status in awaiting_send:
                    match = EMAIL_RE.search(row.text("email"))
                    key = primary_key(row.text("company"), row.text("website"), match.group(0) if match else "")
                    groups.setdefault(key, []).append(row)
    changes = ChangeSet()
    for name in stores:
        changes.touch(name)
    log.info("%d open / awaiting-send row(s) to check", len(groups))
    if not groups:
        commit(ctx, changes)
        return 0

    sender_problems = compliance.sender_problems(cfg)
    if sender_problems:
        report.add("Errors", "Follow-up drafts are off until these are filled in config.yaml: "
                   + ", ".join(sender_problems))
    budget = fc["max_drafts_per_run"] if max_drafts is None else max_drafts
    today = ctx.today
    drafted = 0
    mailbox = ctx.mailbox_factory()
    mailbox.connect()
    try:
        for key, rows in groups.items():
            row = rows[0]
            company = row.text("company") or "(no company name)"
            rtype = _record_type(row)
            days = fc["lead_days"] if rtype == "lead" else fc["quote_days"]
            label = f"{company} [{rtype}, {row.sheet.title}]"
            match = EMAIL_RE.search(row.text("email"))
            email = norm_email(match.group(0)) if match else ""
            status = row.text("status").lower()

            def update(values: dict | None = None, note: str = "") -> None:
                for r in rows:
                    changes.update(r, values, note)

            if not email:
                report.add("Flagged - no email on row", label)
                update(note="Open but no email address on the row - can't check replies or follow up")
                continue
            if ctx.state.is_suppressed(email):
                report.add("Unsubscribed", f"{label} - {email}")
                update({"status": cfg.status("unsubscribed")}, f"{email} is on the unsubscribe list")
                continue

            sent = mailbox.messages_to(email, today - timedelta(days=fc["sent_lookback_days"]))
            last_sent = sent[0] if sent else None
            last = to_date(row.get("last_contacted"))
            values: dict = {}

            if status in awaiting_send:
                drafted_on = to_date(row.get("draft_created")) or to_date(row.get("date_added")) or today
                if not last_sent or _local_date(last_sent, today) < drafted_on:
                    continue  # draft still waiting to be sent
                sent_on = _local_date(last_sent, today)
                values.update(status=cfg.status("open"), last_contacted=sent_on,
                              next_action_date=sent_on + timedelta(days=days))
                report.add("Sent - now Open", f"{label} (sent {fmt_date(sent_on)})")
                last = sent_on
                status = "open"
            elif last_sent:
                sent_on = _local_date(last_sent, today)
                if last is None or sent_on > last:
                    values["last_contacted"] = sent_on
                    nxt = to_date(row.get("next_action_date"))
                    if nxt is None or nxt < sent_on + timedelta(days=days):
                        values["next_action_date"] = sent_on + timedelta(days=days)
                    last = sent_on

            if last is None:
                report.add("Flagged - no date", f"{label}: no 'Last contacted' date and nothing in Sent")
                update(values, "Open but no last-contacted date - add one so follow-ups can be timed")
                continue

            # 1. Replies first.
            domain = "" if is_free_mail(email) else company_domain(email_domain(email))
            replies = mailbox.messages_from(email, last - timedelta(days=1), domain=domain)
            if last_sent and last_sent.when:
                replies = [r for r in replies if r.when is None or _utc(r.when) > _utc(last_sent.when)]
            else:
                replies = [r for r in replies if _local_date(r, today) >= last]
            genuine = [r for r in replies if not r.auto_reply and not r.is_bounce]
            if unsub := next((r for r in genuine if r.is_unsubscribe), None):
                ctx.state.suppress(email, f"replied '{unsub.subject}'")
                values["status"] = cfg.status("unsubscribed")
                report.add("Unsubscribe requests (honoured)", f"{label} - {email}")
                update(values, f"Unsubscribe request received {fmt_date(_local_date(unsub, today))} - "
                               "added to the do-not-contact list")
                continue
            if genuine:
                reply = genuine[0]
                values.update(status=cfg.status("replied"), last_reply=_local_date(reply, today),
                              next_action_date=today)
                report.add("Replies found", f"{label} - \"{reply.subject}\" from {reply.from_addr}")
                update(values, f"Reply received {fmt_date(_local_date(reply, today))}: {reply.subject[:80]}")
                continue
            if bounce := next(iter(mailbox.bounces_for(email, last)), None):
                values["status"] = cfg.status("bounced")
                report.add("Bounced", f"{label} - {email}")
                update(values, f"Email bounced {fmt_date(_local_date(bounce, today))} - check the address")
                continue
            auto_note = ""
            if auto := next((r for r in replies if r.auto_reply), None):
                auto_note = f"Auto-reply only ({fmt_date(_local_date(auto, today))}: {auto.subject[:60]})"

            # 2. Due?
            nxt = values.get("next_action_date") or to_date(row.get("next_action_date"))
            due_on = nxt or last + timedelta(days=days)
            if today < due_on:
                update(values, auto_note)
                continue
            count = to_int(row.get("followup_count"))
            overdue = (today - last).days
            label += f" - {overdue} days since last contact"
            if count >= fc["max_followups_per_contact"]:
                report.add("Flagged - max follow-ups reached", label)
                update(values, f"No reply after {count} follow-ups - consider closing this row")
                continue
            step = count + 1
            existing = ctx.state.draft(key, "followup", step)
            if existing is not None:
                # Drafted on an earlier run; make sure the sheet reflects it.
                made = date.fromisoformat(existing["created_on"])
                values.update(last_contacted=max(made, last), next_action_date=max(made, last) + timedelta(days=days),
                              followup_count=step)
                waiting = mailbox.draft_waiting(existing["message_id"], email)
                report.add("Flagged - follow-up already drafted", f"{label} (#{step}, "
                           f"{'still in Drafts' if waiting else 'drafted'} {fmt_date(made)})")
                update(values, f"Follow-up #{step} drafted {fmt_date(made)}"
                               + (" - still waiting in Drafts" if waiting else ""))
                continue
            if mailbox.draft_waiting(to_addr=email):
                report.add("Flagged - unsent draft waiting", label)
                update(values, "Due for follow-up, but a draft to this address is already in Drafts - "
                               "send or delete it first")
                continue
            if sender_problems or budget <= 0:
                why = "sender details missing in config.yaml" if sender_problems else "per-run draft cap reached"
                report.add("Flagged - due, not drafted", f"{label} ({why})")
                update(values, f"Due for follow-up #{step} ({why})")
                continue

            threaded = bool(last_sent and last_sent.message_id)
            try:
                draft = ctx.agent.structured(
                    system=FOLLOWUP_SYSTEM,
                    prompt=followup_prompt(record_type=rtype, number=step, row_context=row.email_context(),
                                           original=(last_sent.text[:1500] if last_sent else ""),
                                           days_since=overdue, threaded=threaded, cfg=cfg),
                    schema=EMAIL_SCHEMA, effort=fc["effort"])
            except AgentError as exc:
                report.add("Errors", f"{company}: follow-up drafting failed - {exc}")
                continue
            subject = clean(draft["subject"])
            if threaded:
                subject = last_sent.subject if last_sent.subject.lower().startswith("re:") else f"Re: {last_sent.subject}"
            body = compliance.finish_body(cfg, draft["body"], record_type=rtype, source_url=row.text("source_url"))
            if threaded and last_sent.text:
                quoted = "\n".join("> " + line for line in last_sent.text[:2000].splitlines())
                when = fmt_date(_local_date(last_sent, today))
                body = f"{body}\nOn {when}, {cfg['sender']['name']} wrote:\n{quoted}\n"
            msg = build_message(
                from_name=cfg["sender"]["name"], from_addr=cfg["sender"]["email"], to_addr=email,
                to_name=row.text("contact_name"), subject=subject, body=body,
                in_reply_to=last_sent.message_id if threaded else "",
                references=last_sent.references if threaded else "",
                unsubscribe_addr=compliance.unsubscribe_address(cfg) if compliance.needs_unsubscribe(cfg, rtype) else "",
                agent_tag=f"followup {step}; {row.ref}")
            if ctx.dry_run:
                ctx.preview_dir.mkdir(parents=True, exist_ok=True)
                path = ctx.preview_dir / f"followup{step}_{key.split(':', 1)[-1].replace('/', '_')}.eml"
                path.write_bytes(msg.as_bytes())
                report.add("Would draft follow-up", f"{label} (preview: {path})")
                continue
            mailbox.append_draft(msg)
            ctx.state.record_draft(key, "followup", step, email, subject, msg["Message-ID"], row.ref, today)
            budget -= 1
            drafted += 1
            values.update(last_contacted=today, next_action_date=today + timedelta(days=days), followup_count=step)
            report.add("Follow-ups drafted (in your Drafts folder)", f"{label}: #{step} \"{subject}\"")
            report.count("follow-up drafts")
            update(values, f"Follow-up #{step} drafted - review and send it from Drafts"
                           + (f"; {auto_note}" if auto_note else ""))
    finally:
        mailbox.close()

    commit(ctx, changes)
    return drafted
