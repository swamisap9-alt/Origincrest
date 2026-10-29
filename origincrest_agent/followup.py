"""Workflow 3: follow up on open RFQs and leads (weekdays).

For each open row (an open Status, or no status but a "Date sent"): sync
"Last contacted" from your Sent folder, look for a reply (or unsubscribe /
bounce), and if nothing came back by the row's "Followup date" (or within 5
days for quotes, 10 for leads by default) draft a threaded follow-up.

Rows whose Notes say to phone them ("need to be called") get their draft
marked [CALL FIRST], with the phone number and points for the call at the top;
without an email address they get a call reminder addressed to you instead.
"""

from __future__ import annotations

import re
from datetime import date, datetime, timedelta, timezone

from . import compliance
from .dedupe import EMAIL_RE, company_domain, email_domain, is_free_mail, norm_email, primary_key
from .llm import AgentError
from .mailer import MailInfo, build_message
from .prompts import FOLLOWUP_SCHEMA, FOLLOWUP_SYSTEM, followup_prompt
from .runtime import ChangeSet, RunContext, commit
from .utils import clean, fmt_date, log, to_date, to_int
from .workbook import Row, WorkbookStore

CALL_TAG = "[CALL FIRST] "


def _utc(dt: datetime) -> datetime:
    return dt.replace(tzinfo=timezone.utc) if dt.tzinfo is None else dt


def _record_type(row: Row) -> str:
    return "lead" if row.kind == "buyer" else "quote"


def _local_date(info: MailInfo, fallback: date) -> date:
    return info.when.astimezone().date() if info.when else fallback


def _norm(text: str) -> str:
    """Lower case, with dashes and curly apostrophes evened out ("Follow-up" == "follow up")."""
    return re.sub(r"\s+", " ", str(text or "").lower().replace("’", "'").replace("-", " ")).strip()


def _phrase(text: str, phrases) -> str:
    """The first of `phrases` that appears in `text` as whole words, or ''."""
    text = _norm(text)
    for phrase in phrases:
        if (p := _norm(phrase)) and re.search(r"(?<!\w)" + re.escape(p) + r"(?!\w)", text):
            return phrase
    return ""


def _call_block(row: Row, *, step: int, days_since: int, last: date, rtype: str, points: list[str],
                has_email: bool) -> str:
    who = ", ".join(x for x in (row.text("contact_name"), row.text("contact_role")) if x)
    phone = row.text("phone") or "(no phone number on the row - check their website)"
    sent_what = "RFQ" if rtype == "quote" else "first email"
    lines = [
        "=" * 22 + " CALL FIRST " + "=" * 22,
        f"Your note: {row.text('user_notes')}",
        f"Call {row.text('company') or 'them'}{f' ({who})' if who else ''} on {phone}",
        f"Follow-up #{step} is due: {days_since} days since the {sent_what} went out ({fmt_date(last)}).",
    ]
    if about := row.text("services"):
        lines.append(f"About: {about}")
    if points:
        lines += ["Points for the call:", *(f"- {p}" for p in points)]
    if has_email:
        lines.append("The email below is only for after the call, if you still want to write. Before sending it, "
                     'delete this block and "[CALL FIRST]" from the subject.')
    else:
        lines.append("There is no email address on the row, so this reminder is addressed to you.")
    lines += ["After the call, update the row's Notes and Followup date (or its Status) so the agent knows.",
              "=" * 56]
    return "\n".join(lines)


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
    open_statuses = {_norm(s) for s in fc["open_statuses"]}
    awaiting_send = {_norm(cfg.status("intro_drafted")), _norm(cfg.status("rfq_drafted"))}
    call_phrases, done_phrases = fc.get("call_phrases") or [], fc.get("done_phrases") or []
    groups: dict[str, list[Row]] = {}
    for store in stores.values():
        for view in store.views():
            if store.name not in fc["workbooks"] and view.title not in cfg.lead_tabs(store.name):
                continue
            for row in view.rows():
                status = _norm(row.text("status"))
                if status in open_statuses or status in awaiting_send:
                    pass
                elif status or not fc.get("blank_status_is_open", True):
                    continue
                elif row.last_contact_date() is None or _phrase(row.text("user_notes"), done_phrases):
                    continue  # no status: open only once it has a "Date sent", and until the notes close it
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

    budget = fc["max_drafts_per_run"] if max_drafts is None else max_drafts
    today = ctx.today
    drafted = 0
    mailbox = ctx.mailbox_factory()
    mailbox.connect()
    sender = compliance.resolve_sender(cfg, mailbox)
    compliance.describe_sender(sender, cfg, report)
    try:
        for key, rows in groups.items():
            row = rows[0]
            company = row.text("company") or "(no company name)"
            rtype = _record_type(row)
            days = fc["lead_days"] if rtype == "lead" else fc["quote_days"]
            label = f"{company} [{rtype}, {row.sheet.title}]"
            match = EMAIL_RE.search(row.text("email"))
            email = norm_email(match.group(0)) if match else ""
            status = _norm(row.text("status"))
            call = bool(_phrase(row.text("user_notes"), call_phrases))
            last = row.last_contact_date()
            last_sent: MailInfo | None = None
            values: dict = {}
            auto_note = ""

            def update(values: dict | None = None, note: str = "") -> None:
                for r in rows:
                    changes.update(r, values, note)

            if not email and not call:
                if last and today >= (to_date(row.get("next_action_date")) or last + timedelta(days=days)):
                    report.add("Flagged - no email on row", label)
                    update(note="Due for follow-up but there's no email address on the row - add one, or put "
                                "'need to be called' in Notes for a call reminder")
                continue
            if email and ctx.state.is_suppressed(email):
                report.add("Unsubscribed", f"{label} - {email}")
                update({"status": cfg.status("unsubscribed")}, f"{email} is on the unsubscribe list")
                continue

            if email:
                sent = mailbox.messages_to(email, today - timedelta(days=fc["sent_lookback_days"]))
                last_sent = sent[0] if sent else None
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
                        if nxt is None or nxt <= sent_on:  # a Followup date still ahead of that email stays
                            values["next_action_date"] = sent_on + timedelta(days=days)
                        last = sent_on

            if last is None:
                report.add("Flagged - no date", f"{label}: no 'Date sent' / 'Last contacted' and nothing in Sent")
                update(values, "Open but no date sent - add one so follow-ups can be timed")
                continue

            if email:
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
                if auto := next((r for r in replies if r.auto_reply), None):
                    auto_note = f"Auto-reply only ({fmt_date(_local_date(auto, today))}: {auto.subject[:60]})"

            # 2. Due? Your Followup date wins; otherwise the date sent + 5 / 10 days.
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
            what = "Call reminder" if not email else "Follow-up"
            existing = ctx.state.draft(key, "followup", step)
            if existing is not None:
                # Drafted on an earlier run; make sure the sheet reflects it.
                made = date.fromisoformat(existing["created_on"])
                values.update(next_action_date=max(made, last) + timedelta(days=days), followup_count=step)
                if not call:
                    values["last_contacted"] = max(made, last)
                waiting = mailbox.draft_waiting(existing["message_id"], email)
                report.add("Flagged - follow-up already drafted", f"{label} (#{step}, "
                           f"{'still in Drafts' if waiting else 'drafted'} {fmt_date(made)})")
                update(values, f"{what} #{step} drafted {fmt_date(made)}"
                               + (" - still waiting in Drafts" if waiting else ""))
                continue
            if email and mailbox.draft_waiting(to_addr=email):
                report.add("Flagged - unsent draft waiting", label)
                update(values, "Due for follow-up, but a draft to this address is already in Drafts - "
                               "send or delete it first")
                continue
            if sender.problems or budget <= 0:
                why = "sender details missing" if sender.problems else "per-run draft cap reached"
                report.add("Flagged - due, not drafted", f"{label} ({why})")
                update(values, f"Due for {what.lower()} #{step} ({why})" + ("; your notes say call" if call else ""))
                continue

            threaded = bool(last_sent and last_sent.message_id)
            try:
                draft = ctx.agent.structured(
                    system=FOLLOWUP_SYSTEM,
                    prompt=followup_prompt(record_type=rtype, number=step, row_context=row.email_context(),
                                           original=(last_sent.text[:1500] if last_sent else ""),
                                           days_since=overdue, threaded=threaded, cfg=cfg, call=call,
                                           has_email=bool(email)),
                    schema=FOLLOWUP_SCHEMA, effort=fc["effort"])
            except AgentError as exc:
                report.add("Errors", f"{company}: follow-up drafting failed - {exc}")
                continue
            points = [clean(p) for p in draft.get("call_points") or [] if clean(p)][:5]
            block = _call_block(row, step=step, days_since=overdue, last=last, rtype=rtype, points=points,
                                has_email=bool(email)) if call else ""
            if email:
                to_addr, to_name = email, row.text("contact_name")
                subject = clean(draft["subject"])
                if threaded:
                    subject = (last_sent.subject if last_sent.subject.lower().startswith("re:")
                               else f"Re: {last_sent.subject}")
                body = sender.finish_body(cfg, draft["body"], record_type=rtype, source_url=row.text("source_url"))
                if threaded and last_sent.text:
                    quoted = "\n".join("> " + line for line in last_sent.text[:2000].splitlines())
                    when = fmt_date(_local_date(last_sent, today))
                    body = f"{body}\nOn {when}, {sender.name or sender.email} wrote:\n{quoted}\n"
                if call:
                    subject, body = CALL_TAG + subject, f"{block}\n\n{body}"
                unsubscribe = sender.unsubscribe_email if compliance.needs_unsubscribe(cfg, rtype) else ""
            else:
                # Nobody to email: a reminder to yourself, never sent to the contact.
                to_addr, to_name = sender.email, sender.name
                subject = f"[CALL] {company}: follow-up #{step} due"
                body, unsubscribe, threaded = block + "\n", "", False
            msg = build_message(
                from_name=sender.name, from_addr=sender.email, to_addr=to_addr, to_name=to_name,
                subject=subject, body=body,
                in_reply_to=last_sent.message_id if threaded else "",
                references=last_sent.references if threaded else "",
                unsubscribe_addr=unsubscribe, agent_tag=f"followup {step}; {row.ref}")
            if ctx.dry_run:
                ctx.preview_dir.mkdir(parents=True, exist_ok=True)
                prefix = ("call" if not email else "followup") + str(step)
                path = ctx.preview_dir / f"{prefix}_{key.split(':', 1)[-1].replace('/', '_')}.eml"
                path.write_bytes(msg.as_bytes())
                report.add("Would draft follow-up", f"{label}{' (CALL)' if call else ''} (preview: {path})")
                continue
            mailbox.append_draft(msg)
            ctx.state.record_draft(key, "followup", step, to_addr, subject, msg["Message-ID"], row.ref, today)
            budget -= 1
            drafted += 1
            values.update(next_action_date=today + timedelta(days=days), followup_count=step)
            if call:
                # You'll phone rather than send, so this isn't a contact date yet; sending the email later
                # is picked up from your Sent folder.
                phone = row.text("phone") or "no phone on the row"
                if email:
                    report.add("Call needed - follow-up drafted with [CALL FIRST]", f"{label}: {phone}")
                    note = f"Your notes say call: follow-up #{step} drafted with [CALL FIRST] and call points"
                else:
                    report.add("Call reminders drafted (addressed to you)", f"{label}: {phone}")
                    note = f"Your notes say call: call reminder #{step} drafted to you (no email on the row)"
            else:
                values["last_contacted"] = today
                report.add("Follow-ups drafted (in your Drafts folder)", f"{label}: #{step} \"{subject}\"")
                note = f"Follow-up #{step} drafted - review and send it from Drafts"
            report.count("follow-up drafts")
            update(values, note + (f"; {auto_note}" if auto_note else ""))
    finally:
        mailbox.close()

    commit(ctx, changes)
    return drafted
