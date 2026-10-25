from datetime import timedelta

import openpyxl

from origincrest_agent.followup import run_followup
from origincrest_agent.workbook import WorkbookStore

from .conftest import TODAY, mail, utc


def rfq_rows(cfg):
    ws = openpyxl.load_workbook(cfg.workbook_path("rfq"))["RFQ Tracker"]
    headers = [c.value for c in ws[3]]
    return {r[0].value: dict(zip(headers, (c.value for c in r))) for r in ws.iter_rows(min_row=4) if r[0].value}


def sent_rfq(mailbox, days_ago=7):
    msg = mail("Sent", utc(TODAY - timedelta(days=days_ago)), "hello@origincrest.com.au", "quotes@seaway.com.au",
               "RFQ: 1x20ft green coffee Belawan to Adelaide", text="Hi Jo, could you quote 1x20ft...",
               message_id="<rfq-1@origincrest.com.au>")
    mailbox.sent.append(msg)
    return msg


def test_quote_followup_threaded(cfg, agent, mailbox, make_ctx):
    sent_rfq(mailbox)
    ctx = make_ctx("followup")
    assert run_followup(ctx) == 1
    msg = mailbox.drafts[0]
    assert msg["To"] == "Jo Bloggs <quotes@seaway.com.au>"
    assert msg["Subject"] == "Re: RFQ: 1x20ft green coffee Belawan to Adelaide"
    assert msg["In-Reply-To"] == "<rfq-1@origincrest.com.au>"
    assert "List-Unsubscribe" not in msg  # RFQ follow-ups aren't marketing
    assert "> Hi Jo, could you quote" in msg.get_content()

    rows = rfq_rows(cfg)
    seaway = rows["Seaway Logistics"]
    assert seaway["Date Contacted"].date() == TODAY  # existing column reused
    assert seaway["Next action date"].date() == TODAY + timedelta(days=5)
    assert seaway["Follow-ups drafted"] == 1
    assert "Follow-up #1 drafted" in seaway["Agent notes"]
    assert seaway["Notes"] == "Asked for 1x20ft"  # user's notes untouched
    assert rows["Fastfreight Customs"]["Agent notes"] is None  # only 2 days, not due
    assert rows["Harbour 3PL"]["Agent notes"] is None  # not Open

    # Same day again: nothing new.
    assert run_followup(make_ctx("followup")) == 0
    assert len(mailbox.drafts) == 1


def test_reply_found_stops_followup(cfg, agent, mailbox, make_ctx):
    sent = sent_rfq(mailbox)
    mailbox.inbox.append(mail("INBOX", sent.when + timedelta(days=2), "jo@seaway.com.au", "hello@origincrest.com.au",
                              "Re: RFQ: 1x20ft green coffee", text="Quote attached."))
    run_followup(make_ctx("followup"))
    assert mailbox.drafts == []
    seaway = rfq_rows(cfg)["Seaway Logistics"]
    assert seaway["Status"] == "Replied - review"
    assert seaway["Last reply"].date() == (sent.when + timedelta(days=2)).date()


def test_auto_reply_is_not_a_reply(cfg, agent, mailbox, make_ctx):
    sent = sent_rfq(mailbox)
    mailbox.inbox.append(mail("INBOX", sent.when + timedelta(hours=1), "quotes@seaway.com.au",
                              "hello@origincrest.com.au", "Automatic reply: RFQ", auto=True))
    assert run_followup(make_ctx("followup")) == 1
    assert "Auto-reply only" in rfq_rows(cfg)["Seaway Logistics"]["Agent notes"]


def test_unsubscribe_is_honoured(cfg, agent, mailbox, make_ctx):
    sent = sent_rfq(mailbox)
    mailbox.inbox.append(mail("INBOX", sent.when + timedelta(days=1), "quotes@seaway.com.au",
                              "hello@origincrest.com.au", "Re: RFQ", text="Please remove me from your list."))
    ctx = make_ctx("followup")
    run_followup(ctx)
    assert mailbox.drafts == []
    assert rfq_rows(cfg)["Seaway Logistics"]["Status"] == "Unsubscribed - do not contact"
    assert ctx.state.is_suppressed("quotes@seaway.com.au")


def test_lead_intro_sent_becomes_open_and_uses_10_days(cfg, agent, mailbox, make_ctx):
    store = WorkbookStore("rfq", cfg.workbook_path("rfq"), cfg).load()
    view = store.get_or_create("Roasters - Adelaide", "roaster")
    view.append({"company": "Hills Roasting Co", "email": "info@hillsroasting.com.au",
                 "status": "Intro drafted - to review", "draft_created": TODAY - timedelta(days=12),
                 "record_type": "Lead"})
    store.save(cfg.data_dir / "backups")
    mailbox.sent.append(mail("Sent", utc(TODAY - timedelta(days=11)), "hello@origincrest.com.au",
                             "info@hillsroasting.com.au", "Indonesian green for your range", text="Hi there"))
    run_followup(make_ctx("followup"))
    ws = openpyxl.load_workbook(cfg.workbook_path("rfq"))["Roasters - Adelaide"]
    headers = [c.value for c in ws[1]]
    row = dict(zip(headers, (c.value for c in ws[2])))
    assert row["Status"] == "Open"
    # 11 days since the intro went out > 10-day lead window, so a follow-up is drafted with the unsubscribe footer.
    lead_draft = next(m for m in mailbox.drafts if "hillsroasting" in m["To"])
    assert "unsubscribe" in lead_draft.get_content()
    assert lead_draft["List-Unsubscribe"]
    assert row["Next action date"].date() == TODAY + timedelta(days=10)


def test_max_followups_flagged(cfg, agent, mailbox, make_ctx):
    sent_rfq(mailbox)
    store = WorkbookStore("rfq", cfg.workbook_path("rfq"), cfg).load()
    next(store.view("RFQ Tracker").rows()).set("followup_count", 3)
    store.save(cfg.data_dir / "backups")
    ctx = make_ctx("followup")
    run_followup(ctx)
    assert mailbox.drafts == []
    assert "Flagged - max follow-ups reached" in ctx.report.sections


def test_failed_sheet_write_does_not_duplicate_draft(cfg, agent, mailbox, make_ctx):
    sent_rfq(mailbox)
    rfq = cfg.workbook_path("rfq")
    owner = rfq.parent / ("~$" + rfq.name)
    owner.write_text("Excel")
    run_followup(make_ctx("followup"))
    assert len(mailbox.drafts) == 1
    assert rfq_rows(cfg)["Seaway Logistics"]["Status"] == "Open"  # not written: Excel had it open
    owner.unlink()
    run_followup(make_ctx("followup"))
    assert len(mailbox.drafts) == 1  # no second draft
    seaway = rfq_rows(cfg)["Seaway Logistics"]
    assert seaway["Follow-ups drafted"] == 1
    assert seaway["Next action date"].date() == TODAY + timedelta(days=5)
