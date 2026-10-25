from datetime import timedelta

import openpyxl

from origincrest_agent.leads import run_leads
from origincrest_agent.llm import ResearchResult
from origincrest_agent.workbook import WorkbookStore

from .conftest import TODAY, mail, utc


def seed_roasters(cfg, names=("rfq", "master")):
    rows = [
        {"company": "Hills Roasting Co", "email": "info@hillsroasting.com.au", "website": "https://hillsroasting.com.au",
         "status": "New - to review", "source_url": "https://hillsroasting.com.au/contact",
         "added_by": "OriginCrest Agent"},
        {"company": "Low Fit Cafe", "email": "hi@lowfit.com.au", "status": "New - to review",
         "added_by": "OriginCrest Agent"},
        {"company": "No Email Roasters", "website": "https://noemail.com.au", "status": ""},
        {"company": "Existing Customer Roasters", "email": "buy@customer.com.au", "status": "Customer"},
    ]
    for name in names:
        store = WorkbookStore(name, cfg.workbook_path(name), cfg).load()
        view = store.get_or_create("Roasters - Adelaide", "roaster")
        for r in rows:
            view.append(r)
        store.save(cfg.data_dir / "backups")


def research(name, score, email="", **kw):
    data = {"company_name": name, "website": "", "city": "Adelaide", "state": "SA",
            "what_they_roast": "Light-roast single origins, filter focus", "size_estimate": "Small",
            "size_evidence": "Two staff, 15kg roaster", "sourcing": "Buys from two importers",
            "current_indonesian_coffees": "", "wholesale": "Yes", "public_email": email,
            "public_email_source_url": "", "contact_name": "", "contact_role": "", "fit_score": score,
            "best_fit_origins": ["Sumatra", "Flores"], "fit_rationale": "Single-origin focus, no Indonesian lot.",
            "personalisation_hooks": [{"fact": "Sells a washed Ethiopian on filter", "source_url": "https://x"}],
            "red_flags": ""}
    data.update(kw)
    return ResearchResult(data=data)


def sheet(cfg, name="rfq", tab="Roasters - Adelaide"):
    ws = openpyxl.load_workbook(cfg.workbook_path(name))[tab]
    headers = [c.value for c in ws[1]]
    return {r[0].value: dict(zip(headers, (c.value for c in r))) for r in ws.iter_rows(min_row=2)}


def test_research_score_and_draft(cfg, agent, mailbox, make_ctx):
    seed_roasters(cfg)
    # Order: user-entered rows first, then agent rows.
    agent.research_queue = [research("No Email Roasters", 70), research("Hills Roasting Co", 82),
                            research("Low Fit Cafe", 30)]
    ctx = make_ctx("leads")
    assert run_leads(ctx) == 1
    assert len([c for c in agent.calls if c[0] == "research"]) == 3  # once per company, not per workbook

    assert len(mailbox.drafts) == 1
    msg = mailbox.drafts[0]
    assert msg["To"] == "info@hillsroasting.com.au"
    assert msg["List-Unsubscribe"] == "<mailto:hello@origincrest.com.au?subject=unsubscribe>"
    body = msg.get_content()
    assert "Swami K" in body and "1 Test St" in body
    assert "unsubscribe" in body and "hillsroasting.com.au" in body

    for name in ("rfq", "master"):
        rows = sheet(cfg, name)
        hills = rows["Hills Roasting Co"]
        assert hills["Status"] == "Intro drafted - to review"
        assert hills["Fit score (0-100)"] == 82
        assert hills["Best-fit origins"] == "Sumatra, Flores"
        assert hills["Intro draft subject"] == "Indonesian green for your range"
        assert rows["Low Fit Cafe"]["Status"] == "Researched - low fit"
        assert rows["No Email Roasters"]["Status"] == "Researched - no public email"
        assert rows["Existing Customer Roasters"]["Status"] == "Customer"
        assert rows["Existing Customer Roasters"]["Fit score (0-100)"] is None

    # Nothing left to do on a second run: no new API calls, no new drafts.
    calls = len(agent.calls)
    assert run_leads(make_ctx("leads")) == 0
    assert len(agent.calls) == calls and len(mailbox.drafts) == 1


def test_approved_status_overrides_low_score(cfg, agent, mailbox, make_ctx):
    seed_roasters(cfg, names=("rfq",))
    cfg.data["leads"]["workbooks"] = ["rfq"]
    store = WorkbookStore("rfq", cfg.workbook_path("rfq"), cfg).load()
    view = store.view("Roasters - Adelaide")
    for row in view.rows():
        row.set("status", "Approved" if row.text("company") == "Low Fit Cafe" else "Customer")
    store.save(cfg.data_dir / "backups")
    agent.research_queue = [research("Low Fit Cafe", 30)]
    run_leads(make_ctx("leads"))
    assert [m["To"] for m in mailbox.drafts] == ["hi@lowfit.com.au"]


def test_no_drafts_until_sender_details_filled(cfg, agent, mailbox, make_ctx):
    seed_roasters(cfg, names=("rfq",))
    cfg.data["leads"]["workbooks"] = ["rfq"]
    cfg.data["sender"]["address"] = "<<EDIT ME: business postal address>>"
    cfg.data["sender"]["phone"] = ""
    agent.research_queue = [research("No Email Roasters", 70), research("Hills Roasting Co", 82),
                            research("Low Fit Cafe", 30)]
    ctx = make_ctx("leads")
    run_leads(ctx)
    assert mailbox.drafts == []
    assert "sender.address" in " ".join(ctx.report.sections["Errors"])
    assert sheet(cfg)["Hills Roasting Co"]["Researched on"] is None  # picked up again once fixed


def test_already_emailed_contact_goes_to_open(cfg, agent, mailbox, make_ctx):
    seed_roasters(cfg, names=("rfq",))
    cfg.data["leads"]["workbooks"] = ["rfq"]
    sent_day = TODAY - timedelta(days=30)
    mailbox.sent.append(mail("Sent", utc(sent_day), "hello@origincrest.com.au", "info@hillsroasting.com.au",
                             "Coffee chat"))
    agent.research_queue = [research("No Email Roasters", 70), research("Hills Roasting Co", 82),
                            research("Low Fit Cafe", 30)]
    run_leads(make_ctx("leads"))
    hills = sheet(cfg)["Hills Roasting Co"]
    assert mailbox.drafts == []
    assert hills["Status"] == "Open"
    assert hills["Last contacted"].date() == sent_day


def test_draft_cap_defers_rest(cfg, agent, mailbox, make_ctx):
    seed_roasters(cfg, names=("rfq",))
    cfg.data["leads"]["workbooks"] = ["rfq"]
    agent.research_queue = [research("No Email Roasters", 70, email=""), research("Hills Roasting Co", 82),
                            research("Low Fit Cafe", 90)]
    ctx = make_ctx("leads")
    run_leads(ctx, max_drafts=1)
    assert len(mailbox.drafts) == 1
    assert "Deferred (draft cap)" in ctx.report.sections
    agent.research_queue = []  # research is cached: the deferred row needs no new research
    run_leads(make_ctx("leads"), max_drafts=5)
    assert len(mailbox.drafts) == 2
