"""Supplier agent: search, compare, enquiry drafts, follow-ups - and its separation from the outreach agent."""

import hashlib
import imaplib
import os
import time
import uuid
from datetime import date, timedelta
from email.utils import format_datetime
from pathlib import Path

import openpyxl
import pytest
import yaml

from origincrest_agent import cli as outreach_cli
from origincrest_agent.config import DEFAULTS, load_config
from origincrest_agent.followup import run_followup
from origincrest_agent.leads import run_leads
from origincrest_agent.llm import ResearchResult
from origincrest_agent.mailer import Mailbox, quote
from origincrest_agent.runtime import Report, RunContext
from origincrest_agent.state import State
from origincrest_agent.supplier_agent import main as supplier_main
from origincrest_agent.supplier_config import load_supplier_config
from origincrest_agent.suppliers import (FOCUS, dedupe_name, run_focus, run_supplier_drafts, run_supplier_followup,
                                         run_suppliers)

from .conftest import TODAY, FakeVerifier, mail, utc

TAB = "New Suppliers - Indonesia"
SITE = "https://gayomountain.co.id"
PAGES = {SITE + "/contact": "Export desk: export@gayomountain.co.id Tel +62 643 555 0101 WhatsApp +62 812 5550 0102"}
VENDOR_PROFILE = "Robusta Dampit G1 FOB USD 3.21/kg, priced per full 20ft; 70/30 payment"
DEMAND = "Hills Roasters: LCL trial of 600 kg, then one 20ft FCL per quarter"


def supplier(name="Koperasi Gayo Mountain", email="export@gayomountain.co.id", score=82, **kw):
    data = {"company_name": name, "business_type": "Cooperative (exports directly)", "website": SITE,
            "contact_name": "", "contact_role": "", "email": email, "phone": "+62 643 555 0101",
            "whatsapp": "+62 812 5550 0102", "city": "Takengon", "province": "Aceh",
            "origins": "Gayo arabica G1, wet-hulled and washed", "certifications": "Organic, Fairtrade",
            "export_capability": "Exports FCL from Belawan; LCL on request; samples by courier",
            "covers_current_demand": "Yes",
            "comparison_with_current_vendor": "Certified and publishes export contacts; prices not public.",
            "fit_score": score, "fit_rationale": "Certified Gayo lots with an export licence.",
            "suggested_approach": "Email the export desk, then WhatsApp the published number.",
            "source_url": SITE + "/contact", "notes": ""}
    data.update(kw)
    return data


# --- fixtures: an outreach config.yaml and a separate config.suppliers.yaml ------------------------------

@pytest.fixture
def files(tmp_path, cfg):
    """config.yaml as the outreach agent uses it (incl. settings the supplier agent must NOT pick up), and
    config.suppliers.yaml for the supplier agent."""
    outreach = {
        "data_dir": str(tmp_path / "data"),
        "workbooks": {"rfq": {"path": str(cfg.workbook_path("rfq"))}, "master": {"path": str(cfg.workbook_path("master"))}},
        "email": {"imap_host": "mail.example.com", "username": "hello@origincrest.com.au", "drafts_folder": "Drafts"},
        "sender": dict(cfg["sender"]),
        "claude": {"model": "claude-test-model"},
        "followup": {"quote_days": 99, "lead_days": 99, "max_drafts_per_run": 1},
        "leads": {"max_drafts_per_run": 1, "extra_tabs": ["Some outreach tab"]},
        "excel": {"lock_wait_minutes": 0},
    }
    suppliers = {
        "suppliers": {"workbook_path": str(tmp_path / "suppliers" / "suppliers.xlsx"), "focus_per_run": 1,
                      "baseline_vendor": {"name": "PT Baseline Kopi", "profile": VENDOR_PROFILE}},
        "sourcing": {"demand": DEMAND},
        "excel": {"lock_wait_minutes": 0},
    }
    shared_path, own_path = tmp_path / "config.yaml", tmp_path / "config.suppliers.yaml"
    shared_path.write_text(yaml.safe_dump(outreach), encoding="utf-8")
    own_path.write_text(yaml.safe_dump(suppliers), encoding="utf-8")
    return own_path, shared_path


@pytest.fixture
def scfg(files):
    return load_supplier_config(*files)


@pytest.fixture
def make_sctx(scfg, agent, mailbox):
    def _make(workflow="supplier-test", dry_run=False, today=TODAY, cfg=None):
        c = cfg or scfg
        return RunContext(cfg=c, state=State(c.data_dir / "state.db"), today=today, report=Report(workflow, today),
                          dry_run=dry_run, agent_factory=lambda: agent, mailbox_factory=lambda: mailbox,
                          verifier=FakeVerifier(PAGES))
    return _make


def sheet(cfg):
    ws = openpyxl.load_workbook(cfg.workbook_path("suppliers"))[TAB]
    headers = [c.value for c in ws[1]]
    return {r[0].value: dict(zip(headers, (c.value for c in r))) for r in ws.iter_rows(min_row=2)}


def digest(*paths: Path) -> dict:
    return {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in paths if p.exists()}


# --- search rotation ---------------------------------------------------------------------------------------

def test_focus_rotates_every_run(files):
    cfg = load_supplier_config(*files)
    cfg.data["suppliers"]["focus_per_run"] = 2
    runs = -(-len(FOCUS) // 2)  # enough runs to cover the list once
    seen = [tuple(run_focus(cfg, date(2026, 11, 12) + timedelta(days=3 * i))) for i in range(runs)]
    assert len(set(seen)) == len(seen)
    assert {f for run in seen for f in run} == set(FOCUS)
    # every species and the main growing areas are in the rotation
    text = " ".join(FOCUS)
    for word in ("Preanger", "Gunung Kawi", "Dampit", "Temanggung", "Garut", "Gayo", "Toraja", "Lampung",
                 "Liberica", "Flores", "Bali", "Papua"):
        assert word in text, word
    assert run_focus(cfg, date(2026, 11, 12)) == run_focus(cfg, date(2026, 11, 12))


def test_priority_focus_is_searched_every_run(scfg):
    urgent = ["Dampit robusta G1 natural exporters", "Temanggung robusta G1 exporters"]
    scfg.data["suppliers"].update(priority_focus=urgent, focus_per_run=2)
    runs = [run_focus(scfg, date(2026, 11, 13) + timedelta(days=3 * i)) for i in range(4)]
    assert all(len(r) == 2 and r[0] in urgent and r[1] in FOCUS for r in runs)
    assert {r[0] for r in runs} == set(urgent)          # alternates between the urgent searches
    assert len({r[1] for r in runs}) == 4               # while the general rotation keeps moving


def test_indonesian_company_forms_are_ignored_for_duplicates():
    assert dedupe_name("PT. Kopi Nusantara Jaya") == "Kopi Nusantara Jaya"
    assert dedupe_name("CV Kopi Gayo, Tbk") == "Kopi Gayo"
    assert dedupe_name("Ptolemy Coffee") == "Ptolemy Coffee"


# --- search, compare, draft, follow up -------------------------------------------------------------------

def test_discovers_compares_and_drafts_enquiry(scfg, agent, mailbox, make_sctx):
    agent.research_queue = [ResearchResult(data={"suppliers": [
        supplier(),
        supplier("PT. Baseline Kopi", "sales@baselinekopi.example", website="", source_url=""),  # current vendor
        supplier("CV Unverified Kopi", "made-up@nowhere.example", website="https://unverified.example",
                 source_url="https://unverified.example/contact", whatsapp="", phone=""),
    ]})]
    agent.structured_queue = [{"subject": "Green coffee import enquiry from Australia - Gayo arabica",
                               "body": "Dear Koperasi Gayo Mountain team,\n\nWe are ready to import.\n\nKind regards,"}]
    ctx = make_sctx("supplier-search")
    assert run_suppliers(ctx) == 2
    assert run_supplier_drafts(ctx) == 1

    search_prompt = agent.calls[0][1]
    assert "PT Baseline Kopi" in search_prompt and VENDOR_PROFILE in search_prompt and DEMAND in search_prompt
    rows = sheet(scfg)
    assert set(rows) == {"Koperasi Gayo Mountain", "CV Unverified Kopi"}
    assert any("Baseline Kopi" in line for line in ctx.report.sections["Skipped duplicates"])
    row = rows["Koperasi Gayo Mountain"]
    assert row["Email"] == "export@gayomountain.co.id" and row["State"] == "Aceh"
    assert row["Phone"] == "+62 643 555 0101 / WhatsApp +62 812 5550 0102"
    assert row["Comparison with current vendor"].startswith("Certified")
    assert row["Suggested approach"] and row["Fit score (0-100)"] == 82
    assert row["Status"] == "Enquiry drafted - to review" and row["Record type"] == "Supplier"
    assert rows["CV Unverified Kopi"]["Email"] is None             # not on the supplier's pages: dropped
    assert rows["CV Unverified Kopi"]["Status"] == "No public email"

    [msg] = mailbox.drafts
    assert msg["To"] == "export@gayomountain.co.id" and msg["X-OriginCrest-Agent"].startswith("supplier;")
    # the same one-line opt-out as RFQs, not the marketing footer
    assert msg["List-Unsubscribe"] == "<mailto:hello@origincrest.com.au?subject=unsubscribe>"
    assert "reply \"unsubscribe\" and we won't contact you again" in msg.get_content()
    assert "within 5 business days" not in msg.get_content()
    assert "Swami K" in msg.get_content()
    email_prompt = next(p for kind, p in agent.calls if kind == "structured")
    assert DEMAND in email_prompt and "already arranged" in email_prompt
    # Private assessment and everything about the current vendor stay out of the email.
    for private in ("Certified and publishes", "Email the export desk", "Baseline Kopi", "3.21"):
        assert private not in email_prompt

    # The next run finds nothing new and drafts nothing twice.
    agent.research_queue = [ResearchResult(data={"suppliers": [supplier()]})]
    ctx = make_sctx("supplier-search")
    assert run_suppliers(ctx) == 0
    assert run_supplier_drafts(ctx) == 0
    assert len(mailbox.drafts) == 1


def test_blank_demand_and_low_fit(scfg, agent, mailbox, make_sctx):
    scfg.data["sourcing"]["demand"] = ""
    agent.research_queue = [ResearchResult(data={"suppliers": [supplier(score=35)]})]
    ctx = make_sctx()
    assert run_suppliers(ctx) == 1
    assert "sourcing.demand is blank" in ctx.report.sections["Demand"][0]
    assert run_supplier_drafts(ctx) == 0
    assert sheet(scfg)["Koperasi Gayo Mountain"]["Status"] == "Researched - low fit"
    assert not mailbox.drafts


def test_followup_after_three_days(scfg, agent, mailbox, make_sctx):
    agent.research_queue = [ResearchResult(data={"suppliers": [supplier()]})]
    ctx = make_sctx()
    run_suppliers(ctx)
    run_supplier_drafts(ctx)
    draft = mailbox.drafts.pop()  # you send it the same day
    mailbox.sent.append(mail("Sent", utc(TODAY, 10), "hello@origincrest.com.au", "export@gayomountain.co.id",
                             draft["Subject"], "Dear team,\n\nWe are ready to import.", "<first@origincrest>"))

    assert run_supplier_followup(make_sctx("supplier-followup", today=TODAY + timedelta(days=2))) == 0
    assert sheet(scfg)["Koperasi Gayo Mountain"]["Status"] == "Open"
    agent.structured_queue = [{"subject": "x", "body": "Dear team,\n\nFollowing up.\n\nKind regards,",
                               "call_points": []}]
    # 3 days (the supplier setting), not config.yaml's 99-day follow-up settings
    assert run_supplier_followup(make_sctx("supplier-followup", today=TODAY + timedelta(days=3))) == 1
    [msg] = mailbox.drafts
    assert msg["Subject"].startswith("Re: ") and msg["In-Reply-To"] == "<first@origincrest>"
    prompt = agent.calls[-1][1]
    assert "(supplier)" in prompt and DEMAND in prompt
    for private in ("Certified and publishes", "Baseline Kopi", "3.21"):
        assert private not in prompt
    row = sheet(scfg)["Koperasi Gayo Mountain"]
    assert row["Follow-ups drafted"] == 1
    assert row["Next action date"].date() == TODAY + timedelta(days=6)


def test_dry_run_writes_nothing(scfg, agent, mailbox, make_sctx):
    agent.research_queue = [ResearchResult(data={"suppliers": [supplier()]})]
    ctx = make_sctx(dry_run=True)
    assert run_suppliers(ctx) == 1
    assert run_supplier_drafts(ctx) == 0
    assert not scfg.workbook_path("suppliers").exists() and not mailbox.drafts
    assert "Koperasi Gayo Mountain" in ctx.report.sections["Would add"][0]
    assert "covers current demand: Yes" in ctx.report.sections["Would add"][0]
    assert "isn't created in a dry run" in ctx.report.sections["Nothing to draft"][0]


def test_every_search_is_reported_with_its_reasons(scfg, agent, make_sctx):
    scfg.data["suppliers"].update(priority_focus=["Dampit robusta G1 natural exporters"], focus_per_run=2)
    agent.research_queue = [
        ResearchResult(data={"suppliers": [], "search_notes": "Only marketplace listings, no identifiable exporter."}),
        ResearchResult(data={"suppliers": [supplier()], "search_notes": "Searched in Indonesian too."}),
    ]
    ctx = make_sctx(dry_run=True)
    assert run_suppliers(ctx) == 1
    first, second = ctx.report.sections["Searches"]
    assert first.startswith("Dampit robusta G1 natural exporters: 0 found, 0 new. Only marketplace listings")
    assert ": 1 found, 1 new. Searched in Indonesian too." in second


def test_every_supplier_gets_the_current_demand_and_is_asked_to_quote(scfg, agent, mailbox, make_sctx):
    """The demand never filters suppliers: one without the customer's coffees is still drafted, and every
    enquiry pitches OriginCrest's experience, leads with the current order and asks for their quotation."""
    scfg.data["sourcing"]["experience"] = "OriginCrest prices and quotes Java coffees to Australian roasters."
    agent.research_queue = [ResearchResult(data={"suppliers": [
        supplier(covers_current_demand="No", origins="Kerinci arabica, washed", score=55)]})]
    ctx = make_sctx()
    run_suppliers(ctx)
    assert run_supplier_drafts(ctx) == 1
    assert sheet(scfg)["Koperasi Gayo Mountain"]["Covers current demand"] == "No"
    prompt = agent.calls[-1][1]
    assert DEMAND in prompt and "Experience: OriginCrest prices and quotes Java coffees" in prompt
    assert "your quotation: FOB price per kg" in prompt and "your current offer list" in prompt
    assert '"Covers current demand": "No"' in prompt
    for private in ("Baseline Kopi", "3.21"):
        assert private not in prompt
    from origincrest_agent.supplier_prompts import SUPPLIER_INTRO_SYSTEM, SUPPLIER_SYSTEM
    assert "The current demand - the heart of the email, for every supplier" in SUPPLIER_INTRO_SYSTEM
    assert "closest equivalents" in SUPPLIER_INTRO_SYSTEM
    assert "leave a supplier out, or score it lower, because it doesn't sell those coffees" in SUPPLIER_SYSTEM
    assert "(not to the current order)" in SUPPLIER_SYSTEM


# --- separation from the outreach agent ------------------------------------------------------------------

def test_supplier_settings_take_only_the_mailbox_from_config_yaml(files, tmp_path):
    cfg = load_supplier_config(*files)
    # shared, read from config.yaml
    assert cfg["email"]["imap_host"] == "mail.example.com" and cfg["claude"]["model"] == "claude-test-model"
    assert cfg["sender"]["name"] == "Swami K"
    assert cfg.shared_used == ["claude", "email", "sender"]
    # not shared: config.yaml's follow-up / leads / workbooks / data folder don't apply
    assert cfg["followup"]["quote_days"] == DEFAULTS["followup"]["quote_days"]
    assert cfg["followup"]["max_drafts_per_run"] == 15 and cfg["followup"]["supplier_days"] == 3
    assert cfg["leads"]["max_drafts_per_run"] == DEFAULTS["leads"]["max_drafts_per_run"]
    assert cfg["leads"]["extra_tabs"] == [TAB]
    assert set(cfg["workbooks"]) == {"suppliers"} and cfg["discovery"]["profiles"] == {}
    assert cfg.data_dir == tmp_path / "data" / "suppliers"
    # the supplier file can override a shared setting
    own_path, shared_path = files
    data = yaml.safe_load(own_path.read_text())
    data["email"] = {"drafts_folder": "Supplier drafts"}
    own_path.write_text(yaml.safe_dump(data))
    cfg = load_supplier_config(own_path, shared_path)
    assert cfg["email"]["drafts_folder"] == "Supplier drafts" and cfg["email"]["imap_host"] == "mail.example.com"


def test_supplier_sections_left_in_config_yaml_are_ignored_and_flagged(files):
    own_path, shared_path = files
    data = yaml.safe_load(shared_path.read_text())
    data["suppliers"] = {"baseline_vendor": {"name": "Pasted into the wrong file"}}
    data["sourcing"] = {"demand": "wrong file"}
    shared_path.write_text(yaml.safe_dump(data))
    cfg = load_supplier_config(own_path, shared_path)
    assert cfg["suppliers"]["baseline_vendor"]["name"] == "PT Baseline Kopi"
    assert cfg["sourcing"]["demand"] == DEMAND
    assert cfg.ignored_in_shared == ["suppliers", "sourcing"]
    # ... and the outreach agent doesn't use them either
    outreach = load_config(shared_path)
    assert "suppliers" not in outreach["workbooks"] and outreach.lead_tabs("master") == {
        "New Leads - Adelaide", "New Leads - Other States", "Some outreach tab"}


def test_outreach_defaults_have_no_supplier_settings():
    assert "suppliers" not in DEFAULTS and "sourcing" not in DEFAULTS
    assert set(DEFAULTS["workbooks"]) == {"rfq", "master"}
    assert "supplier_days" not in DEFAULTS["followup"] and "supplier_drafted" not in DEFAULTS["statuses"]


def test_supplier_runs_never_change_config_files_or_outreach_workbooks(files, cfg, scfg, agent, mailbox, make_sctx):
    own_path, shared_path = files
    watched = [own_path, shared_path, cfg.workbook_path("rfq"), cfg.workbook_path("master")]
    before = digest(*watched)
    agent.research_queue = [ResearchResult(data={"suppliers": [supplier()]})]
    ctx = make_sctx()
    run_suppliers(ctx)
    run_supplier_drafts(ctx)
    run_supplier_followup(make_sctx(today=TODAY + timedelta(days=5)))
    assert digest(*watched) == before
    # its own state, reports and backups live in its own folder
    assert (scfg.data_dir / "state.db").exists() and scfg.data_dir.parent == cfg.data_dir.parent / "data"
    assert not list((scfg.data_dir / "backups").glob("rfq*")) and not list((scfg.data_dir / "backups").glob("master*"))


def test_outreach_runs_leave_the_supplier_workbook_alone(cfg, scfg, agent, mailbox, make_sctx, make_ctx):
    agent.research_queue = [ResearchResult(data={"suppliers": [supplier()]})]
    ctx = make_sctx()
    run_suppliers(ctx)
    run_supplier_drafts(ctx)
    drafts_before = len(mailbox.drafts)
    supplier_file = scfg.workbook_path("suppliers")
    before = digest(supplier_file)
    run_leads(make_ctx("leads"))
    run_followup(make_ctx("followup"))
    assert digest(supplier_file) == before
    assert not any("gayomountain" in m["To"] for m in mailbox.drafts[drafts_before:])
    # and the outreach state database knows nothing about suppliers
    state = State(cfg.data_dir / "state.db")
    assert state.db.execute("SELECT COUNT(*) FROM drafts WHERE workflow='supplier'").fetchone()[0] == 0


def test_supplier_cli(files, capsys):
    own_path, shared_path = files
    data = yaml.safe_load(shared_path.read_text())
    data["sourcing"] = {"demand": "pasted in the wrong file"}
    shared_path.write_text(yaml.safe_dump(data))
    assert supplier_main(["--config", str(own_path), "--shared-config", str(shared_path), "check"]) == 1  # no IMAP
    out = capsys.readouterr().out
    assert f"[OK] Supplier agent settings: {own_path}" in out
    assert "(read only): claude, email, sender" in out
    assert "config.yaml has a 'sourcing:' section" in out
    assert "[OK] compared with PT Baseline Kopi" in out and "[OK] demand in the enquiries" in out
    # search / followup refuse to run without the supplier settings file
    assert supplier_main(["--config", str(own_path.with_name("missing.yaml")), "--shared-config", str(shared_path),
                          "search"]) == 2
    # the outreach CLI has no supplier commands
    with pytest.raises(SystemExit):
        outreach_cli.main(["suppliers"])


# --- the real mailbox (Dovecot or similar), like tests/test_imap_integration.py -------------------------

SERVER = os.environ.get("OC_TEST_IMAP")


@pytest.mark.skipif(not SERVER, reason="set OC_TEST_IMAP=host:port:user:password")
def test_enquiry_and_followup_drafts_in_a_real_mailbox(scfg, agent, make_sctx):
    host, port, user, password = SERVER.split(":", 3)
    tag = uuid.uuid4().hex[:8]
    domain = f"kopi{tag}.co.id"
    site = f"https://{domain}"
    contact = f"export@{domain}"

    def mbox():
        return Mailbox(host, int(port), user, password, security="none")

    agent.research_queue = [ResearchResult(data={"suppliers": [
        supplier(f"PT Kopi {tag}", contact, website=site, source_url=site + "/contact", phone="", whatsapp="")]})]
    agent.structured_queue = [{"subject": f"Green coffee import enquiry {tag}",
                               "body": "Dear team,\n\nWe are ready to import.\n\nKind regards,"}]
    ctx = make_sctx()
    ctx.mailbox_factory = mbox
    ctx.verifier = FakeVerifier({site + "/contact": f"Export desk: {contact}"})
    run_suppliers(ctx)
    assert run_supplier_drafts(ctx) == 1

    with mbox() as mb:
        [draft] = [d for d in mb.agent_drafts() if d.subject == f"Green coffee import enquiry {tag}"]
        assert draft.to_addrs == [contact]
        # "Send" it: the draft leaves Drafts and a copy lands in Sent, dated today
        uids = mb._search(mb.drafts, ["TO", quote(contact)])
        raw = mb.imap.uid("FETCH", uids[0], "(BODY.PEEK[])")[1][0][1]
    imap = imaplib.IMAP4(host, int(port))
    imap.login(user, password)
    imap.select(quote("Drafts"))
    imap.uid("STORE", uids[0], "+FLAGS", r"(\Deleted)")
    imap.expunge()
    sent_when = utc(TODAY, 10)
    raw = raw.replace(b"\r\nDate: ", b"\r\nX-Old-Date: ", 1).replace(
        b"\r\nX-Old-Date: ", ("\r\nDate: " + format_datetime(sent_when) + "\r\nX-Old-Date: ").encode(), 1)
    imap.append(quote("Sent Mail"), r"(\Seen)", imaplib.Time2Internaldate(sent_when.timestamp()), raw)
    imap.logout()
    time.sleep(0.2)

    agent.structured_queue = [{"subject": "x", "body": "Dear team,\n\nFollowing up.\n\nKind regards,",
                               "call_points": []}]
    follow = make_sctx("supplier-followup", today=TODAY + timedelta(days=3))
    follow.mailbox_factory = mbox
    assert run_supplier_followup(follow) == 1
    with mbox() as mb:
        uids = mb._search(mb.drafts, ["TO", quote(contact)])
        [reply] = mb._fetch(mb.drafts, uids)
    assert reply.subject == f"Re: Green coffee import enquiry {tag}"
    assert draft.message_id in reply.references
    assert sheet(scfg)[f"PT Kopi {tag}"]["Follow-ups drafted"] == 1


def test_unsubscribe_reply_stops_all_contact(scfg, agent, mailbox, make_sctx):
    agent.research_queue = [ResearchResult(data={"suppliers": [supplier()]})]
    ctx = make_sctx()
    run_suppliers(ctx)
    run_supplier_drafts(ctx)
    draft = mailbox.drafts.pop()
    mailbox.sent.append(mail("Sent", utc(TODAY, 10), "hello@origincrest.com.au", "export@gayomountain.co.id",
                             draft["Subject"], "Dear team", "<first@origincrest>"))
    mailbox.inbox.append(mail("INBOX", utc(TODAY + timedelta(days=1), 9), "export@gayomountain.co.id",
                              "hello@origincrest.com.au", "Re: " + draft["Subject"], "unsubscribe"))
    assert run_supplier_followup(make_sctx("supplier-followup", today=TODAY + timedelta(days=4))) == 0
    assert sheet(scfg)["Koperasi Gayo Mountain"]["Status"] == "Unsubscribed - do not contact"
    assert State(scfg.data_dir / "state.db").is_suppressed("export@gayomountain.co.id")
    assert not mailbox.drafts
