"""Supplier onboarding: read what suppliers sent, fill the onboarding workbook, draft questionnaires."""

import email
import email.policy
import imaplib
import io
import os
import uuid
import zipfile
from datetime import timedelta
from email.message import EmailMessage
from email.utils import format_datetime, make_msgid
from pathlib import Path

import openpyxl
import pytest
import yaml

from origincrest_agent.mailer import Mailbox, quote
from origincrest_agent.onboarding import (EXTRACT_SYSTEM, file_content, onboarding_folder, questionnaire_xlsx,
                                          run_onboarding, workbook_path)
from origincrest_agent.onboarding_questions import KEYS
from origincrest_agent.runtime import Report, RunContext
from origincrest_agent.state import State
from origincrest_agent.supplier_agent import main as supplier_main
from origincrest_agent.supplier_config import load_supplier_config
from origincrest_agent.suppliers import ensure_workbook

from .conftest import TODAY, FakeVerifier, utc

SUPPLIER = "PT Kopi Test"
ANSWERS = {
    "answers": [
        {"key": "A1", "answer": "PT Kopi Test (brand: Kopi Tengger)", "source": "profile.docx", "status": "Provided",
         "note": ""},
        {"key": "C1", "answer": "Robusta G1 natural: FOB USD 5.10/kg for 1 MT LCL", "source": "price list.pdf",
         "status": "Provided", "note": ""},
        {"key": "C4", "answer": "50% deposit", "source": "email 20/10/2026", "status": "Partly",
         "note": "Balance timing not stated"},
        {"key": "B1", "answer": "Dampit robusta", "source": "OriginCrest research", "status": "From research - to confirm",
         "note": ""},
    ],
    "red_flags": ["Asks for the deposit to a personal account"],
    "summary": "Prices and company details received; registration and documents still to come.",
}


def docx_bytes(text: str) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as zf:
        zf.writestr("word/document.xml", "<w:document><w:body><w:p><w:r><w:t>" + text
                    + "</w:t></w:r></w:p></w:body></w:document>")
    return out.getvalue()


def xlsx_bytes(rows) -> bytes:
    wb = openpyxl.Workbook()
    for r in rows:
        wb.active.append(r)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def raw_email(frm, subject, body, when, attachments=(), msgid=None, headers=None) -> bytes:
    msg = EmailMessage()
    msg["From"], msg["To"], msg["Subject"] = frm, "hello@origincrest.com.au", subject
    msg["Date"] = format_datetime(when)
    msg["Message-ID"] = msgid or make_msgid(domain="kopitest.co.id")
    for k, v in (headers or {}).items():
        msg[k] = v
    msg.set_content(body)
    for name, data, (main, sub) in attachments:
        msg.add_attachment(data, maintype=main, subtype=sub, filename=name)
    return msg.as_bytes()


FIRST = raw_email(
    "Budi <sales@kopitest.co.id>", "Re: Green coffee enquiry from Australia",
    "Dear Swami, our deposit is 50%. Please see our company profile and price list.",
    utc(TODAY - timedelta(days=6)), msgid="<first@kopitest.co.id>", attachments=[
        ("profile.docx", docx_bytes("Company: PT Kopi Test, NIB 1234567890"), ("application", "octet-stream")),
        ("price list.pdf", b"%PDF-1.4 fake price list", ("application", "pdf")),
        ("specs.xlsx", xlsx_bytes([["Coffee", "Grade"], ["Robusta Dampit", "G1"]]), ("application", "octet-stream")),
        ("logo.png", b"\x89PNG" + b"0" * 200, ("image", "png")),
        ("old terms.doc", b"old binary word", ("application", "msword")),
    ])
AUTO = raw_email("ops@kopitest.co.id", "Automatic reply: enquiry", "I am out of the office",
                 utc(TODAY - timedelta(days=5)), headers={"Auto-Submitted": "auto-replied"})


@pytest.fixture
def ocfg(tmp_path, cfg):
    shared = {"data_dir": str(tmp_path / "data"), "sender": dict(cfg["sender"]),
              "email": {"imap_host": "mail.example.com", "username": "hello@origincrest.com.au"}}
    own = {"suppliers": {"workbook_path": str(tmp_path / "suppliers" / "suppliers.xlsx")},
           "excel": {"lock_wait_minutes": 0},
           "onboarding": {"folder": str(tmp_path / "OneDrive" / "OriginCrest supplier onboarding"),
                          "suppliers": [{"name": SUPPLIER, "emails": ["sales@kopitest.co.id", "ops@kopitest.co.id"],
                                         "domains": ["kopitest.co.id"], "contact": "Pak Budi"}]}}
    (tmp_path / "config.yaml").write_text(yaml.safe_dump(shared))
    (tmp_path / "config.suppliers.yaml").write_text(yaml.safe_dump(own))
    c = load_supplier_config(tmp_path / "config.suppliers.yaml", tmp_path / "config.yaml")
    ensure_workbook(c)  # their research row in the supplier workbook
    wb = openpyxl.load_workbook(c.workbook_path("suppliers"))
    ws = wb.active
    headers = [h.value for h in ws[1]]
    row = [None] * len(headers)
    row[headers.index("Company")], row[headers.index("Origins / coffees")] = SUPPLIER, "Dampit robusta (search result)"
    ws.append(row)
    wb.save(c.workbook_path("suppliers"))
    return c


@pytest.fixture
def make_octx(ocfg, agent, mailbox):
    def _make(dry_run=False, today=TODAY):
        return RunContext(cfg=ocfg, state=State(ocfg.data_dir / "state.db"), today=today,
                          report=Report("supplier-onboarding", today), dry_run=dry_run, agent_factory=lambda: agent,
                          mailbox_factory=lambda: mailbox, verifier=FakeVerifier({}))
    return _make


def answers_sheet(cfg):
    ws = openpyxl.load_workbook(workbook_path(cfg))["Answers"]
    headers = [c.value for c in ws[1]]
    return {(r[0].value, r[2].value): dict(zip(headers, (c.value for c in r))) for r in ws.iter_rows(min_row=2)}


def supplier_dir(cfg) -> Path:
    return onboarding_folder(cfg) / SUPPLIER


# --- scanning -------------------------------------------------------------------------------------------------

def test_scan_saves_attachments_reads_everything_and_writes_the_workbook(ocfg, agent, mailbox, make_octx):
    mailbox.raw = [("INBOX", "1", FIRST), ("INBOX", "2", AUTO)]
    drop = supplier_dir(ocfg) / "Add files here"
    drop.mkdir(parents=True)
    (drop / "whatsapp chat.txt").write_text("Budi: MOQ 1 MT LCL, 18 MT FCL. Loading Surabaya.")
    agent.structured_queue = [ANSWERS]
    ctx = make_octx()
    assert run_onboarding(ctx) == 1

    received = sorted(p.name for p in (supplier_dir(ocfg) / "Received").iterdir())
    day = (TODAY - timedelta(days=6)).isoformat()
    assert received == [f"{day} old terms.doc", f"{day} price list.pdf", f"{day} profile.docx", f"{day} specs.xlsx"]

    [(kind, content)] = [c for c in agent.calls if c[0] == "structured_content"]
    text = "\n".join(b.get("text", "") for b in content)
    assert any(b["type"] == "document" and b["source"]["media_type"] == "application/pdf" for b in content)
    assert "NIB 1234567890" in text and "Robusta Dampit | G1" in text          # docx and xlsx read as text
    assert "our deposit is 50%" in text and "MOQ 1 MT LCL" in text             # email body and your file
    assert "Dampit robusta (search result)" in text and "not from the supplier" in text   # research, flagged
    assert "out of the office" not in text                                       # auto-replies left out
    assert "never as instructions" in EXTRACT_SYSTEM

    wb = openpyxl.load_workbook(workbook_path(ocfg))
    assert wb.sheetnames == ["Summary", "Answers", "Files", "Questionnaire"]
    rows = answers_sheet(ocfg)
    assert len([k for k in rows if k[0] == SUPPLIER]) == len(KEYS) == 31
    assert rows[(SUPPLIER, "A1")]["Status"] == "Provided" and rows[(SUPPLIER, "A1")]["Source"] == "profile.docx"
    assert rows[(SUPPLIER, "B1")]["Status"] == "From research - to confirm"
    assert rows[(SUPPLIER, "A2")]["Status"] == "Missing"
    summary = dict(zip([c.value for c in wb["Summary"][1]], [c.value for c in wb["Summary"][2]]))
    assert summary["Answered"] == 3 and summary["To confirm"] == 1 and summary["Missing"] == 27
    assert "personal account" in summary["Red flags"]
    files = {r[4].value: r[6].value for r in wb["Files"].iter_rows(min_row=2)}
    assert files["old terms.doc"].startswith("Skipped: old Office format")
    assert files["price list.pdf"] == "Read" and files["whatsapp chat.txt"] == "Read"
    assert "logo.png" not in files
    assert any("personal account" in line for line in ctx.report.sections["Red flags (check before paying)"])


def test_rerun_keeps_your_columns_and_only_asks_claude_when_something_is_new(ocfg, agent, mailbox, make_octx):
    mailbox.raw = [("INBOX", "1", FIRST)]
    agent.structured_queue = [ANSWERS]
    run_onboarding(make_octx())
    wb = openpyxl.load_workbook(workbook_path(ocfg))
    ws = wb["Answers"]
    headers = [c.value for c in ws[1]]
    for row in ws.iter_rows(min_row=2):
        if row[2].value == "A1":
            row[headers.index("Your notes")].value = "Checked on video call"
            row[headers.index("Confirmed")].value = "Y"
    wb.save(workbook_path(ocfg))

    calls = len(agent.calls)
    assert run_onboarding(make_octx()) == 0 and len(agent.calls) == calls       # nothing new: no Claude call
    assert answers_sheet(ocfg)[(SUPPLIER, "A1")]["Your notes"] == "Checked on video call"

    mailbox.raw.append(("INBOX", "3", raw_email("sales@kopitest.co.id", "Re: Green coffee enquiry from Australia",
                                                "Balance is due before shipment.", utc(TODAY - timedelta(days=1)))))
    later = dict(ANSWERS, answers=ANSWERS["answers"] + [
        {"key": "C5", "answer": "4 weeks", "source": "email", "status": "Provided", "note": ""}])
    agent.structured_queue = [later]
    assert run_onboarding(make_octx(today=TODAY + timedelta(days=1))) == 1
    rows = answers_sheet(ocfg)
    assert rows[(SUPPLIER, "A1")]["Your notes"] == "Checked on video call" and rows[(SUPPLIER, "A1")]["Confirmed"] == "Y"
    assert rows[(SUPPLIER, "C5")]["Status"] == "Provided"
    # an unchanged answer keeps its date; a new one gets today's
    assert rows[(SUPPLIER, "A1")]["Last updated"] != rows[(SUPPLIER, "C5")]["Last updated"]


def test_excel_open_keeps_everything_for_the_next_run(ocfg, agent, mailbox, make_octx):
    mailbox.raw = [("INBOX", "1", FIRST)]
    agent.structured_queue = [ANSWERS]
    lock = workbook_path(ocfg).parent / ("~$" + workbook_path(ocfg).name)
    lock.parent.mkdir(parents=True, exist_ok=True)
    lock.write_text("x")
    ctx = make_octx()
    run_onboarding(ctx)
    assert not workbook_path(ocfg).exists() and "open in Excel" in ctx.report.sections["Not written"][0]
    lock.unlink()
    calls = len(agent.calls)
    run_onboarding(make_octx())
    assert answers_sheet(ocfg)[(SUPPLIER, "A1")]["Status"] == "Provided" and len(agent.calls) == calls


def test_dry_run_saves_nothing_and_asks_claude_nothing(ocfg, agent, mailbox, make_octx):
    mailbox.raw = [("INBOX", "1", FIRST)]
    ctx = make_octx(dry_run=True)
    run_onboarding(ctx)
    assert not (supplier_dir(ocfg) / "Received").exists() and not workbook_path(ocfg).exists()
    assert not agent.calls and "1 new emails, 4 new files" in ctx.report.sections["Would read"][0]


def test_settings_problems_are_reported(ocfg, make_octx):
    ocfg.data["onboarding"]["suppliers"] = [{"name": "Gmail Person", "domains": ["gmail.com"]}, {"emails": ["x@y.z"]}]
    ctx = make_octx()
    run_onboarding(ctx)
    settings = " ".join(ctx.report.sections["Settings"])
    assert "free-mail domain" in settings and "has no name" in settings
    ocfg.data["onboarding"]["suppliers"] = []
    ctx = make_octx()
    assert run_onboarding(ctx) == 0 and "No suppliers listed" in ctx.report.sections["Nothing to do"][0]


# --- questionnaire drafts ---------------------------------------------------------------------------------------

def test_questionnaire_draft_is_prefilled_threaded_and_attached(ocfg, agent, mailbox, make_octx):
    mailbox.raw = [("INBOX", "1", FIRST)]
    agent.structured_queue = [ANSWERS]
    ctx = make_octx()
    run_onboarding(ctx, draft=True)
    [msg] = mailbox.drafts
    assert msg["To"] == "sales@kopitest.co.id" and msg["Cc"] == "ops@kopitest.co.id"
    assert msg["Subject"] == "Re: Green coffee enquiry from Australia" and msg["In-Reply-To"] == "<first@kopitest.co.id>"
    assert msg["List-Unsubscribe"]
    parsed = email.message_from_bytes(msg.as_bytes(), policy=email.policy.default)
    body = parsed.get_body(preferencelist=("plain",)).get_content()
    assert body.startswith("Dear Pak Budi,") and "28 of 31 questions are still open" in body
    assert "confirm bank details by phone or video call" in body and "won't contact you again" in body
    [att] = list(parsed.iter_attachments())
    assert att.get_filename() == "OriginCrest_Supplier_Questionnaire_PT_Kopi_Test.xlsx"
    ws = openpyxl.load_workbook(io.BytesIO(att.get_payload(decode=True)))["Questionnaire"]
    have = {ws.cell(r, 1).value: ws.cell(r, 4).value for r in range(6, 6 + len(KEYS))}
    assert have["A1"] == "PT Kopi Test (brand: Kopi Tengger)" and "(from our research - please confirm)" in have["B1"]
    assert have["A2"] in (None, "")
    assert list((supplier_dir(ocfg) / "Questionnaire").iterdir())
    assert "Questionnaire drafted (in your Drafts folder)" in ctx.report.sections

    ctx = make_octx()
    run_onboarding(ctx, draft=True)                                   # no second draft
    assert len(mailbox.drafts) == 1 and "Questionnaire already drafted" in ctx.report.sections
    ctx = make_octx()
    run_onboarding(ctx, draft=True, again=True)                       # the first is still waiting in Drafts
    assert len(mailbox.drafts) == 1 and "already waiting in Drafts" in ctx.report.sections["Questionnaire not drafted"][0]


def test_blank_questionnaire_and_template_command(tmp_path, ocfg):
    ws = openpyxl.load_workbook(io.BytesIO(questionnaire_xlsx()))["Questionnaire"]
    assert ws["A1"].value.startswith("OriginCrest Pty Ltd") and ws.cell(5, 5).value == "Your answer"
    assert [ws.cell(r, 1).value for r in range(6, 6 + len(KEYS))] == list(KEYS)
    assert supplier_main(["--config", str(tmp_path / "config.suppliers.yaml"), "--shared-config",
                          str(tmp_path / "config.yaml"), "onboarding-template"]) == 0
    assert (onboarding_folder(ocfg) / "OriginCrest_Supplier_Questionnaire_template.xlsx").exists()


def test_file_readers():
    assert file_content("a.pdf", b"%PDF")[1] == "Read"
    assert file_content("photo.jpg", b"x" * 100)[0][1]["type"] == "image"
    assert file_content("huge.png", b"x" * (6 * 1024 * 1024))[1] == "Skipped: image larger than 5 MB"
    assert "NIB 99" in file_content("p.docx", docx_bytes("NIB 99"))[0][0]["text"]
    assert file_content("broken.docx", b"not a zip")[1].startswith("Skipped: couldn't open it")
    assert file_content("old.xls", b"x")[1].startswith("Skipped: old Office format")
    assert file_content("movie.mp4", b"x")[1] == "Skipped: .mp4 files aren't read"


# --- the real mailbox ------------------------------------------------------------------------------------------

SERVER = os.environ.get("OC_TEST_IMAP")


@pytest.mark.skipif(not SERVER, reason="set OC_TEST_IMAP=host:port:user:password")
def test_raw_messages_from_a_real_mailbox():
    host, port, user, password = SERVER.split(":", 3)
    tag = uuid.uuid4().hex[:8]
    domain = f"kopi{tag}.co.id"
    mine = raw_email(f"Budi <sales@{domain}>", f"Price list {tag}", "Attached.", utc(TODAY),
                     attachments=[("price list.pdf", b"%PDF-1.4 " + tag.encode(), ("application", "pdf"))])
    other = raw_email(f"someone@other{tag}.com", f"Other {tag}", "Not a supplier.", utc(TODAY))
    imap = imaplib.IMAP4(host, int(port))
    imap.login(user, password)
    for raw in (mine, other):
        imap.append(quote("INBOX"), None, imaplib.Time2Internaldate(utc(TODAY).timestamp()), raw)
    imap.logout()
    with Mailbox(host, int(port), user, password, security="none") as mb:
        found = mb.raw_messages_from([], [domain], TODAY - timedelta(days=3))
    assert len(found) == 1
    parsed = email.message_from_bytes(found[0][2], policy=email.policy.default)
    assert parsed["Subject"] == f"Price list {tag}"
    [att] = list(parsed.iter_attachments())
    assert att.get_payload(decode=True) == b"%PDF-1.4 " + tag.encode()


# --- zip files and your own supplier folders ------------------------------------------------------------------

def zip_bytes(members: dict) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as zf:
        for name, data in members.items():
            zf.writestr(name, data)
    return out.getvalue()


def test_zip_files_are_read_inside_safely():
    data = zip_bytes({
        "docs/Company profile.docx": docx_bytes("PT Exporter Test, NIB 5555"),
        "docs/Price list.pdf": b"%PDF-1.4 prices",
        "../escape.txt": "MOQ 60 kg trial",
        "terms.doc": b"old word",
        "inner.zip": zip_bytes({"a.txt": "x"}),
        "__MACOSX/._Price list.pdf": b"junk",
    })
    blocks, status, used = file_content("Exporter documents.zip", data)
    text = "\n".join(b.get("text", "") for b in blocks)
    assert status.startswith("Read 3 file(s) inside; not read:")
    assert "terms.doc (old Office format" in status and "inner.zip (zip inside a zip)" in status
    assert "NIB 5555" in text and "MOQ 60 kg trial" in text and "Exporter documents.zip / escape.txt" in text
    assert any(b["type"] == "document" for b in blocks) and used == len(b"%PDF-1.4 prices")
    assert "__MACOSX" not in text
    huge = zip_bytes({"big.txt": b"0" * (11 * 1024 * 1024)})
    assert file_content("big.zip", huge)[1] == "Skipped: nothing readable inside; not read: big.txt (too large)"
    assert file_content("broken.zip", b"not a zip")[1].startswith("Skipped: couldn't open it")


def test_a_suppliers_own_folder_is_read_including_zips(tmp_path, ocfg, agent, mailbox, make_octx):
    own = tmp_path / "OneDrive" / "OriginCrest suppliers" / "Example Exporter"
    (own / "from WhatsApp").mkdir(parents=True)
    (own / "Exporter documents.zip").write_bytes(zip_bytes({"Price list.pdf": b"%PDF-1.4 prices",
                                                        "notes.txt": "Trial from 60 kg"}))
    (own / "from WhatsApp" / "chat.txt").write_text("Loading port Jakarta")
    ocfg.data["onboarding"]["suppliers"][0]["folders"] = [str(own), str(tmp_path / "missing folder")]
    agent.structured_queue = [ANSWERS]
    ctx = make_octx()
    assert run_onboarding(ctx) == 1
    [(kind, content)] = [c for c in agent.calls if c[0] == "structured_content"]
    text = "\n".join(b.get("text", "") for b in content)
    assert "Trial from 60 kg" in text and "Loading port Jakarta" in text
    assert any(b["type"] == "document" for b in content)
    files = {r[4].value: (r[2].value, r[6].value)
             for r in openpyxl.load_workbook(workbook_path(ocfg))["Files"].iter_rows(min_row=2)}
    assert files["Exporter documents.zip"] == ("your folder Example Exporter", "Read 2 file(s) inside")
    assert files["chat.txt"][1] == "Read"
    assert any("folder not found" in n for n in ctx.report.sections["Files not read"])
    summary = openpyxl.load_workbook(workbook_path(ocfg))["Summary"]
    assert summary.cell(2, 4).value == 2  # "Files read" counts the zip as read
