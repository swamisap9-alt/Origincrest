"""Runs against a real IMAP server when OC_TEST_IMAP=host:port:user:password is set (plain IMAP, e.g. Dovecot)."""

import imaplib
import os
import time
import uuid
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage
from email.utils import format_datetime

import pytest

from origincrest_agent.followup import run_followup
from origincrest_agent.mailer import Mailbox, build_message, quote

from .conftest import TODAY, utc

SERVER = os.environ.get("OC_TEST_IMAP")
pytestmark = pytest.mark.skipif(not SERVER, reason="set OC_TEST_IMAP=host:port:user:password")


def _parts():
    host, port, user, password = SERVER.split(":", 3)
    return host, int(port), user, password


def _put(folder, *, from_addr, to_addr, subject, body, when, headers=None, message_id=None):
    host, port, user, password = _parts()
    msg = EmailMessage()
    msg["From"], msg["To"], msg["Subject"] = from_addr, to_addr, subject
    msg["Date"] = format_datetime(when)
    msg["Message-ID"] = message_id or f"<{uuid.uuid4()}@test>"
    for k, v in (headers or {}).items():
        msg[k] = v
    msg.set_content(body)
    imap = imaplib.IMAP4(host, port)
    imap.login(user, password)
    imap.append(quote(folder), None, imaplib.Time2Internaldate(when.timestamp()), msg.as_bytes())
    imap.logout()
    return msg["Message-ID"]


def _mailbox():
    host, port, user, password = _parts()
    return Mailbox(host, port, user, password, security="none")


def test_folders_drafts_and_searches():
    tag = uuid.uuid4().hex[:8]
    contact = f"quotes-{tag}@seaway{tag}.com.au"
    sent_when = utc(TODAY - timedelta(days=6))
    rfq_id = _put("Sent Mail", from_addr="swami@example.com", to_addr=contact, subject="RFQ 20ft",
                  body="Please quote", when=sent_when)
    _put("INBOX", from_addr=f"jo@seaway{tag}.com.au", to_addr="swami@example.com", subject="Re: RFQ 20ft",
         body="Quote attached\n\nOn Mon, Swami wrote:\n> Please quote", when=sent_when + timedelta(days=1))
    _put("INBOX", from_addr=contact, to_addr="swami@example.com", subject="Automatic reply: RFQ 20ft",
         body="I'm away", when=sent_when + timedelta(hours=2), headers={"Auto-Submitted": "auto-replied"})
    _put("INBOX", from_addr="mailer-daemon@googlemail.com", to_addr="swami@example.com",
         subject="Delivery Status Notification (Failure)", body=f"Address not found: {contact}",
         when=sent_when + timedelta(minutes=5))

    with _mailbox() as mb:
        assert mb.drafts == "Drafts" and mb.sent == "Sent Mail" and mb.reply_folders == ["INBOX"]
        sent = mb.messages_to(contact, TODAY - timedelta(days=30))
        assert [m.message_id for m in sent] == [rfq_id]
        replies = mb.messages_from(contact, TODAY - timedelta(days=30), domain=f"seaway{tag}.com.au")
        assert {m.subject for m in replies} == {"Re: RFQ 20ft", "Automatic reply: RFQ 20ft"}
        genuine = [m for m in replies if not m.auto_reply]
        assert len(genuine) == 1 and genuine[0].text == "Quote attached"
        assert len(mb.bounces_for(contact, TODAY - timedelta(days=30))) == 1

        msg = build_message(from_name="Swami", from_addr="swami@example.com", to_addr=contact, to_name="",
                            subject="Re: RFQ 20ft", body="Following up", in_reply_to=rfq_id)
        assert not mb.draft_waiting(msg["Message-ID"], contact)
        mb.append_draft(msg)
        assert mb.draft_waiting(msg["Message-ID"])
        assert mb.draft_waiting(to_addr=contact)


def test_followup_end_to_end(cfg, agent, make_ctx, monkeypatch):
    """Workflow 3 against the real server: thread lookup, draft saved, sheet updated."""
    tag = uuid.uuid4().hex[:8]
    contact = f"rfq-{tag}@fastfreight{tag}.com.au"
    from origincrest_agent.workbook import WorkbookStore

    store = WorkbookStore("rfq", cfg.workbook_path("rfq"), cfg).load()
    rows = list(store.view("RFQ Tracker").rows())
    rows[0].set("email", contact)
    rows[1].set("status", "Quoted")
    store.save(cfg.data_dir / "backups")
    rfq_id = _put("Sent Mail", from_addr="swami@example.com", to_addr=contact, subject=f"RFQ {tag}",
                  body="Could you quote LCL Surabaya-Adelaide?", when=utc(TODAY - timedelta(days=8)))
    time.sleep(0.2)
    ctx = make_ctx("followup")
    ctx.mailbox_factory = _mailbox
    assert run_followup(ctx) == 1
    with _mailbox() as mb:
        assert mb.draft_waiting(to_addr=contact)
        uids = mb._search(mb.drafts, ["TO", quote(contact)])
        draft = mb._fetch(mb.drafts, uids)[0]
    assert draft.subject == f"Re: RFQ {tag}"
    assert rfq_id in draft.references


def test_signature_learned_from_sent_folder():
    tag = uuid.uuid4().hex[:6]
    signature = f"Swami K {tag}\nOriginCrest | 0400 000 000"
    for i in range(2):
        _put("Sent Mail", from_addr="Swami K <swami@example.com>", to_addr=f"lead{i}-{tag}@example.com",
             subject=f"Note {i}", body=f"Hi,\n\nNote {i}.\n\nKind regards,\n{signature}",
             when=datetime.now(timezone.utc) + timedelta(seconds=i))
    with _mailbox() as mb:
        learned = mb.learn_sender()
    assert learned.email == "swami@example.com"
    assert learned.name == "Swami K"
    assert learned.signature == signature


def test_agent_drafts_listed_and_duplicate_drafts_folders_found():
    host, port, user, password = _parts()
    imap = imaplib.IMAP4(host, port)
    imap.login(user, password)
    imap.create(quote("Old/Drafts"))  # a second Drafts-like folder, as some hosts have
    imap.logout()
    tag = uuid.uuid4().hex[:6]
    msg = build_message(from_name="", from_addr="swami@example.com", to_addr=f"rfq-{tag}@example.com", to_name="",
                        subject=f"RFQ {tag}", body="Please quote", agent_tag="rfq; test")
    with _mailbox() as mb:
        mb.append_draft(msg)
        assert mb.folder_candidates("drafts") == ["Drafts", "Old/Drafts"]
        assert f"RFQ {tag}" in [d.subject for d in mb.agent_drafts()]
        assert mb.agent_drafts("Old/Drafts") == []
        assert mb.message_count("Drafts") >= 1


def test_signature_comes_from_newest_mail_across_sent_folders():
    """Servers often have both "Sent" and Outlook's "Sent Items"; the newest emails win wherever they are."""
    host, port, user, password = _parts()
    imap = imaplib.IMAP4(host, port)
    imap.login(user, password)
    imap.create(quote("Sent Items"))
    imap.logout()
    tag = uuid.uuid4().hex[:6]
    old_sig = f"Swami K\nOld number 0400 000 000 {tag}"
    new_sig = f"Swaminathan Kailasam\nDirector, OriginCrest\nM 0400 999 888 {tag}"
    now = datetime.now(timezone.utc)
    for i in range(2):
        _put("Sent Mail", from_addr="Swami K <swami@example.com>", to_addr=f"old{i}-{tag}@example.com",
             subject="Old", body=f"Hi,\n\nOld {i}.\n\nRegards,\n{old_sig}", when=now - timedelta(days=30, seconds=i))
        _put("Sent Items", from_addr="Swami K <swami@example.com>", to_addr=f"new{i}-{tag}@example.com",
             subject="New", body=f"Hi,\n\nNew {i}.\n\nKind regards,\n{new_sig}", when=now + timedelta(minutes=5, seconds=i))
    with _mailbox() as mb:
        assert mb.sent_folders == ["Sent Mail", "Sent Items"]
        learned = mb.learn_sender()
        assert learned.signature == new_sig
        assert "Sent Items" in learned.folders
        assert [m.subject for m in mb.messages_to(f"new0-{tag}@example.com", TODAY - timedelta(days=5))] == ["New"]
