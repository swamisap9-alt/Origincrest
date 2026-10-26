from datetime import datetime, timezone

from origincrest_agent.mailer import MailInfo, build_message, parse_message, strip_quoted


def info(subject="Re: hello", text="", from_addr="jo@x.com.au"):
    return MailInfo("INBOX", "1", datetime.now(timezone.utc), from_addr, [], subject, "<a@b>", "", False, text)


def test_unsubscribe_detection():
    assert info(text="Please remove me from your list").is_unsubscribe
    assert info(subject="Unsubscribe").is_unsubscribe
    assert info(text="Don't email us again").is_unsubscribe
    assert not info(text="Please remove the pallets from bay 3 and send a quote").is_unsubscribe
    assert not info(text="We can opt out of insurance on this shipment").is_unsubscribe


def test_bounce_detection():
    assert info(subject="Delivery Status Notification (Failure)", from_addr="mailer-daemon@googlemail.com").is_bounce
    assert not info(subject="Re: RFQ").is_bounce


def test_quoted_history_removed():
    text = "Sounds good, send samples.\n\nOn Mon, 20 Oct 2026, Swami wrote:\n> unsubscribe footer here"
    assert strip_quoted(text) == "Sounds good, send samples."
    outlook = "Thanks!\n________________________________\nFrom: Swami\nSent: Monday\n...unsubscribe..."
    assert "unsubscribe" not in strip_quoted(outlook)


def test_auto_reply_headers():
    raw = (b"From: a@x.com\r\nTo: b@y.com\r\nSubject: Out of Office: back Monday\r\n"
           b"Date: Mon, 20 Oct 2026 09:00:00 +1030\r\nMessage-ID: <1@x>\r\n\r\nAway")
    msg = parse_message(raw, "INBOX", "1")
    assert msg.auto_reply and msg.when.utcoffset().total_seconds() == 37800


def test_build_message_threading():
    msg = build_message(from_name="Swami", from_addr="hello@origincrest.com.au", to_addr="jo@x.com.au",
                        to_name="Jo", subject="Re: RFQ", body="Hi", in_reply_to="<1@x>", references="<0@x>",
                        unsubscribe_addr="hello@origincrest.com.au")
    assert msg["In-Reply-To"] == "<1@x>"
    assert msg["References"] == "<0@x> <1@x>"
    assert msg["Message-ID"].endswith("@origincrest.com.au>")
    assert msg["List-Unsubscribe"].startswith("<mailto:")


class _FakeImap:
    def __init__(self, lines):
        self.lines = lines

    def list(self):
        return "OK", self.lines


def _discover(lines, **kw):
    from origincrest_agent.mailer import Mailbox

    mb = Mailbox("mail.example.com.au", 993, "swami@origincrest.com.au", "pw", **kw)
    mb.imap = _FakeImap(lines)
    mb._discover_folders()
    return mb


def test_folders_found_by_name_when_server_has_no_special_use_flags():
    # Typical cPanel/Courier layout: "." separator, everything under INBOX, no \\Drafts / \\Sent flags.
    mb = _discover([b'(\\HasChildren) "." INBOX', b'(\\HasNoChildren) "." INBOX.Drafts',
                    b'(\\HasNoChildren) "." "INBOX.Sent Items"', b'(\\HasNoChildren) "." INBOX.Trash'])
    assert (mb.drafts, mb.sent, mb.reply_folders) == ("INBOX.Drafts", "INBOX.Sent Items", ["INBOX"])


def test_folder_name_sent_as_literal():
    mb = _discover([b'(\\HasNoChildren \\Drafts) "/" Drafts', (b'(\\HasNoChildren \\Sent) "/" {13}', b'Sent Messages')])
    assert (mb.drafts, mb.sent) == ("Drafts", "Sent Messages")


def test_own_addresses_and_security_option():
    import pytest
    from origincrest_agent.mailer import Mailbox

    mb = Mailbox("h", 993, "swami", "pw", own_addresses=["Hello@OriginCrest.com.au"])
    assert mb.own_addresses == {"hello@origincrest.com.au"}  # a non-email login name isn't added
    with pytest.raises(ValueError):
        Mailbox("h", 993, "u", "pw", security="tls1")


SIG = "Swami Kailasam\nDirector | OriginCrest\nM 0400 111 222 | www.origincrest.com.au"


def test_signature_learned_after_signoff():
    from origincrest_agent.mailer import learn_signature

    texts = [f"Hi Jo,\n\nQuote please.\n\nKind regards,\n{SIG}",
             f"Hello,\n\nThanks for the samples!\n\nCheers,\n{SIG}",
             "Sent from my phone - short reply"]
    assert learn_signature(texts) == SIG


def test_templated_emails_dont_swallow_the_body():
    from origincrest_agent.mailer import learn_signature

    body = "Could you quote 1x20ft FCL Belawan to Adelaide?\nGreen coffee in jute bags.\n\n"
    texts = [f"Hi {name},\n\n{body}Kind regards,\n{SIG}" for name in ("Jo", "Sam", "team")]
    assert learn_signature(texts) == SIG


def test_signature_without_signoff_and_agent_footer_removed():
    from origincrest_agent.mailer import learn_signature

    texts = [f"Hi Jo, see attached.\n\n{SIG}",
             f"Morning, confirming Tuesday.\n\n{SIG}\n\n--\nIf you'd rather not hear from OriginCrest again, reply "
             'with "unsubscribe"']
    assert learn_signature(texts) == SIG
    assert learn_signature(["Hi, quick one.\nThanks", "Another note"]) == ""


def test_html_only_sent_email_text():
    raw = (b"From: Swami <s@o.com.au>\r\nTo: a@b.com\r\nSubject: x\r\nDate: Mon, 20 Oct 2026 09:00:00 +1030\r\n"
           b"Content-Type: text/html; charset=utf-8\r\n\r\n"
           b"<div>Hi Jo,</div><div>Quote please.</div><p>Kind regards,<br>Swami&nbsp;Kailasam<br>OriginCrest</p>")
    msg = parse_message(raw, "Sent", "1")
    assert msg.from_name == "Swami"
    assert msg.text.splitlines()[-3:] == ["Kind regards,", "Swami Kailasam", "OriginCrest"]
