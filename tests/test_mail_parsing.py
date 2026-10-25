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
