"""IMAP mailbox access: save drafts, find replies, sent mail and bounces.

The agent never sends email. It only APPENDs messages to the Drafts folder,
where you review, edit and send (or delete) them yourself.
"""

from __future__ import annotations

import email
import email.policy
import imaplib
import re
import ssl
import time
from dataclasses import dataclass
from datetime import date, datetime
from email.message import EmailMessage
from email.utils import formataddr, formatdate, getaddresses, make_msgid, parsedate_to_datetime

from .utils import log

_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
_LIST_RE = re.compile(r'^\((?P<flags>[^)]*)\)\s+(?P<delim>"(?:[^"\\]|\\.)*"|NIL)\s+(?P<name>.+)$')
_GUESSES = {
    "drafts": ["[Gmail]/Drafts", "Drafts", "INBOX.Drafts", "Draft", "INBOX/Drafts"],
    "sent": ["[Gmail]/Sent Mail", "Sent", "Sent Items", "Sent Messages", "INBOX.Sent", "INBOX/Sent"],
}
_AUTO_SUBJECT = re.compile(
    r"^(automatic reply|auto[- ]?reply|auto:|out of (the )?office|ooo\b|autoreply|away from (the )?office"
    r"|i am (currently )?out of|on leave)", re.I)
_BOUNCE_SUBJECT = re.compile(
    r"(undeliverable|delivery status notification|mail delivery (failed|subsystem)|returned mail|"
    r"delivery failure|failure notice|could not be delivered)", re.I)
_UNSUB = re.compile(
    r"\b(unsubscribe|remove (me|us)\b|take (me|us) off|opt[- ]?out of (these|your|future)|"
    r"stop (emailing|e-mailing|contacting|sending (me|us))|no longer wish to (receive|hear)|"
    r"(do not|don't|please don't) (contact|email|e-mail) (me|us))", re.I)


def imap_date(d: date) -> str:
    return f"{d.day:02d}-{_MONTHS[d.month - 1]}-{d.year}"


def quote(text: str) -> str:
    return '"' + text.replace("\\", "\\\\").replace('"', '\\"') + '"'


@dataclass
class MailInfo:
    folder: str
    uid: str
    when: datetime | None
    from_addr: str
    to_addrs: list[str]
    subject: str
    message_id: str
    references: str
    auto_reply: bool
    text: str  # first part of the body, quoted history removed

    @property
    def is_unsubscribe(self) -> bool:
        return bool(_UNSUB.search(self.subject) or _UNSUB.search(self.text[:600]))

    @property
    def is_bounce(self) -> bool:
        return bool(re.search(r"mailer-daemon|postmaster", self.from_addr, re.I) or _BOUNCE_SUBJECT.search(self.subject))


def strip_quoted(text: str) -> str:
    lines: list[str] = []
    for line in text.splitlines():
        stripped = line.strip()
        if stripped.startswith(">"):
            continue
        if re.match(r"^(On .+ wrote:|-{2,}\s*Original Message\s*-{2,}|From: .+|Sent from my )", stripped, re.I):
            break
        lines.append(line)
    return "\n".join(lines).strip()


def _body_text(msg) -> str:
    try:
        part = msg.get_body(preferencelist=("plain", "html"))
        if part is None:
            return ""
        content = part.get_content()
        if part.get_content_type() == "text/html":
            content = re.sub(r"<(script|style)[^>]*>.*?</\1>|<[^>]+>", " ", content, flags=re.S | re.I)
        return content
    except (KeyError, LookupError, UnicodeDecodeError, AttributeError):
        return ""


def parse_message(raw: bytes, folder: str, uid: str, include_body: bool = True) -> MailInfo:
    msg = email.message_from_bytes(raw, policy=email.policy.default)
    try:
        when = parsedate_to_datetime(str(msg["Date"])) if msg["Date"] else None
    except (TypeError, ValueError):
        when = None
    from_addr = (getaddresses([str(msg.get("From", ""))]) or [("", "")])[0][1].lower()
    to_addrs = [a.lower() for _, a in getaddresses([str(msg.get("To", "")), str(msg.get("Cc", ""))]) if a]
    subject = str(msg.get("Subject", "")).strip()
    auto_submitted = str(msg.get("Auto-Submitted", "no")).lower()
    precedence = str(msg.get("Precedence", "")).lower()
    auto = (
        (auto_submitted not in ("", "no"))
        or bool(msg.get("X-Autoreply") or msg.get("X-Autorespond") or msg.get("X-Auto-Response-Suppress"))
        or precedence in ("auto_reply", "bulk", "junk")
        or bool(_AUTO_SUBJECT.search(subject))
    )
    text = strip_quoted(_body_text(msg))[:3000] if include_body else ""
    return MailInfo(folder, uid, when, from_addr, to_addrs, subject, str(msg.get("Message-ID", "")).strip(),
                    str(msg.get("References", "")), auto, text)


def build_message(*, from_name: str, from_addr: str, to_addr: str, to_name: str, subject: str, body: str,
                  in_reply_to: str = "", references: str = "", unsubscribe_addr: str = "",
                  agent_tag: str = "") -> EmailMessage:
    msg = EmailMessage()
    msg["From"] = formataddr((from_name, from_addr))
    msg["To"] = formataddr((to_name, to_addr)) if to_name else to_addr
    msg["Subject"] = subject
    msg["Date"] = formatdate(localtime=True)
    domain = from_addr.rsplit("@", 1)[-1] if "@" in from_addr else None
    msg["Message-ID"] = make_msgid(domain=domain)
    if in_reply_to:
        msg["In-Reply-To"] = in_reply_to
        msg["References"] = (references + " " + in_reply_to).strip()
    if unsubscribe_addr:
        msg["List-Unsubscribe"] = f"<mailto:{unsubscribe_addr}?subject=unsubscribe>"
    if agent_tag:
        msg["X-OriginCrest-Agent"] = agent_tag
    msg.set_content(body)
    return msg


class Mailbox:
    def __init__(self, host: str, port: int, username: str, password: str, *, drafts_folder: str = "",
                 sent_folder: str = "", reply_folders: list[str] | None = None, max_fetch_bytes: int = 5_000_000,
                 use_ssl: bool = True):
        self.host, self.port, self.username, self.password = host, port, username, password
        self.use_ssl = use_ssl
        self.drafts = drafts_folder
        self.sent = sent_folder
        self.reply_folders = list(reply_folders or [])
        self.max_fetch_bytes = max_fetch_bytes
        self.imap: imaplib.IMAP4 | None = None
        self.folders: dict[str, set[str]] = {}
        self._bounce_cache: dict[str, list[MailInfo]] = {}

    @classmethod
    def from_config(cls, cfg) -> "Mailbox":
        e = cfg["email"]
        password = cfg.secret(e["password_env"])
        if not e["username"] or not password:
            raise RuntimeError(
                f"Email not configured: set email.username in config.yaml and {e['password_env']} in .env")
        return cls(e["imap_host"], int(e["imap_port"]), e["username"], password,
                   drafts_folder=e["drafts_folder"], sent_folder=e["sent_folder"],
                   reply_folders=e["reply_folders"], max_fetch_bytes=int(e["max_fetch_bytes"]),
                   use_ssl=e.get("use_ssl", True))

    # connection --------------------------------------------------------
    def __enter__(self) -> "Mailbox":
        self.connect()
        return self

    def __exit__(self, *exc) -> None:
        self.close()

    def connect(self) -> None:
        if self.use_ssl:
            self.imap = imaplib.IMAP4_SSL(self.host, self.port, ssl_context=ssl.create_default_context())
        else:
            self.imap = imaplib.IMAP4(self.host, self.port)
        self.imap.login(self.username, self.password)
        self._discover_folders()

    def close(self) -> None:
        if self.imap is not None:
            try:
                self.imap.logout()
            except (imaplib.IMAP4.error, OSError):
                pass
            self.imap = None

    def _discover_folders(self) -> None:
        typ, data = self.imap.list()
        self.folders = {}
        for line in data or []:
            if not line:
                continue
            m = _LIST_RE.match(line.decode("utf-8", "replace") if isinstance(line, bytes) else str(line))
            if not m:
                continue
            name = m.group("name").strip()
            if name.startswith('"') and name.endswith('"'):
                name = name[1:-1].replace('\\"', '"').replace("\\\\", "\\")
            self.folders[name] = {f.lower() for f in m.group("flags").split()}

        def special(flag: str, kind: str) -> str:
            for name, flags in self.folders.items():
                if flag in flags:
                    return name
            for guess in _GUESSES.get(kind, []):
                if guess in self.folders:
                    return guess
            return ""

        self.drafts = self.drafts or special("\\drafts", "drafts")
        self.sent = self.sent or special("\\sent", "sent")
        if not self.reply_folders:
            all_mail = special("\\all", "all")
            self.reply_folders = [all_mail] if all_mail else ["INBOX"]
        if not self.drafts:
            raise RuntimeError("Couldn't find the Drafts folder; set email.drafts_folder in config.yaml")
        log.debug("IMAP folders: drafts=%s sent=%s replies=%s", self.drafts, self.sent, self.reply_folders)

    # drafts ------------------------------------------------------------
    def append_draft(self, msg: EmailMessage) -> None:
        typ, resp = self.imap.append(quote(self.drafts), r"(\Draft \Seen)",
                                     imaplib.Time2Internaldate(time.time()), msg.as_bytes())
        if typ != "OK":
            raise RuntimeError(f"IMAP APPEND to {self.drafts} failed: {resp}")

    def draft_waiting(self, message_id: str = "", to_addr: str = "") -> bool:
        """True if a draft with this Message-ID, or any draft to this address, is still in Drafts."""
        if message_id and self._search(self.drafts, ["HEADER", "Message-ID", quote(message_id)]):
            return True
        if to_addr and self._search(self.drafts, ["TO", quote(to_addr)]):
            return True
        return False

    # searching ---------------------------------------------------------
    def _select(self, folder: str) -> bool:
        typ, _ = self.imap.select(quote(folder), readonly=True)
        return typ == "OK"

    def _search(self, folder: str, criteria: list[str]) -> list[str]:
        if not folder or not self._select(folder):
            return []
        try:
            typ, data = self.imap.uid("SEARCH", *criteria)
        except imaplib.IMAP4.error as exc:
            log.debug("IMAP search %s in %s failed: %s", criteria, folder, exc)
            return []
        if typ != "OK" or not data or not data[0]:
            return []
        return data[0].decode().split()

    def _fetch(self, folder: str, uids: list[str], include_body: bool = True) -> list[MailInfo]:
        out: list[MailInfo] = []
        if not uids:
            return out
        typ, sizes = self.imap.uid("FETCH", ",".join(uids), "(RFC822.SIZE)")
        size_by_uid: dict[str, int] = {}
        for item in sizes or []:
            raw = item[0] if isinstance(item, tuple) else item
            if isinstance(raw, bytes):
                m_uid, m_size = re.search(rb"UID (\d+)", raw), re.search(rb"RFC822\.SIZE (\d+)", raw)
                if m_uid and m_size:
                    size_by_uid[m_uid.group(1).decode()] = int(m_size.group(1))
        for uid in uids:
            whole = include_body and size_by_uid.get(uid, 0) <= self.max_fetch_bytes
            part = "BODY.PEEK[]" if whole else "BODY.PEEK[HEADER]"
            typ, data = self.imap.uid("FETCH", uid, f"({part})")
            raw = next((d[1] for d in data or [] if isinstance(d, tuple) and len(d) > 1), None)
            if raw:
                out.append(parse_message(raw, folder, uid, include_body=whole))
        return out

    def messages_from(self, address: str, since: date, domain: str = "") -> list[MailInfo]:
        """Messages received from an address (or anyone @domain) since a date, newest first."""
        found: dict[str, MailInfo] = {}
        for folder in self.reply_folders:
            uids = set(self._search(folder, ["SINCE", imap_date(since), "FROM", quote(address)])) if address else set()
            if domain:
                uids |= set(self._search(folder, ["SINCE", imap_date(since), "FROM", quote(domain)]))
            for info in self._fetch(folder, sorted(uids, key=int)[-20:]):
                if info.from_addr == self.username.lower():
                    continue
                same_domain = domain and info.from_addr.endswith(("@" + domain, "." + domain))
                if info.from_addr != address.lower() and not same_domain:
                    continue
                found.setdefault(info.message_id or f"{folder}:{info.uid}", info)
        return sorted(found.values(), key=_sort_key, reverse=True)

    def messages_to(self, address: str, since: date) -> list[MailInfo]:
        """Messages in the Sent folder addressed to `address` since a date, newest first."""
        if not self.sent:
            return []
        uids = self._search(self.sent, ["SINCE", imap_date(since), "TO", quote(address)])
        return sorted(self._fetch(self.sent, uids[-10:]), key=_sort_key, reverse=True)

    def bounces_for(self, address: str, since: date) -> list[MailInfo]:
        key = imap_date(since)
        if key not in self._bounce_cache:
            bounces: list[MailInfo] = []
            for folder in self.reply_folders:
                uids = set(self._search(folder, ["SINCE", key, "FROM", quote("mailer-daemon")]))
                uids |= set(self._search(folder, ["SINCE", key, "FROM", quote("postmaster")]))
                bounces += [m for m in self._fetch(folder, sorted(uids, key=int)[-50:]) if m.is_bounce]
            self._bounce_cache[key] = bounces
        address = address.lower()
        return [b for b in self._bounce_cache[key] if address in b.text.lower()]


def _sort_key(info: MailInfo) -> float:
    if info.when is None:
        return 0.0
    try:
        return info.when.timestamp()
    except (OverflowError, OSError, ValueError):
        return 0.0
