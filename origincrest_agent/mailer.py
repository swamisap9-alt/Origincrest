"""IMAP mailbox access: save drafts, find replies, sent mail and bounces.

The agent never sends email. It only APPENDs messages to the Drafts folder,
where you review, edit and send (or delete) them yourself.
"""

from __future__ import annotations

import email
import email.policy
import html
import imaplib
import re
import ssl
import time
from collections import Counter
from dataclasses import dataclass
from datetime import date, datetime
from email.message import EmailMessage
from email.utils import formataddr, formatdate, getaddresses, make_msgid, parsedate_to_datetime

from .utils import log

_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
_LIST_RE = re.compile(r'^\((?P<flags>[^)]*)\)\s+(?P<delim>"(?:[^"\\]|\\.)*"|NIL)\s+(?P<name>.+)$')
# Folder names to fall back on when the server doesn't flag its special folders
# (compared on the last part of the name, so "INBOX.Sent" and "INBOX/Sent Items" match too).
_FOLDER_NAMES = {
    "drafts": ["drafts", "draft"],
    "sent": ["sent", "sent items", "sent messages", "sent mail", "sent-mail", "outbox sent"],
}
SECURITY_MODES = ("ssl", "starttls", "none")
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
    text: str  # the message's own text, quoted history removed
    from_name: str = ""

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
            content = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", content, flags=re.S | re.I)
            content = re.sub(r"<br\s*/?>|</(p|div|tr|li|h\d)>", "\n", content, flags=re.I)
            content = html.unescape(re.sub(r"<[^>]+>", "", content))
            content = "\n".join(re.sub(r"[ \t\xa0]+", " ", line).strip() for line in content.splitlines())
        return content
    except (KeyError, LookupError, UnicodeDecodeError, AttributeError):
        return ""


def parse_message(raw: bytes, folder: str, uid: str, include_body: bool = True) -> MailInfo:
    msg = email.message_from_bytes(raw, policy=email.policy.default)
    try:
        when = parsedate_to_datetime(str(msg["Date"])) if msg["Date"] else None
    except (TypeError, ValueError):
        when = None
    from_name, from_addr = (getaddresses([str(msg.get("From", ""))]) or [("", "")])[0]
    from_addr = from_addr.lower()
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
    text = strip_quoted(_body_text(msg))[:8000] if include_body else ""
    return MailInfo(folder, uid, when, from_addr, to_addrs, subject, str(msg.get("Message-ID", "")).strip(),
                    str(msg.get("References", "")), auto, text, from_name.strip())


# --- Learning your signature from the Sent folder ---------------------------

_SIGNOFF = re.compile(
    r"^(kind|best|warm|warmest|many)?\s*(regards|thanks|thank you|wishes|cheers)\b.{0,20}$|"
    r"^(cheers|thanks|thank you|sincerely|yours sincerely|yours faithfully|yours truly|best|all the best|"
    r"talk soon|speak soon)\b[ ,.!]*$", re.I)
_AGENT_FOOTER = re.compile(r"\n--\n[^\n]*(unsubscribe|requests like this)", re.I)
_CONTACT_LINE = re.compile(r"@|www\.|https?://|\+?\d[\d\s().-]{6,}\d", re.I)


@dataclass
class LearnedSender:
    name: str
    email: str
    signature: str
    messages_checked: int


def _norm_line(line: str) -> str:
    return re.sub(r"\s+", " ", line).strip().lower()


def learn_signature(texts: list[str]) -> str:
    """Your signature: the block that ends most of your sent emails, after the sign-off line.

    `texts` are your own message texts (quoted history removed), newest first.
    """
    docs: list[list[str]] = []
    for text in texts:
        text = _AGENT_FOOTER.split(text)[0]  # emails the agent drafted end with its opt-out footer
        lines = [line.rstrip() for line in text.splitlines()]
        while lines and not lines[-1].strip():
            lines.pop()
        if lines:
            docs.append(lines)
    if len(docs) < 2:
        return ""

    def tidy(lines: list[str]) -> str:
        while lines and not lines[0].strip():
            lines = lines[1:]
        return "\n".join(lines).strip()

    # 1. Most emails have a sign-off ("Kind regards,"); the signature is what follows it.
    after_signoff: list[str] = []
    for lines in docs:
        tail = lines[-30:]
        idx = max((i for i, line in enumerate(tail) if _SIGNOFF.match(line.strip())), default=None)
        if idx is not None and idx < len(tail) - 1:
            after_signoff.append(tidy(tail[idx + 1:]))
    votes = Counter(_norm_line(sig.replace("\n", " | ")) for sig in after_signoff if sig)
    for sig in after_signoff:  # newest first: a changed signature wins once it's in two emails
        if sig and votes[_norm_line(sig.replace("\n", " | "))] >= 2:
            return sig

    # 2. Otherwise: the longest run of closing lines shared by several emails, if it looks like a signature.
    need = max(2, (len(docs) + 2) // 3)
    for k in range(12, 1, -1):
        blocks = Counter(tuple(_norm_line(l) for l in lines[-k:] if l.strip())
                         for lines in docs if len([l for l in lines if l.strip()]) >= k)
        if not blocks:
            continue
        block, count = blocks.most_common(1)[0]
        if count >= need and len(block) >= 2 and any(_CONTACT_LINE.search(l) for l in block) \
                and all(len(l) <= 90 for l in block):
            for lines in docs:
                if tuple(_norm_line(l) for l in lines[-k:] if l.strip()) == block:
                    sig_lines = lines[-k:]
                    while sig_lines and (not sig_lines[0].strip() or _SIGNOFF.match(sig_lines[0].strip())):
                        sig_lines = sig_lines[1:]
                    return tidy(sig_lines)
    return ""


def learn_sender(messages: list["MailInfo"]) -> LearnedSender | None:
    """Your From name/address and signature, from your recent sent emails (newest first)."""
    mine = [m for m in messages if m.from_addr]
    if not mine:
        return None
    address = Counter(m.from_addr for m in mine).most_common(1)[0][0]
    names = Counter(m.from_name for m in mine if m.from_addr == address and m.from_name)
    own = [m.text for m in mine if m.from_addr == address and m.text]
    return LearnedSender(names.most_common(1)[0][0] if names else "", address, learn_signature(own), len(own))


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
                 security: str = "ssl", own_addresses: list[str] | None = None):
        if security not in SECURITY_MODES:
            raise ValueError(f"email.security must be one of {', '.join(SECURITY_MODES)}")
        self.host, self.port, self.username, self.password = host, port, username, password
        self.security = security
        # Messages from these addresses are ours, never "replies".
        self.own_addresses = {a.strip().lower() for a in [username, *(own_addresses or [])] if a and "@" in a}
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
        missing = [name for name, value in (("email.imap_host", e["imap_host"]), ("email.username", e["username"]))
                   if not value or "<<" in str(value)]
        if missing or not password:
            raise RuntimeError("Email not configured: set " + ", ".join(missing + (
                [] if password else [f"{e['password_env']} in .env"])))
        return cls(e["imap_host"], int(e["imap_port"]), e["username"], password,
                   drafts_folder=e["drafts_folder"], sent_folder=e["sent_folder"],
                   reply_folders=e["reply_folders"], max_fetch_bytes=int(e["max_fetch_bytes"]),
                   security=str(e.get("security", "ssl")).lower(),
                   own_addresses=[cfg["sender"].get("email", ""), cfg["sender"].get("unsubscribe_email", "")])

    # connection --------------------------------------------------------
    def __enter__(self) -> "Mailbox":
        self.connect()
        return self

    def __exit__(self, *exc) -> None:
        self.close()

    def connect(self) -> None:
        context = ssl.create_default_context()
        if self.security == "ssl":
            self.imap = imaplib.IMAP4_SSL(self.host, self.port, ssl_context=context, timeout=60)
        else:
            self.imap = imaplib.IMAP4(self.host, self.port, timeout=60)
            if self.security == "starttls":
                self.imap.starttls(ssl_context=context)
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
            if isinstance(line, tuple):  # name sent as a literal: (b'(\\Flags) "." {12}', b'INBOX.Drafts')
                line = re.sub(rb"\{\d+\}$", b"", line[0]).strip() + b' "' + line[1] + b'"'
            m = _LIST_RE.match(line.decode("utf-8", "replace") if isinstance(line, bytes) else str(line))
            if not m:
                continue
            name = m.group("name").strip()
            if name.startswith('"') and name.endswith('"'):
                name = name[1:-1].replace('\\"', '"').replace("\\\\", "\\")
            self.folders[name] = {f.lower() for f in m.group("flags").split()}

        def special(kind: str) -> str:
            candidates = self.folder_candidates(kind)
            return candidates[0] if candidates else ""

        self.drafts = self.drafts or special("drafts")
        self.sent = self.sent or special("sent")
        if not self.reply_folders:
            all_mail = special("all")
            self.reply_folders = [all_mail] if all_mail else ["INBOX"]
        if not self.drafts:
            raise RuntimeError("Couldn't find the Drafts folder; set email.drafts_folder in config.yaml")
        log.debug("IMAP folders: drafts=%s sent=%s replies=%s", self.drafts, self.sent, self.reply_folders)

    def folder_candidates(self, kind: str) -> list[str]:
        """Folders that look like Drafts / Sent / All mail: server-flagged ones first, then by name."""
        flag = {"drafts": "\\drafts", "sent": "\\sent", "all": "\\all"}[kind]
        names = [name for name, flags in self.folders.items() if flag in flags]
        for wanted in _FOLDER_NAMES.get(kind, []):
            names += [n for n in self.folders
                      if re.split(r"[./]", n)[-1].strip().lower() == wanted and n not in names]
        return names

    def message_count(self, folder: str) -> int:
        typ, data = self.imap.select(quote(folder), readonly=True)
        try:
            return int(data[0]) if typ == "OK" else -1
        except (TypeError, ValueError, IndexError):
            return -1

    def agent_drafts(self, folder: str = "") -> list[MailInfo]:
        """Drafts the agent saved (they carry an X-OriginCrest-Agent header), newest first."""
        folder = folder or self.drafts
        uids = self._search(folder, ["HEADER", "X-OriginCrest-Agent", '""'])
        return sorted(self._fetch(folder, uids[-50:], include_body=False), key=_sort_key, reverse=True)

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
                if info.from_addr in self.own_addresses:
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

    def recent_sent(self, limit: int = 25) -> list[MailInfo]:
        if not self.sent:
            return []
        uids = self._search(self.sent, ["ALL"])
        return sorted(self._fetch(self.sent, uids[-limit:]), key=lambda m: (_sort_key(m), int(m.uid)), reverse=True)

    def learn_sender(self) -> LearnedSender | None:
        return learn_sender(self.recent_sent())

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
