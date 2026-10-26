"""Who the emails come from, their signature, and the Spam Act 2003 (Cth) opt-out.

Every commercial email must identify the sender, say how to contact them, and
carry a working unsubscribe option that is honoured within 5 business days.

By default the sender name, From address and signature are learned from your
Sent folder, so drafts carry the signature you already use. `email.signature`
switches to the details in config.yaml ("config") or to no signature ("none").
The opt-out footer is always appended by code, never by the model.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .dedupe import host_of

_PLACEHOLDER = re.compile(r"<<.*?>>|edit me", re.I)
SIGNATURE_MODES = ("auto", "config", "none")


def _filled(value) -> str:
    text = str(value or "").strip()
    return "" if _PLACEHOLDER.search(text) else text


def config_sender_problems(cfg) -> list[str]:
    """What's missing for a signature built from config.yaml (email.signature: config)."""
    s = cfg["sender"]
    problems = [f"sender.{key}" for key in ("name", "business_name", "email") if not _filled(s.get(key))]
    if not (_filled(s.get("address")) or _filled(s.get("phone"))):
        problems.append("sender.address (or sender.phone)")
    return problems


def config_signature(cfg) -> str:
    s = {k: _filled(v) for k, v in cfg["sender"].items()}
    lines = [s["name"] + (f" | {s['title']}" if s.get("title") else "")]
    lines.append(s["business_name"] + (f" (ABN {s['abn']})" if s.get("abn") else ""))
    if s.get("address"):
        lines.append(s["address"])
    contact = " | ".join(x for x in (s.get("phone"), s.get("email"), s.get("website")) if x)
    if contact:
        lines.append(contact)
    return "\n".join(line for line in lines if line.strip())


def needs_unsubscribe(cfg, record_type: str) -> bool:
    return record_type in cfg["compliance"]["unsubscribe_footer_for"]


@dataclass
class Sender:
    name: str
    email: str
    signature: str
    signature_source: str  # "your Sent folder", "config.yaml" or "none"
    business_name: str
    unsubscribe_email: str
    problems: list[str] = field(default_factory=list)

    def unsubscribe_footer(self, source_url: str = "", record_type: str = "lead") -> str:
        if record_type == "quote":
            return (f"If you'd rather not receive requests like this from {self.business_name}, reply "
                    "\"unsubscribe\" and we won't contact you again.")
        where = host_of(source_url)
        reason = f"You're receiving this because your business email is published on {where}. " if where else ""
        return (reason + f"If you'd rather not hear from {self.business_name} again, reply with \"unsubscribe\" "
                f"(or email {self.unsubscribe_email} with the subject \"unsubscribe\") and we'll remove you "
                "within 5 business days.")

    def finish_body(self, cfg, body: str, *, record_type: str, source_url: str = "") -> str:
        text = body.rstrip()
        if self.signature:
            text += "\n\n" + self.signature
        if needs_unsubscribe(cfg, record_type):
            text += "\n\n--\n" + self.unsubscribe_footer(source_url, record_type)
        return text + "\n"


def resolve_sender(cfg, mailbox=None) -> Sender:
    """Sender details for this run. `mailbox` (connected) lets them be learned from your Sent folder."""
    s = cfg["sender"]
    mode = str(cfg["email"].get("signature", "auto")).lower()
    business = _filled(s.get("business_name")) or "OriginCrest"
    learned = mailbox.learn_sender() if (mailbox is not None and mode == "auto") else None
    username = str(cfg["email"].get("username", ""))
    address = (_filled(s.get("email")) or (learned.email if learned else "")
               or (username if "@" in username and not _PLACEHOLDER.search(username) else ""))
    learned_name = learned.name if learned and "@" not in learned.name else ""  # some apps put the address there
    name = _filled(s.get("name")) or learned_name
    problems: list[str] = []
    if mode not in SIGNATURE_MODES:
        problems.append(f"email.signature must be one of {', '.join(SIGNATURE_MODES)}")
    signature, source = "", "none"
    if mode == "auto" and learned and learned.signature:
        signature, source = learned.signature, "your Sent folder"
    elif mode in ("auto", "config") and not config_sender_problems(cfg):
        signature, source = config_signature(cfg), "config.yaml"
    elif mode == "config":
        problems += config_sender_problems(cfg)
    if not address:
        problems.append("the From address (set email.username to your email address, or sender.email)")
    unsubscribe = _filled(s.get("unsubscribe_email")) or address
    return Sender(name, address, signature, source, business, unsubscribe, problems)


def describe_sender(sender: Sender, cfg, report) -> None:
    if sender.problems:
        report.add("Errors", "No drafts until this is set in config.yaml: " + ", ".join(sender.problems))
    elif sender.signature_source == "none" and str(cfg["email"].get("signature", "auto")).lower() == "auto":
        report.add("Signature", "Couldn't find your signature in the Sent folder, so drafts have none - add it "
                                "when you send, or set email.signature: config")
    else:
        report.add("Signature", f"Drafts from {sender.name or '(no name)'} <{sender.email}>, signature from "
                                f"{sender.signature_source}")
