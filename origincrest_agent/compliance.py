"""Spam Act 2003 (Cth) guardrails for outreach emails.

Every commercial email must identify the sender, say how to contact them, and
carry a working unsubscribe option that is honoured within 5 business days.
The footer below is appended by code (never by the model), and drafting is
refused until the sender details in config.yaml are filled in.
"""

from __future__ import annotations

import re

from .dedupe import host_of

_PLACEHOLDER = re.compile(r"<<.*?>>|edit me", re.I)


def sender_problems(cfg) -> list[str]:
    s = cfg["sender"]
    problems = []
    for key in ("name", "business_name", "email"):
        if not s.get(key) or _PLACEHOLDER.search(str(s[key])):
            problems.append(f"sender.{key}")
    if not (s.get("address") or s.get("phone")) or _PLACEHOLDER.search(str(s.get("address", ""))):
        problems.append("sender.address (or sender.phone)")
    return problems


def unsubscribe_address(cfg) -> str:
    return cfg["sender"].get("unsubscribe_email") or cfg["sender"]["email"]


def signature(cfg) -> str:
    s = cfg["sender"]
    lines = [s["name"] + (f" | {s['title']}" if s.get("title") else "")]
    lines.append(s["business_name"] + (f" (ABN {s['abn']})" if s.get("abn") else ""))
    if s.get("address"):
        lines.append(s["address"])
    contact = " | ".join(x for x in (s.get("phone"), s.get("email"), s.get("website")) if x)
    if contact:
        lines.append(contact)
    return "\n".join(lines)


def unsubscribe_footer(cfg, source_url: str = "", record_type: str = "lead") -> str:
    if record_type == "quote":
        return (f"If you'd rather not receive requests like this from {cfg['sender']['business_name']}, reply "
                "\"unsubscribe\" and we won't contact you again.")
    where = host_of(source_url)
    reason = f"You're receiving this because your business email is published on {where}. " if where else ""
    return (reason + f"If you'd rather not hear from {cfg['sender']['business_name']} again, reply with "
            f"\"unsubscribe\" (or email {unsubscribe_address(cfg)} with the subject \"unsubscribe\") and "
            "we'll remove you within 5 business days.")


def needs_unsubscribe(cfg, record_type: str) -> bool:
    return record_type in cfg["compliance"]["unsubscribe_footer_for"]


def finish_body(cfg, body: str, *, record_type: str, source_url: str = "") -> str:
    text = body.rstrip() + "\n\n" + signature(cfg)
    if needs_unsubscribe(cfg, record_type):
        text += "\n\n--\n" + unsubscribe_footer(cfg, source_url, record_type)
    return text + "\n"
