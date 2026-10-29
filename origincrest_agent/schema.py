"""Canonical spreadsheet fields and the header names they are recognised by.

Existing sheets keep their own header names: a column titled "Company Name" or
"Contact Person" is matched to the canonical field, and only fields with no
matching column are added (at the right-hand end of the header row).
"""

from __future__ import annotations

import re
from dataclasses import dataclass


def norm_header(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(text).lower()).strip()


@dataclass(frozen=True)
class Field:
    key: str
    header: str
    synonyms: tuple[str, ...] = ()
    width: int = 18
    is_date: bool = False
    # Matched only when nothing more specific claimed the field (e.g. a bare "Name").
    weak_synonyms: tuple[str, ...] = ()


FIELDS: list[Field] = [
    Field("company", "Company", (
        "company", "company name", "business", "business name", "organisation", "organization",
        "organisation name", "organization name", "forwarder", "freight forwarder", "3pl", "3pl name",
        "provider", "provider name", "supplier", "supplier name", "roaster", "roaster name", "roastery",
        "buyer", "buyer name", "customer", "account", "account name", "trading name", "company business",
    ), width=30, weak_synonyms=("name",)),
    Field("category", "Category", (
        "category", "type", "business type", "service type", "segment", "company type", "provider type",
        "business category", "category type", "sector", "industry", "prospect type", "contact type",
        "classification", "categories",
    ), width=24),
    Field("contact_name", "Contact name", (
        "contact name", "contact person", "key contact", "primary contact", "person", "attention", "attn",
        "poc", "contact first name", "full name",
    ), width=22, weak_synonyms=("contact",)),
    Field("contact_role", "Role", (
        "role", "title", "position", "job title", "designation", "contact role", "contact title",
    ), width=18),
    Field("email", "Email", (
        "email", "e mail", "email address", "contact email", "mail", "email id", "emails", "e mail address",
    ), width=30),
    Field("phone", "Phone", (
        "phone", "phone number", "telephone", "tel", "mobile", "contact number", "ph", "phone no",
        "contact no", "phone mobile", "phone numbers",
    ), width=16),
    Field("website", "Website", (
        "website", "web", "url", "site", "web site", "homepage", "web address", "website url",
    ), width=28),
    Field("city", "City", (
        "city", "suburb", "location", "town", "city suburb", "office", "office location", "branch", "base",
        "city town",
    ), width=16),
    Field("state", "State", ("state", "state territory", "territory"), width=8),
    Field("services", "Services / offering", (
        "services", "service", "services offering", "services offered", "offering", "capabilities",
        "specialty", "speciality", "what they do", "description", "coffee handling",
    ), width=40),
    Field("talking_points", "Talking points", (
        "talking points", "email notes", "notes for email", "message notes", "pitch notes", "personal note",
        "email brief",
    ), width=40),
    Field("source_url", "Source URL", (
        "source url", "source", "source link", "found at", "reference url", "where found", "source website",
    ), width=34),
    Field("status", "Status", (
        "status", "stage", "rfq status", "lead status", "quote status", "pipeline status", "current status",
    ), width=22),
    Field("date_added", "Date added", (
        "date added", "added on", "date created", "created", "date found",
    ), width=12, is_date=True),
    Field("added_by", "Added by", ("added by", "created by", "entered by"), width=16),
    Field("record_type", "Record type", (
        "record type", "quote lead", "quote or lead", "kind",
    ), width=12),
    Field("last_contacted", "Last contacted", (
        "last contacted", "last contact", "last contact date", "date contacted", "contacted on",
        "date of last contact", "last emailed", "last contacted date",
    ), width=14, is_date=True),
    # Your own "Date sent": read, never overwritten (the agent records its follow-ups in Last contacted).
    Field("date_sent", "Date sent", (
        "date sent", "sent date", "sent on", "rfq sent", "rfq sent date", "date rfq sent", "date requested",
        "rfq date", "date emailed",
    ), width=14, is_date=True),
    # Your own notes: read (e.g. "need to be called"), never written.
    Field("user_notes", "Notes", ("notes", "note", "comments", "comment", "remarks", "my notes"), width=40),
    Field("next_action_date", "Next action date", (
        "next action date", "follow up date", "next follow up", "next follow up date", "follow up on",
        "followup date", "chase date", "next step date", "next contact date", "due date",
    ), width=14, is_date=True),
    Field("last_reply", "Last reply", (
        "last reply", "reply date", "response date", "last response", "date replied",
        "reply received", "last reply date",
    ), width=14, is_date=True),
    Field("followup_count", "Follow-ups drafted", (
        "follow ups drafted", "follow ups", "follow up count", "followups", "no of follow ups",
        "number of follow ups",
    ), width=10),
    Field("agent_notes", "Agent notes", ("agent notes", "automation notes", "agent note"), width=48),
    # Lead research (coffee-business rows only).
    Field("roast_profile", "Coffee offering", (
        "coffee offering", "what they roast", "what they roast sell", "roast profile", "coffee range",
    ), width=40),
    Field("size", "Size estimate", ("size estimate", "size", "company size", "scale"), width=16),
    Field("sourcing", "Sourcing", ("sourcing", "sourcing approach", "green sourcing"), width=34),
    Field("indonesian_now", "Current Indonesian coffees", (
        "current indonesian coffees", "indonesian offering", "indonesian coffees",
    ), width=30),
    Field("fit_score", "Fit score (0-100)", ("fit score 0 100", "fit score", "lead score", "score"), width=10),
    Field("fit_origins", "Best-fit origins", ("best fit origins", "target origins"), width=20),
    Field("fit_rationale", "Fit rationale", ("fit rationale", "fit notes", "why fit"), width=48),
    Field("researched_on", "Researched on", ("researched on", "research date"), width=12, is_date=True),
    Field("draft_created", "Draft created", (
        "draft created", "intro draft created", "draft date",
    ), width=12, is_date=True),
    Field("draft_subject", "Draft subject", ("draft subject", "intro draft subject"), width=34),
]

FIELD_BY_KEY: dict[str, Field] = {f.key: f for f in FIELDS}

# Columns written for every row the discovery workflow adds.
DISCOVERY_KEYS = [
    "company", "category", "contact_name", "contact_role", "email", "phone", "website", "city", "state",
    "services", "source_url", "status", "date_added", "added_by", "record_type", "agent_notes",
]
FOLLOWUP_KEYS = ["record_type", "last_contacted", "next_action_date", "last_reply", "followup_count", "agent_notes"]
LEAD_KEYS = [
    "roast_profile", "size", "sourcing", "indonesian_now", "fit_score", "fit_origins", "fit_rationale",
    "researched_on", "draft_created", "draft_subject",
]

# Header layout for tabs the agent creates ("Talking points" is yours: notes the emails should build on).
# "buyer" = coffee businesses (roasters, importers, wholesalers, traders...); "freight" = logistics.
_NEW_TAB_BASE = DISCOVERY_KEYS[:10] + ["talking_points"] + DISCOVERY_KEYS[10:-1]
NEW_TAB_KEYS = {
    "buyer": _NEW_TAB_BASE + ["last_contacted", "next_action_date", "last_reply", "followup_count"]
    + LEAD_KEYS + ["agent_notes"],
    "freight": _NEW_TAB_BASE + ["last_contacted", "next_action_date", "last_reply", "followup_count",
                                "draft_created", "draft_subject", "agent_notes"],
}

_FREIGHT_WORDS = re.compile(
    r"freight|forward|3pl|logistic|warehous|customs|broker|shipping|transport|fumigat|courier|carrier|"
    r"storage|cold chain|devan|container|fulfil", re.I)


def category_kind(category: str) -> str:
    """'freight' for logistics categories (they get RFQs), 'buyer' for everything else (intro emails)."""
    return "freight" if _FREIGHT_WORDS.search(category or "") else "buyer"

# Agent bookkeeping columns that aren't useful to whoever writes the email.
NOT_FOR_EMAIL = ["status", "source_url", "date_added", "added_by", "record_type", "agent_notes", "researched_on",
                 "draft_created", "draft_subject", "followup_count", "fit_score"]

_STRONG: dict[str, str] = {}
_WEAK: dict[str, str] = {}
for _f in FIELDS:
    for _s in (_f.key, _f.header, *_f.synonyms):
        _STRONG.setdefault(norm_header(_s), _f.key)
    for _s in _f.weak_synonyms:
        _WEAK.setdefault(norm_header(_s), _f.key)


def match_header(text: str) -> str | None:
    return _STRONG.get(norm_header(text))


def match_weak_header(text: str) -> str | None:
    return _WEAK.get(norm_header(text))
