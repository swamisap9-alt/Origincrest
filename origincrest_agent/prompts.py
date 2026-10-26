"""Prompts and output schemas for the three workflows."""

from __future__ import annotations

import json
import re

from .llm import obj, strict_tool

STATES = ["SA", "VIC", "NSW", "QLD", "WA", "TAS", "ACT", "NT", ""]
ORIGINS = ["Sumatra", "Sulawesi", "Java", "Flores"]

CONTACT_RULES = """\
Rules for contact details (strict):
- Record only business contact details that are published on a web page you actually viewed: the \
business's own site or a business directory listing. Give that page as the source URL.
- Never guess, construct or pattern-fill an email address (e.g. firstname@domain). If no email is \
published, leave it empty.
- Prefer role-based business addresses (info@, sales@, wholesale@, hello@, enquiries@). Leave out \
private individuals' details and personal social-media profiles.
- If a page says the business doesn't want unsolicited commercial email, leave the email empty and say \
so in the notes.
- Leave any field you couldn't confirm as an empty string. Accuracy matters more than completeness."""


DEFAULT_SUMMARY = ("OriginCrest supplies Indonesian single-origin green coffee - Sumatra, Sulawesi, Java and "
                   "Flores - to Australian roasters.")


def _filled(text) -> str:
    """Config text with any unfilled <<EDIT ME>> placeholder removed."""
    text = re.sub(r"<<.*?>>", "", str(text or ""))
    return re.sub(r"[\s\-:]+$", "", text).strip()


def offer_text(cfg) -> str:
    """Background about OriginCrest. Optional in config.yaml; a default is used when it's blank."""
    offer = cfg["offer"]
    lines = [f"Business: {cfg['sender']['business_name']}",
             f"Summary: {_filled(offer.get('summary')) or DEFAULT_SUMMARY}"]
    origins = [o for o in (_filled(x) for x in offer.get("origins") or []) if o]
    if origins:
        lines.append("Origins offered: " + "; ".join(origins))
    for fact in filter(None, (_filled(f) for f in offer.get("facts") or [])):
        lines.append(f"Fact: {fact}")
    if _filled(offer.get("call_to_action")):
        lines.append(f"Preferred next step: {_filled(offer['call_to_action'])}")
    return "\n".join(lines)


# --- Workflow 1: discovery -------------------------------------------------

DISCOVERY_SYSTEM = f"""\
You find Australian businesses for OriginCrest, an importer of Indonesian single-origin green coffee \
(Sumatra, Sulawesi, Java, Flores). Each request names one category from OriginCrest's contact sheets: \
either coffee businesses that could buy or trade OriginCrest's coffee (roasters, importers, wholesalers, \
traders and similar), or logistics providers that can move and store coffee (freight forwarders, customs \
brokers, 3PLs). Search the web, then record what you found with the record_companies tool.

{CONTACT_RULES}

Only include businesses that are in Australia, appear to be operating now, and genuinely belong to the \
requested category. Skip anything in the "already known" list, including obvious variants of the same name."""

COMPANY_SCHEMA = obj({
    "company_name": {"type": "string"},
    "business_type": {"type": "string", "description": "Short label, e.g. 'Specialty roaster', 'Freight forwarder'"},
    "website": {"type": "string", "description": "Company's own website URL, or empty"},
    "contact_name": {"type": "string", "description": "Named business contact if published, else empty"},
    "contact_role": {"type": "string"},
    "email": {"type": "string", "description": "Published business email, else empty"},
    "phone": {"type": "string", "description": "Published business phone, else empty"},
    "city": {"type": "string", "description": "City or suburb of the relevant office"},
    "state": {"type": "string", "enum": STATES},
    "adelaide_presence": {"type": "boolean", "description": "Based in Adelaide/SA or has an Adelaide office"},
    "services": {"type": "string", "description": "One or two sentences: what they sell / roast / which services"},
    "coffee_evidence": {"type": "string", "description": "Evidence they deal in coffee / handle coffee freight"},
    "source_url": {"type": "string", "description": "Page where the contact details are published"},
    "notes": {"type": "string"},
})

RECORD_COMPANIES = strict_tool(
    "record_companies",
    "Record the businesses you found. Call once, after you've finished searching.",
    obj({"companies": {"type": "array", "items": COMPANY_SCHEMA}}),
)

_FREIGHT_TASK = """\
They must handle green coffee imports and/or roasted coffee storage and distribution. Look for evidence \
they handle coffee or similar food and agricultural imports: FCL/LCL sea freight from Indonesia \
(Belawan/Medan, Surabaya, Makassar, Jakarta), customs clearance and DAFF biosecurity for green coffee, \
fumigation, container unpacking, food-grade or bonded warehousing, pallet storage, pick-pack fulfilment and \
interstate distribution. Put that evidence in coffee_evidence."""

_BUYER_TASK = """\
They should be businesses that buy, roast, import, wholesale or trade coffee and could take Indonesian \
single-origin green coffee from OriginCrest. Specialty-grade focus is a plus. Café chains that don't roast, \
machine sellers and instant-coffee brands don't count unless the category says so. Put what shows they \
deal in coffee in coffee_evidence."""


def discovery_prompt(*, kind: str, category: str, region: str, focus: list[str], n: int, known: list[str],
                     examples: list[str]) -> str:
    if region == "adelaide":
        where = "based in Adelaide or elsewhere in South Australia, or with an Adelaide office"
    else:
        where = "in the other Australian states and territories (outside South Australia)"
    lines = [f'Find up to {n} businesses in the category "{category}" {where}.',
             _FREIGHT_TASK if kind == "freight" else _BUYER_TASK]
    if focus:
        lines.append(f"Focus this week: {'; '.join(focus)}.")
    if examples:
        lines.append("Businesses already in OriginCrest's sheet under this category (find others like them): "
                     + ", ".join(examples))
    lines.append(f"Already known (skip these): {', '.join(known) if known else '(none yet)'}")
    lines.append("When you've finished searching, call record_companies once with everything that qualifies.")
    return "\n\n".join(lines)


# --- Workflow 2: research, intro emails and RFQs ---------------------------

RESEARCH_SYSTEM = f"""\
You research Australian coffee businesses (roasters, importers, wholesalers, traders and similar) for \
OriginCrest and judge how well each one fits OriginCrest's Indonesian single-origin green coffees. Use web \
search and web fetch on the business's own website and online shop first, then press, awards and \
directories. Record your findings with the record_research tool.

Score fit from 0 to 100:
- Specialty focus and single-origin range (up to 30)
- Whether they buy, import or distribute green coffee in volumes OriginCrest can serve (up to 25)
- Indonesian interest: they sell Indonesian coffees now, have before, or have a gap an Indonesian lot \
would fill (up to 25)
- Practicalities: wholesale/café supply, scale, location - Adelaide/SA is a plus (up to 20)
Importers and traders can be partners (buying or distributing OriginCrest lots) or competitors: score how \
likely a partnership is. Be calibrated: 70+ needs clear evidence of specialty single-origin buying; under \
40 for businesses that don't deal in coffee, closed businesses, or commodity/instant focus.

personalisation_hooks are 2-3 specific facts from the business's own materials that an intro email could \
mention (a current single origin, their roasting or sourcing approach, an award), each with its source URL. \
Use business facts only - nothing personal about individuals.

{CONTACT_RULES}"""

RESEARCH_SCHEMA = obj({
    "company_name": {"type": "string"},
    "website": {"type": "string"},
    "city": {"type": "string"},
    "state": {"type": "string", "enum": STATES},
    "what_they_roast": {"type": "string", "description": "What they roast / sell / import: range, style, blends"},
    "size_estimate": {"type": "string", "enum": ["Micro", "Small", "Medium", "Large", "Unknown"]},
    "size_evidence": {"type": "string"},
    "sourcing": {"type": "string", "description": "How they buy green: direct trade, named importers..."},
    "current_indonesian_coffees": {"type": "string", "description": "Indonesian coffees they sell now, or empty"},
    "wholesale": {"type": "string", "enum": ["Yes", "No", "Unknown"]},
    "public_email": {"type": "string"},
    "public_email_source_url": {"type": "string"},
    "contact_name": {"type": "string"},
    "contact_role": {"type": "string"},
    "fit_score": {"type": "integer"},
    "best_fit_origins": {"type": "array", "items": {"type": "string", "enum": ORIGINS}},
    "fit_rationale": {"type": "string", "description": "Two or three sentences"},
    "personalisation_hooks": {"type": "array", "items": obj({
        "fact": {"type": "string"}, "source_url": {"type": "string"}})},
    "red_flags": {"type": "string", "description": "Closed, doesn't deal in coffee, etc. Empty if none"},
})

RECORD_RESEARCH = strict_tool(
    "record_research",
    "Record what you found about the business. Call once, after researching.",
    RESEARCH_SCHEMA,
)


def research_prompt(row_context: dict[str, str], cfg) -> str:
    return f"""Research this business from our contacts sheet and score its fit.

Row from the sheet:
{json.dumps(row_context, indent=2, ensure_ascii=False)}

What OriginCrest offers:
{offer_text(cfg)}

When you've finished, call record_research once."""


EMAIL_SCHEMA = obj({
    "subject": {"type": "string"},
    "body": {"type": "string"},
})

INTRO_SYSTEM = """\
You write short first-contact emails from OriginCrest to Australian coffee businesses (roasters, \
importers, wholesalers, traders and similar). Each email is written for one contact, from their row in \
OriginCrest's contacts spreadsheet and the web research on their business.

Tailor it to the contact: use their name and role if the row has them, and build on anything OriginCrest \
noted about them in the row (notes, talking points, how they met, category, services). Open with something \
specific about their business from the research hooks, connect it to one or two OriginCrest origins that \
suit them (for an importer or trader, frame it as supply or partnership), and offer one easy next step. Ask \
at most one question.

Write like a person who knows coffee, not a marketing department: Australian English, plain text, \
110-170 words, no hype, no emojis, no bullet lists.

Use only facts given to you: the spreadsheet row and research for the business, and the OriginCrest \
background for OriginCrest. Don't state prices, certifications, volumes or availability unless they're \
given. End the body with a short sign-off such as "Kind regards," - the signature and unsubscribe line are \
added automatically, so don't write them. Subject: under 8 words, specific, not clickbait."""


def intro_prompt(research: dict, row_context: dict[str, str], contact_name: str, cfg) -> str:
    greeting = f"Address it to {contact_name.split()[0]}." if contact_name else "Open with 'Hi there,' or the team's name."
    return f"""Write the intro email for this contact. {greeting}

Their row in OriginCrest's contacts spreadsheet:
{json.dumps(row_context, indent=2, ensure_ascii=False)}

Research on the roaster:
{json.dumps(research, indent=2, ensure_ascii=False)}

OriginCrest background:
{offer_text(cfg)}"""


RFQ_SYSTEM = """\
You write requests for quotation (RFQs) from OriginCrest, an importer of Indonesian green coffee, to \
Australian freight forwarders, customs brokers and 3PL / warehouse operators. Each RFQ is written for one \
company, from its row in OriginCrest's contacts spreadsheet.

Tailor it to the company: ask only for the services they actually offer (from the row: services, category, \
evidence they handle coffee), mention their location or lane where relevant (e.g. Port Adelaide), and use \
the contact's name if the row has one. Make it easy to quote: say what the cargo is (green coffee, usually \
in jute or GrainPro bags), the lane or service, and list what the quote should cover.

Take the shipment details (origins and ports, destination, container size, volumes, frequency, incoterms, \
services needed) from OriginCrest's requirements and examples of earlier RFQs. Use them only to understand \
what OriginCrest needs; never mention the other companies in them. If a detail you need isn't given, write \
a clearly marked placeholder in square brackets for OriginCrest to fill in before sending, e.g. \
[number of 20ft containers per year]. Never invent numbers.

Plain text, Australian English, 120-200 words, professional and direct; a short list of what to quote is \
fine. End the body with a short sign-off such as "Kind regards," - the signature and opt-out line are added \
automatically, so don't write them. Subject: under 10 words, e.g. "RFQ: green coffee FCL Belawan to \
Adelaide"."""


def rfq_prompt(row_context: dict[str, str], requirements: str, examples: list[dict[str, str]],
               earlier_rfq: str, contact_name: str, cfg) -> str:
    greeting = f"Address it to {contact_name.split()[0]}." if contact_name else "Open with 'Hi team,'."
    parts = [f"Write the RFQ for this company. {greeting}",
             "Their row in OriginCrest's contacts spreadsheet:\n" + json.dumps(row_context, indent=2, ensure_ascii=False),
             "OriginCrest's requirements:\n" + (requirements or "(none written down - use the examples below)")]
    if examples:
        parts.append("Rows from OriginCrest's RFQ tracker (what it has asked other providers to quote):\n"
                     + json.dumps(examples, indent=2, ensure_ascii=False))
    if earlier_rfq:
        parts.append("An RFQ OriginCrest sent before (for reference only):\n" + earlier_rfq)
    parts.append("OriginCrest background:\n" + offer_text(cfg))
    return "\n\n".join(parts)


# --- Workflow 3: follow-ups -------------------------------------------------

FOLLOWUP_SYSTEM = """\
You write brief, polite follow-up emails for OriginCrest, an importer of Indonesian green coffee.

Two situations:
- quote: OriginCrest asked a freight forwarder, customs broker or 3PL for a quote and hasn't heard back. \
Briefly restate what was requested (from the row details and the original email), ask whether they can \
quote or need anything else, and mention timing if the details give it.
- lead: OriginCrest emailed a coffee business (roaster, importer, wholesaler, trader) and hasn't heard back. Keep it lighter than the first email, add \
one new useful thing (e.g. offer to send samples, mention an origin that suits them) and give an easy out.

Plain text, Australian English, 50-110 words; later follow-ups are shorter. Use only facts in the row \
details, the original email and the OriginCrest background. End with a short sign-off such as "Thanks," - \
the signature and any unsubscribe line are added automatically, so don't write them. If you are told \
the email continues an existing thread, write the body only; the subject is set for you, so return the \
subject you'd use anyway."""


def followup_prompt(*, record_type: str, number: int, row_context: dict[str, str], original: str,
                    days_since: int, threaded: bool, cfg) -> str:
    thread_note = "This continues the original email thread." if threaded else "There is no earlier thread to reply to."
    return f"""Write follow-up #{number} ({record_type}). It's been {days_since} days since we last \
contacted them with no reply. {thread_note}

Row details:
{json.dumps(row_context, indent=2, ensure_ascii=False)}

Our last email to them:
{original or "(not found in Sent - rely on the row details)"}

OriginCrest background:
{offer_text(cfg)}"""
