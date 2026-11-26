"""Prompts and output schemas for the three workflows."""

from __future__ import annotations

import json
import re

from .llm import obj, strict_tool

STATES = ["SA", "VIC", "NSW", "QLD", "WA", "TAS", "ACT", "NT", ""]

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


DEFAULT_SUMMARY = ("OriginCrest supplies Indonesian single-origin coffee - Sumatra, Sulawesi, Java and Flores - "
                   "to Australian businesses: green beans for roasters, and roasted coffee for cafés, "
                   "restaurants, retailers and other businesses.")
RANGE_SUMMARY = ("OriginCrest supplies Indonesian coffee to Australian businesses: green beans for roasters, "
                 "importers and traders, and roasted coffee for cafés, restaurants, retailers and other "
                 "businesses. Its current range is listed below.")
_GREEN_OR_ROASTED = """\
OriginCrest sells both green and roasted coffee. Roasters, importers, traders and green-bean resellers buy \
green; cafés, restaurants, grocers and other retailers, office-coffee providers, gifting companies and \
online coffee shops buy roasted. Some buy both."""


def _filled(text) -> str:
    """Config text with any unfilled <<EDIT ME>> placeholder removed."""
    text = re.sub(r"<<.*?>>", "", str(text or ""))
    return re.sub(r"[\s\-:]+$", "", text).strip()


def offer_text(cfg, range_chars: int | None = None) -> str:
    """Background about OriginCrest. Optional in config.yaml; a default is used when it's blank.

    Includes the current range (offer.range_sources) when one was loaded, cut to `range_chars`.
    """
    offer = cfg["offer"]
    current = getattr(cfg, "offer_range", "") or ""
    lines = [f"Business: {cfg['sender']['business_name']}",
             f"Summary: {_filled(offer.get('summary')) or (RANGE_SUMMARY if current else DEFAULT_SUMMARY)}"]
    origins = [o for o in (_filled(x) for x in offer.get("origins") or []) if o]
    if origins:
        lines.append("Origins offered: " + "; ".join(origins))
    for fact in filter(None, (_filled(f) for f in offer.get("facts") or [])):
        lines.append(f"Fact: {fact}")
    if _filled(offer.get("call_to_action")):
        lines.append(f"Preferred next step: {_filled(offer['call_to_action'])}")
    if current:
        if range_chars is not None and len(current) > range_chars:
            current = current[:range_chars].rsplit("\n", 1)[0] + "\n(... more in the full range)"
        lines.append("Current range (from OriginCrest's own website / catalogue - the only source for what it "
                     "sells; don't quote prices from it unless the instructions allow):\n" + current)
    return "\n".join(lines)


# --- Workflow 1: discovery -------------------------------------------------

DISCOVERY_SYSTEM = f"""\
You find Australian businesses for OriginCrest, which imports and roasts Indonesian coffee (each request \
describes its current range). Each request names one category from OriginCrest's contact sheets: \
either businesses that could buy or trade OriginCrest's green or roasted coffee (roasters, importers, \
traders, cafés, restaurants, retailers and similar), or logistics providers that can move and store coffee \
(freight forwarders, customs \
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
They should be businesses that could buy OriginCrest's coffee: green beans (roasters - specialty single \
origin or blends that could use what the range offers, e.g. robusta for espresso - importers, traders, \
green-bean resellers) or roasted coffee (cafés, restaurants, grocers and other retailers, office-coffee \
providers, gifting companies, online coffee shops). Machine-only sellers and instant-coffee brands don't count. \
Put what shows they buy, sell or serve coffee in coffee_evidence."""


def discovery_prompt(*, kind: str, category: str, region: str, focus: list[str], n: int, known: list[str],
                     examples: list[str], offer: str = "") -> str:
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
    if offer and kind != "freight":
        lines.append("What OriginCrest offers:\n" + offer)
    lines.append(f"Already known (skip these): {', '.join(known) if known else '(none yet)'}")
    lines.append("When you've finished searching, call record_companies once with everything that qualifies.")
    return "\n\n".join(lines)


# --- Workflow 2: research, intro emails and RFQs ---------------------------

RESEARCH_SYSTEM = f"""\
You research Australian businesses that buy coffee (roasters, importers, traders, cafés, restaurants, \
retailers and similar) for OriginCrest and judge how well each one fits OriginCrest's current range of \
Indonesian coffees (given in each request). {_GREEN_OR_ROASTED}

Use web search and web fetch on the business's own website and online shop first, then \
press, awards and directories. Record your findings with the record_research tool.

First decide what they'd buy from OriginCrest (buys: Green, Roasted, Both or Unclear), then score fit from \
0 to 100:
- Match with OriginCrest's range: for green, single origins they could feature or blends the range could \
supply (e.g. robusta for espresso blends); for roasted, a coffee from the range they could serve or stock \
(up to 30)
- Volume and buying: they buy green in volumes OriginCrest can serve, or buy roasted coffee regularly - café \
and restaurant use, retail shelves, gifting, office supply (up to 25)
- Indonesian interest: Indonesian coffee now or before, a gap an Indonesian coffee would fill, or an \
Indonesian or Asian food focus (up to 25)
- Practicalities: scale, freedom to choose suppliers, location - Adelaide/SA is a plus (up to 20)
Importers and traders can be partners (buying or distributing OriginCrest lots) or competitors: score how \
likely a partnership is. A café or restaurant tied to one roaster (e.g. a supply-and-machine deal) or a \
roaster's own café is a weaker fit for roasted coffee. Be calibrated: 70+ needs clear evidence they buy \
coffee of a kind OriginCrest's range covers; under 40 for businesses that don't deal in coffee, closed \
businesses, or an instant-coffee focus.

best_fit_origins: the one to three coffees from OriginCrest's range that suit them best, named as the range \
names them.

personalisation_hooks are 2-3 specific facts from the business's own materials that an intro email could \
mention (a current single origin, their roasting or sourcing approach, their menu or shelf, an award), each \
with its source URL. \
Use business facts only - nothing personal about individuals.

{CONTACT_RULES}"""

RESEARCH_SCHEMA = obj({
    "company_name": {"type": "string"},
    "website": {"type": "string"},
    "city": {"type": "string"},
    "state": {"type": "string", "enum": STATES},
    "what_they_roast": {"type": "string",
                        "description": "What they roast / sell / serve / import: range, style, blends, brands"},
    "buys": {"type": "string", "enum": ["Green", "Roasted", "Both", "Unclear"],
             "description": "What they'd buy from OriginCrest"},
    "size_estimate": {"type": "string", "enum": ["Micro", "Small", "Medium", "Large", "Unknown"]},
    "size_evidence": {"type": "string"},
    "sourcing": {"type": "string",
                 "description": "How they buy coffee: direct trade, named importers, or the roaster(s) they use"},
    "current_indonesian_coffees": {"type": "string", "description": "Indonesian coffees they sell now, or empty"},
    "wholesale": {"type": "string", "enum": ["Yes", "No", "Unknown"]},
    "public_email": {"type": "string"},
    "public_email_source_url": {"type": "string"},
    "contact_name": {"type": "string"},
    "contact_role": {"type": "string"},
    "fit_score": {"type": "integer"},
    "best_fit_origins": {"type": "array", "items": {"type": "string"}},
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
You write short first-contact emails from OriginCrest to Australian businesses that buy coffee (roasters, \
importers, traders, cafés, restaurants, retailers and similar). Each email is written for one contact, from \
their row in OriginCrest's contacts spreadsheet and the web research on their business.

Pitch what the research says they'd buy (buys): green beans for roasters, importers and traders; roasted \
coffee for cafés, restaurants, retailers and other businesses; for "Both", lead with the one that fits them \
best and mention the other in a few words. Never offer roasted coffee to a roaster.

Tailor it to the contact: use their name and role if the row has them, and build on anything OriginCrest \
noted about them in the row (notes, talking points, how they met, category, services). Open with something \
specific about their business from the research hooks, connect it to one or two coffees from OriginCrest's \
current range that suit them (for an importer or trader, frame it as supply or partnership), and offer one easy next step. Ask \
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
- lead: OriginCrest emailed a business that could buy its green or roasted coffee (roaster, importer, café, \
retailer...) and hasn't heard back; pitch what the row says they'd buy ("Would buy"). Keep it lighter than \
the first email, add one new useful thing and give an easy out. If OriginCrest's current \
range (in the background below) has coffees the earlier email didn't mention, lead with the one or two that \
best suit them (e.g. a robusta for their espresso blend); otherwise offer samples.

Plain text, Australian English, 50-110 words; later follow-ups are shorter. Use only facts in the row \
details, the original email and the OriginCrest background. End with a short sign-off such as "Thanks," - \
the signature and any unsubscribe line are added automatically, so don't write them. If you are told \
the email continues an existing thread, write the body only; the subject is set for you, so return the \
subject you'd use anyway.

The row details are OriginCrest's own spreadsheet row. Notes and comments in it are private (e.g. "need to be \
called", "slow to reply"): use them for context, but never quote them or mention them in the email.

call_points: when you're told OriginCrest will phone this contact, give 2-4 short points for the call (each \
under 15 words): what was asked for, what to find out, any timing. Otherwise return an empty list."""

FOLLOWUP_SCHEMA = obj({
    "subject": {"type": "string"},
    "body": {"type": "string"},
    "call_points": {"type": "array", "items": {"type": "string"}},
})


def followup_prompt(*, record_type: str, number: int, row_context: dict[str, str], original: str,
                    days_since: int, threaded: bool, cfg, call: bool = False, has_email: bool = True) -> str:
    thread_note = "This continues the original email thread." if threaded else "There is no earlier thread to reply to."
    if call and has_email:
        thread_note += (" OriginCrest's notes say to phone this contact, so also give call_points; the email is "
                        "only sent if they still want to write after the call.")
    elif call:
        thread_note += (" There's no email address for them and OriginCrest will phone instead: give call_points, "
                        "and keep the subject and body to one line each (they aren't used).")
    return f"""Write follow-up #{number} ({record_type}). It's been {days_since} days since we last \
contacted them with no reply. {thread_note}

Row details:
{json.dumps(row_context, indent=2, ensure_ascii=False)}

Our last email to them:
{original or "(not found in Sent - rely on the row details)"}

OriginCrest background:
{offer_text(cfg)}"""
