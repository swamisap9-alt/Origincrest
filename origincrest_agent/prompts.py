"""Prompts and output schemas for the three workflows."""

from __future__ import annotations

import json

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


def offer_text(cfg) -> str:
    offer = cfg["offer"]
    lines = [f"Business: {cfg['sender']['business_name']}", f"Summary: {offer['summary']}"]
    if offer.get("origins"):
        lines.append("Origins offered: " + "; ".join(offer["origins"]))
    for fact in offer.get("facts") or []:
        lines.append(f"Fact: {fact}")
    if offer.get("call_to_action"):
        lines.append(f"Preferred next step: {offer['call_to_action']}")
    return "\n".join(lines)


# --- Workflow 1: discovery -------------------------------------------------

DISCOVERY_SYSTEM = f"""\
You find Australian businesses for OriginCrest, an importer of Indonesian single-origin green coffee \
(Sumatra, Sulawesi, Java, Flores). You search the web for specialty roasters (potential buyers) and for \
freight forwarders, customs brokers and 3PLs that can move and store coffee, then record what you found \
with the record_companies tool.

{CONTACT_RULES}

Only include businesses that are in Australia and appear to be operating now. Skip anything in the \
"already known" list, including obvious variants of the same name."""

COMPANY_SCHEMA = obj({
    "company_name": {"type": "string"},
    "category": {"type": "string", "enum": [
        "Specialty roaster", "Freight forwarder", "Customs broker", "3PL / warehouse",
        "Freight forwarder + 3PL", "Other logistics"]},
    "website": {"type": "string", "description": "Company's own website URL, or empty"},
    "contact_name": {"type": "string", "description": "Named business contact if published, else empty"},
    "contact_role": {"type": "string"},
    "email": {"type": "string", "description": "Published business email, else empty"},
    "phone": {"type": "string", "description": "Published business phone, else empty"},
    "city": {"type": "string", "description": "City or suburb of the relevant office"},
    "state": {"type": "string", "enum": STATES},
    "adelaide_presence": {"type": "boolean", "description": "Based in Adelaide/SA or has an Adelaide office"},
    "services": {"type": "string", "description": "One or two sentences: what they roast / which services"},
    "coffee_evidence": {"type": "string", "description": "Evidence they are specialty / handle coffee"},
    "source_url": {"type": "string", "description": "Page where the contact details are published"},
    "notes": {"type": "string"},
})

RECORD_COMPANIES = strict_tool(
    "record_companies",
    "Record the businesses you found. Call once, after you've finished searching.",
    obj({"companies": {"type": "array", "items": COMPANY_SCHEMA}}),
)


def discovery_prompt(*, kind: str, region: str, focus: list[str], n: int, known: list[str]) -> str:
    if region == "adelaide":
        where = "based in Adelaide or elsewhere in South Australia, or with an Adelaide office"
    elif kind == "roaster":
        where = "anywhere in Australia outside South Australia"
    else:
        where = "anywhere in Australia (national networks and operators in every state)"
    focus_text = "; ".join(focus)
    known_text = ", ".join(known) if known else "(none yet)"
    if kind == "roaster":
        task = f"""\
Find up to {n} specialty coffee roasters {where}.
Focus this week: {focus_text}.
Good prospects roast specialty-grade coffee, sell single origins (filter and/or espresso) and buy green \
coffee they could source from Indonesia. Wholesale roasters that supply cafés are especially relevant. \
Café chains that don't roast, machine sellers and instant-coffee brands don't count."""
    else:
        task = f"""\
Find up to {n} freight forwarders, customs brokers and 3PL / warehouse operators {where} that handle \
green coffee imports and/or roasted coffee storage and distribution.
Focus this week: {focus_text}.
Look for evidence they handle coffee or similar food and agricultural imports: FCL/LCL sea freight from \
Indonesia (Belawan/Medan, Surabaya, Makassar, Jakarta), customs clearance and DAFF biosecurity for green \
coffee, fumigation, container unpacking, food-grade or bonded warehousing, pallet storage, pick-pack \
fulfilment and interstate distribution. Put that evidence in coffee_evidence."""
    return f"""{task}

Already known (skip these): {known_text}

When you've finished searching, call record_companies once with everything that qualifies."""


# --- Workflow 2: lead research and intro email ----------------------------

RESEARCH_SYSTEM = f"""\
You research Australian coffee roasters for OriginCrest and judge how well each one fits OriginCrest's \
Indonesian single-origin green coffees. Use web search and web fetch on the roaster's own website and \
online shop first, then press, awards and directories. Record your findings with the record_research tool.

Score fit from 0 to 100:
- Specialty focus and single-origin range (up to 30)
- How they buy green coffee and whether the volume suits an importer (up to 25)
- Indonesian interest: they sell Indonesian coffees now, have before, or have a gap an Indonesian lot \
would fill (up to 25)
- Practicalities: wholesale/café supply, roastery scale, location - Adelaide/SA is a plus (up to 20)
Be calibrated: 70+ needs clear evidence of specialty single-origin buying; under 40 for non-roasters, \
closed businesses, or commodity/instant focus.

personalisation_hooks are 2-3 specific facts from the roaster's own materials that an intro email could \
mention (a current single origin, their roasting approach, an award), each with its source URL. Use \
business facts only - nothing personal about individuals.

{CONTACT_RULES}"""

RESEARCH_SCHEMA = obj({
    "company_name": {"type": "string"},
    "website": {"type": "string"},
    "city": {"type": "string"},
    "state": {"type": "string", "enum": STATES},
    "what_they_roast": {"type": "string", "description": "Range, roast style, filter/espresso, blends"},
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
    "red_flags": {"type": "string", "description": "Closed, not a roaster, no wholesale, etc. Empty if none"},
})

RECORD_RESEARCH = strict_tool(
    "record_research",
    "Record what you found about the roaster. Call once, after researching.",
    RESEARCH_SCHEMA,
)


def research_prompt(row_context: dict[str, str], cfg) -> str:
    return f"""Research this roaster from our contacts sheet and score its fit.

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
You write short first-contact emails from OriginCrest to Australian specialty coffee roasters.

Write like a person who knows coffee, not a marketing department: Australian English, plain text, \
110-170 words, no hype, no emojis, no bullet lists. Open with something specific about their coffee \
taken from the research hooks, connect it to one or two OriginCrest origins that suit their range, and \
offer one easy next step. Ask at most one question.

Use only facts given to you: the research for the roaster, and the OriginCrest profile for OriginCrest. \
Don't state prices, certifications, volumes or availability unless they're in the profile. End the body \
with a short sign-off such as "Kind regards," - the signature and unsubscribe line are added \
automatically, so don't write them. Subject: under 8 words, specific, not clickbait."""


def intro_prompt(research: dict, contact_name: str, cfg) -> str:
    greeting = f"Address it to {contact_name.split()[0]}." if contact_name else "Open with 'Hi there,' or the team's name."
    return f"""Write the intro email. {greeting}

Research on the roaster:
{json.dumps(research, indent=2, ensure_ascii=False)}

OriginCrest profile:
{offer_text(cfg)}"""


# --- Workflow 3: follow-ups -------------------------------------------------

FOLLOWUP_SYSTEM = """\
You write brief, polite follow-up emails for OriginCrest, an importer of Indonesian green coffee.

Two situations:
- quote: OriginCrest asked a freight forwarder, customs broker or 3PL for a quote and hasn't heard back. \
Briefly restate what was requested (from the row details and the original email), ask whether they can \
quote or need anything else, and mention timing if the details give it.
- lead: OriginCrest emailed a roaster and hasn't heard back. Keep it lighter than the first email, add \
one new useful thing (e.g. offer to send samples, mention an origin that suits them) and give an easy out.

Plain text, Australian English, 50-110 words; later follow-ups are shorter. Use only facts in the row \
details, the original email and the OriginCrest profile. End with a short sign-off such as "Thanks," - \
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

OriginCrest profile:
{offer_text(cfg)}"""
