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


# --- Supplier discovery: Indonesian exporters, compared with the current vendor ----------------

SUPPLIER_CONTACT_RULES = CONTACT_RULES + """
- WhatsApp numbers count as phone numbers: record one only if the business publishes it for trade enquiries."""

SUPPLIER_SYSTEM = f"""\
You find Indonesian green-coffee suppliers for OriginCrest, an Australian importer that buys Indonesian \
single-origin green coffee and sells it to Australian roasters. OriginCrest is the buyer here. Search the \
web, then record what you found with the record_suppliers tool.

Who qualifies: businesses in Indonesia that can export green (unroasted) coffee themselves - licensed \
exporters, producer cooperatives with their own export licence, mills and processors that export, and \
established trading houses. They must appear to be operating now. Leave out roasted-coffee brands and \
cafes, brokers and importers based outside Indonesia, marketplace listings with no identifiable company \
(Alibaba, IndoTrading and similar pages are leads to a company, never the source for one), and farms \
that sell only through someone else's export licence.

Evidence to look for and record: origins and growing regions, species, grades and processing methods; \
certifications (Organic, Fairtrade, Rainforest Alliance, 4C, SCA scores); export track record, loading \
port, minimum order, whether they ship part-container (LCL) as well as full containers (FCL), sample \
policy and incoterms; any shipments to Australia. Record only what a page says - "Not stated" beats a guess.

Comparison: OriginCrest already knows one vendor, described in the request. For each supplier, write two \
or three sentences comparing it with that vendor on the points the evidence supports (origins and range, \
certifications, export scale and minimums, how reachable and transparent they are, Australia experience), \
and say plainly what can't be compared because it isn't known. Never invent facts about either side.

Fit score 0-100 for OriginCrest's needs:
- Sells the origins OriginCrest wants, specialty grade, with traceable lots (up to 30)
- Proven export capability: licence, documents, shipments abroad, ideally to Australia (up to 30)
- Can start small (samples, part-container) and scale to full containers (up to 20)
- Reachable and transparent: named contacts, published business email or WhatsApp, an active site (up to 20)

suggested_approach: one or two sentences on the best way to open the conversation with this supplier - \
which published channel to use first (email, WhatsApp, website enquiry form, LinkedIn company page), \
who to address, and what to lead with. Published business channels only.

{{SUPPLIER_CONTACT_RULES}}

Skip anything in the "already known" list, including obvious variants of the same name.""".replace(
    "{{SUPPLIER_CONTACT_RULES}}", SUPPLIER_CONTACT_RULES)

SUPPLIER_SCHEMA = obj({
    "company_name": {"type": "string"},
    "business_type": {"type": "string", "description": "e.g. 'Exporter', 'Cooperative (exports directly)', "
                                                       "'Mill and exporter', 'Trading house'"},
    "website": {"type": "string", "description": "Company's own website URL, or empty"},
    "contact_name": {"type": "string", "description": "Named business contact if published, else empty"},
    "contact_role": {"type": "string"},
    "email": {"type": "string", "description": "Published business email, else empty"},
    "phone": {"type": "string", "description": "Published business phone with country code, else empty"},
    "whatsapp": {"type": "string", "description": "Published business WhatsApp number, else empty"},
    "city": {"type": "string"},
    "province": {"type": "string", "description": "e.g. Aceh, North Sumatra, West Java"},
    "origins": {"type": "string", "description": "Origins, regions, species, grades and processes they sell"},
    "certifications": {"type": "string", "description": "Certifications stated on their pages, or 'Not stated'"},
    "export_capability": {"type": "string", "description": "Export track record, port, minimum order, LCL/FCL, "
                                                           "samples, incoterms, Australia experience"},
    "comparison_with_current_vendor": {"type": "string"},
    "fit_score": {"type": "integer"},
    "fit_rationale": {"type": "string", "description": "One or two sentences"},
    "suggested_approach": {"type": "string"},
    "source_url": {"type": "string", "description": "Page where the contact details are published"},
    "notes": {"type": "string"},
})

RECORD_SUPPLIERS = strict_tool(
    "record_suppliers",
    "Record the suppliers you found. Call once, after you've finished searching.",
    obj({"suppliers": {"type": "array", "items": SUPPLIER_SCHEMA}}),
)


def sourcing_text(cfg, for_search: bool = False) -> str:
    """OriginCrest as a buyer: who it is, the demand behind the enquiry, and what's already arranged."""
    s = cfg["sourcing"]
    lines = [f"Business: {cfg['sender']['business_name']}",
             "About: " + (_filled(s.get("intro")) or "An Australian importer of Indonesian single-origin green "
                                                    "coffee, supplying Australian roasters.")]
    wanted = [w for w in (_filled(x) for x in s.get("wanted") or []) if w]
    if wanted:
        lines.append("Origins wanted: " + ", ".join(wanted))
    demand = _filled(s.get("demand"))
    if demand:
        lines.append(f"Current demand: {demand}")
    elif not for_search:
        lines.append("Current demand: (not written down - use clearly marked [placeholders] for the customer, "
                     "the LCL trial quantity and the FCL quantity and frequency)")
    if _filled(s.get("logistics")):
        lines.append(f"Logistics: {_filled(s['logistics'])}")
    if not for_search:
        asks = [a for a in (_filled(x) for x in s.get("ask") or []) if a]
        if asks:
            lines.append("What to ask the supplier for: " + "; ".join(asks))
    return "\n".join(lines)


def baseline_text(cfg) -> str:
    vendor = cfg["suppliers"]["baseline_vendor"]
    name = _filled(vendor.get("name")) or "the current vendor"
    profile = _filled(vendor.get("profile"))
    return f"{name}: " + (profile or "no details on file. Compare on what the supplier's own evidence shows and "
                                     "list what OriginCrest should ask both sides to compare properly.")


def supplier_discovery_prompt(*, focus: list[str], n: int, known: list[str], cfg) -> str:
    return "\n\n".join([
        f"Find up to {n} Indonesian green-coffee suppliers that can export to Australia.",
        f"Focus for this search: {'; '.join(focus)}.",
        "OriginCrest as a buyer:\n" + sourcing_text(cfg, for_search=True),
        "Current vendor to compare each supplier with:\n" + baseline_text(cfg),
        f"Already known (skip these): {', '.join(known) if known else '(none yet)'}",
        "When you've finished searching, call record_suppliers once with everything that qualifies.",
    ])


SUPPLIER_INTRO_SYSTEM = """\
You write first-contact sourcing enquiries from OriginCrest, an Australian importer of Indonesian green \
coffee, to Indonesian exporters, cooperatives and mills. OriginCrest is the buyer. Each email is written \
for one supplier, from its row in OriginCrest's supplier spreadsheet.

Structure, in this order:
1. One or two sentences introducing OriginCrest and why this supplier (name a specific origin, process or \
certification of theirs from the row).
2. Readiness: OriginCrest is ready to import into Australia now, and the logistics side (as given in the \
background) is already arranged, so the supplier only needs to deliver to the agreed incoterm.
3. The demand, stated concretely from the background: the customer, the LCL trial order and the FCL \
requirement that follows. A short list is fine here.
4. What OriginCrest asks for, as a short list, and one easy next step (a reply, or a call on WhatsApp).

Use the demand exactly as given. If a number, customer or date isn't given, write a clearly marked \
placeholder in square brackets for OriginCrest to fill in before sending, e.g. [LCL trial quantity in kg]. \
Never invent volumes, prices, dates or certifications. Never mention any other supplier or vendor, or that \
OriginCrest is comparing suppliers.

Plain text, clear international business English (many readers use English as a second language): short \
sentences, no idioms, no hype, 170-300 words (more only if the demand needs it). Use the contact's name if the row has one, otherwise \
"Dear <company> team,". End the body with a short sign-off such as "Kind regards," - the signature is added \
automatically, so don't write it. Subject: under 10 words and specific, e.g. "Green coffee import enquiry \
from Australia - Gayo arabica"."""


def supplier_intro_prompt(row_context: dict[str, str], contact_name: str, cfg) -> str:
    greeting = (f"Address it to {contact_name}." if contact_name
                else "No contact name is known: address the company's team.")
    return f"""Write the sourcing enquiry for this supplier. {greeting}

Their row in OriginCrest's supplier spreadsheet:
{json.dumps(row_context, indent=2, ensure_ascii=False)}

OriginCrest as a buyer:
{sourcing_text(cfg)}"""


# --- Workflow 3: follow-ups -------------------------------------------------

FOLLOWUP_SYSTEM = """\
You write brief, polite follow-up emails for OriginCrest, an importer of Indonesian green coffee.

Three situations:
- quote: OriginCrest asked a freight forwarder, customs broker or 3PL for a quote and hasn't heard back. \
Briefly restate what was requested (from the row details and the original email), ask whether they can \
quote or need anything else, and mention timing if the details give it.
- lead: OriginCrest emailed a coffee business (roaster, importer, wholesaler, trader) and hasn't heard back. Keep it lighter than the first email, add \
one new useful thing (e.g. offer to send samples, mention an origin that suits them) and give an easy out.
- supplier: OriginCrest, as the buyer, sent a sourcing enquiry to an Indonesian exporter, cooperative or mill \
and hasn't heard back. Restate in one or two lines who OriginCrest is and the order it wants to place (the \
LCL trial and the FCL volume, from the original email), repeat the one thing you most need (offer list and \
minimum order, or samples), and offer a WhatsApp or video call as an alternative to email. Use short \
sentences and plain international English. Never mention other suppliers.

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
{sourcing_text(cfg) if record_type == "supplier" else offer_text(cfg)}"""
