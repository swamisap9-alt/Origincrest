"""Prompts and schemas for the supplier agent (Indonesian green-coffee suppliers).

Kept apart from prompts.py so the outreach agent's prompts never change because of supplier work.
"""

from __future__ import annotations

import json

from .llm import obj, strict_tool
from .prompts import CONTACT_RULES, _filled


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
that sell only through someone else's export licence. The Indonesian Ministry of Trade's exporter \
directory (InaExport) and exporter-association member lists (AEKI, GAEKI) are good sources: a company \
listed there counts as identified, and contact details published there count as published.

How to search: search in Indonesian as well as English - many exporters and cooperatives only publish in \
Indonesian (e.g. "eksportir kopi robusta Dampit", "koperasi kopi Temanggung ekspor", "green bean kopi \
arabika Preanger ekspor"). When the focus names a specific coffee or area, look for the processors, \
cooperatives and exporters in that exact area first, then exporters elsewhere who sell that coffee.

Evidence to look for and record: origins and growing regions, species, grades and processing methods; \
certifications (Organic, Fairtrade, Rainforest Alliance, 4C, SCA scores); export track record, loading \
port, minimum order, whether they ship part-container (LCL) as well as full containers (FCL), sample \
policy and incoterms; any shipments to Australia. Record only what a page says - "Not stated" beats a guess.

Comparison: OriginCrest already knows one vendor, described in the request. For each supplier, write two \
or three sentences comparing it with that vendor on the points the evidence supports (origins and range, \
certifications, export scale and minimums, how reachable and transparent they are, Australia experience), \
and say plainly what can't be compared because it isn't known. Never invent facts about either side.

Current demand: when the request gives OriginCrest's current demand (specific coffees for a customer), set \
covers_current_demand to "Yes" if the supplier's own pages show they sell those coffees (same area, \
species, grade and process, or a stated close match), "Partly" if they sell some of them, "No" if none, \
and "No demand given" when the request has no demand.

Fit score 0-100 for OriginCrest's needs:
- Sells the coffees in the current demand (or, failing that, the origins OriginCrest wants), specialty or \
export grade, with traceable lots (up to 30)
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
    "covers_current_demand": {"type": "string", "enum": ["Yes", "Partly", "No", "No demand given"]},
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
    obj({"suppliers": {"type": "array", "items": SUPPLIER_SCHEMA},
         "search_notes": {"type": "string", "description": (
             "Two or three sentences: what you searched (including Indonesian-language searches) and which "
             "candidates you left out and why. Most important when you record few or none.")}}),
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
3. The demand - depending on the row's "Covers current demand":
   - "Yes", "Partly", blank or missing: the current demand, stated concretely from the background (the \
customer, the LCL trial order and the FCL requirement that follows); for "Partly", only the coffees this \
supplier sells. A short list is fine here.
   - "No": this supplier doesn't sell the coffees in the current order, so don't ask them to supply it or \
list its coffees. Instead say OriginCrest imports for Australian roasters, starting with part-container \
(LCL) lots and moving to full 20 ft containers as demand grows, and that it is interested in their own \
coffees (from the row). Don't name the customer.
4. What OriginCrest asks for, as a short list (for "No": about their own coffees), and one easy next step \
(a reply, or a call on WhatsApp).

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


SUPPLIER_FOLLOWUP_SYSTEM = """\
You write brief, polite follow-up emails for OriginCrest, an Australian importer of Indonesian green \
coffee. OriginCrest, as the buyer, sent a sourcing enquiry to an Indonesian exporter, cooperative or mill \
and hasn't heard back.

Restate in one or two lines who OriginCrest is and the order it wants to place (the trial by \
part-container (LCL) and the full-container (FCL) step after it, from the background and the original \
email), repeat the one thing OriginCrest most needs (FOB price and minimum order, or samples), and offer a \
WhatsApp or video call as an alternative to email. If the row's "Covers current demand" is "No", restate \
OriginCrest's interest in the supplier's own coffees instead of the customer's order. Use short sentences and plain international English \
(many readers use English as a second language). Never mention other suppliers or vendors.

Plain text, 50-120 words; later follow-ups are shorter. Use only facts in the row details, the original \
email and the background. End with a short sign-off such as "Kind regards," - the signature is added \
automatically, so don't write it. If you are told the email continues an existing thread, write the body \
only; the subject is set for you, so return the subject you'd use anyway.

The row details are OriginCrest's own spreadsheet row. Notes in it are private: use them for context, but \
never quote them or mention them in the email.

call_points: when you're told OriginCrest will phone this supplier, give 2-4 short points for the call (each \
under 15 words). Otherwise return an empty list."""


def supplier_followup_prompt(*, number: int, row_context: dict[str, str], original: str, days_since: int,
                             threaded: bool, cfg, call: bool = False, has_email: bool = True) -> str:
    thread_note = "This continues the original email thread." if threaded else "There is no earlier thread to reply to."
    if call and has_email:
        thread_note += (" OriginCrest's notes say to phone this supplier, so also give call_points; the email is "
                        "only sent if they still want to write after the call.")
    elif call:
        thread_note += (" There's no email address for them and OriginCrest will phone instead: give call_points, "
                        "and keep the subject and body to one line each (they aren't used).")
    return f"""Write follow-up #{number} (supplier). It's been {days_since} days since we last \
contacted them with no reply. {thread_note}

Row details:
{json.dumps(row_context, indent=2, ensure_ascii=False)}

Our last email to them:
{original or "(not found in Sent - rely on the row details)"}

OriginCrest as a buyer:
{sourcing_text(cfg)}"""
