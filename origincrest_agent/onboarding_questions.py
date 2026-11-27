"""OriginCrest's supplier onboarding questionnaire for Indonesian green-coffee exporters.

Each question has a stable key (used to match answers across runs), a section and the text sent to
suppliers. Edit the wording freely; keep the keys unchanged so earlier answers stay matched.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Question:
    key: str
    section: str
    text: str


SECTIONS = {
    "A": "Company and legal",
    "B": "Coffees and quality",
    "C": "Commercial terms",
    "D": "Packaging and logistics",
    "E": "Export documents and Australian biosecurity",
    "F": "Compliance and references",
}

QUESTIONS: tuple[Question, ...] = (
    # A - Company and legal
    Question("A1", "A", "Registered company name (PT / CV / cooperative) and any trading or brand names"),
    Question("A2", "A", "Business registration number (NIB) and tax number (NPWP)"),
    Question("A3", "A", "Coffee export registration or licence (and ICO exporter code if you have one)"),
    Question("A4", "A", "Year established, ownership, and number of staff"),
    Question("A5", "A", "Registered office address, and the address of your mill or warehouse"),
    Question("A6", "A", "Contacts for sales, export documents and quality: name, role, email, WhatsApp"),
    Question("A7", "A", "Bank name and account holder name (must be the company). We confirm full bank details "
                        "by phone or video call before any payment"),
    # B - Coffees and quality
    Question("B1", "B", "Coffees you offer: origin/region, species (Arabica/Robusta/Liberica), variety, "
                        "process (natural, washed, semi-washed, honey), grade, screen size, altitude"),
    Question("B2", "B", "Harvest calendar and current availability by coffee (tonnes)"),
    Question("B3", "B", "Quality specification: moisture, defect count / grading standard (e.g. SNI), "
                        "cupping scores (SCA) and who cups"),
    Question("B4", "B", "Quality control: sorting and grading, lab tests, pre-shipment samples, retained "
                        "reference samples"),
    Question("B5", "B", "Certifications (Organic, Fairtrade, Rainforest Alliance, 4C, Halal, HACCP, ISO): "
                        "certificate numbers and expiry dates"),
    Question("B6", "B", "Traceability: can you trace each lot to farm, cooperative or washing station?"),
    Question("B7", "B", "Grower relationships and sustainability: how you buy from farmers and pay them"),
    # C - Commercial terms
    Question("C1", "C", "Price list: FOB price per kg for each coffee, for a trial / LCL lot, about 2 tonnes, "
                        "and a 20 ft FCL; how long prices are valid"),
    Question("C2", "C", "Incoterms you offer (FOB, CFR, CIF) and your loading port(s)"),
    Question("C3", "C", "Minimum order for LCL and for FCL; can one LCL shipment mix several coffees?"),
    Question("C4", "C", "Payment terms: deposit %, balance timing, method (TT, L/C), currency"),
    Question("C5", "C", "Lead time from deposit to ready-to-ship, and shipping frequency to Australia"),
    Question("C6", "C", "Sample policy: green and roasted samples, sizes, cost, courier"),
    Question("C7", "C", "Quality claims: how you handle a lot that doesn't match the approved sample"),
    # D - Packaging and logistics
    Question("D1", "D", "Packaging: bag type and size (jute / plastic, 30 / 60 kg), liners (GrainPro or "
                        "inner plastic), labels and ICO marks"),
    Question("D2", "D", "Container loading: bags per 20 ft, fumigation, container lining or desiccants"),
    Question("D3", "D", "Shipping experience: LCL shipments, shipments to Australia (how many, last one), "
                        "preferred forwarder"),
    # E - Export documents and Australian biosecurity
    Question("E1", "E", "Documents you provide: commercial invoice, packing list, bill of lading, certificate "
                        "of origin, ICO certificate, phytosanitary certificate, fumigation certificate, PEB"),
    Question("E2", "E", "Are the export documents and charges included in your FOB price? If not, their cost"),
    Question("E3", "E", "Experience meeting Australia's biosecurity (BICON) conditions for green coffee: free "
                        "of soil, live insects, other seeds and contaminants"),
    Question("E4", "E", "Document timing: drafts before shipment, originals by courier or telex release"),
    # F - Compliance and references
    Question("F1", "F", "Food safety and storage: HACCP or similar, pest control, warehouse conditions"),
    Question("F2", "F", "Any legal disputes, export restrictions or company matters that could affect "
                        "shipments?"),
    Question("F3", "F", "Two overseas buyer references we may contact (ideally Australia or New Zealand)"),
)

KEYS = tuple(q.key for q in QUESTIONS)
BY_KEY = {q.key: q for q in QUESTIONS}
