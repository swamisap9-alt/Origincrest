from datetime import date, timedelta

import openpyxl

from origincrest_agent.cli import main
from origincrest_agent.followup import run_followup
from origincrest_agent.leads import run_leads
from origincrest_agent.llm import ResearchResult
from origincrest_agent.suppliers import FOCUS, run_focus, run_suppliers

from .conftest import TODAY, mail, utc

TAB = "New Suppliers - Indonesia"
SITE = "https://gayomountain.co.id"


def supplier(name="Koperasi Gayo Mountain", email="export@gayomountain.co.id", score=82, **kw):
    data = {"company_name": name, "business_type": "Cooperative (exports directly)", "website": SITE,
            "contact_name": "", "contact_role": "", "email": email, "phone": "+62 643 555 0101",
            "whatsapp": "+62 812 5550 0102", "city": "Takengon", "province": "Aceh",
            "origins": "Gayo arabica G1, wet-hulled and washed", "certifications": "Organic, Fairtrade",
            "export_capability": "Exports FCL from Belawan; LCL on request; samples by courier",
            "comparison_with_current_vendor": "Certified and publishes export contacts; prices not public.",
            "fit_score": score, "fit_rationale": "Certified Gayo lots with an export licence.",
            "suggested_approach": "Email the export desk, then WhatsApp the published number.",
            "source_url": SITE + "/contact", "notes": ""}
    data.update(kw)
    return data


PAGES = {SITE + "/contact": "Export desk: export@gayomountain.co.id Tel +62 643 555 0101 WhatsApp +62 812 5550 0102"}


def sheet(cfg):
    ws = openpyxl.load_workbook(cfg.workbook_path("suppliers"))[TAB]
    headers = [c.value for c in ws[1]]
    return {r[0].value: dict(zip(headers, (c.value for c in r))) for r in ws.iter_rows(min_row=2)}


def set_cfg(cfg, tmp_path, **sourcing):
    cfg.data["workbooks"]["suppliers"]["path"] = str(tmp_path / "suppliers" / "suppliers.xlsx")
    cfg.data["suppliers"]["focus_per_run"] = 1
    cfg.data["sourcing"].update(sourcing)


def test_focus_rotates_every_run():
    from origincrest_agent.config import DEFAULTS, Config

    cfg = Config(DEFAULTS)
    seen = [tuple(run_focus(cfg, date(2026, 11, 12) + timedelta(days=3 * i))) for i in range(len(FOCUS) // 2)]
    assert len(set(seen)) == len(seen)
    assert {f for run in seen for f in run} == set(FOCUS)
    # the same run day always gives the same focus
    assert run_focus(cfg, date(2026, 11, 12)) == run_focus(cfg, date(2026, 11, 12))


def test_discovers_compares_and_drafts_enquiry(cfg, tmp_path, agent, mailbox, make_ctx):
    set_cfg(cfg, tmp_path, demand="Soho Roasters: LCL trial of 600 kg, then one 20ft FCL per quarter")
    agent.research_queue = [ResearchResult(data={"suppliers": [
        supplier(),
        supplier("PT Putra Pulau", "sales@putrapulau.example", website="", source_url=""),   # the current vendor
        supplier("CV Unverified Kopi", "made-up@nowhere.example", website="https://unverified.example",
                 source_url="https://unverified.example/contact", whatsapp="", phone=""),
    ]})]
    agent.structured_queue = [{"subject": "Green coffee import enquiry from Australia - Gayo arabica",
                               "body": "Dear Koperasi Gayo Mountain team,\n\nWe are ready to import.\n\nKind regards,"}]
    ctx = make_ctx("suppliers", pages=PAGES)
    assert run_suppliers(ctx) == 2
    assert run_leads(ctx, only_kind="supplier") == 1

    search_prompt = agent.calls[0][1]
    assert "PT Putra Pulau" in search_prompt and "Soho Roasters" in search_prompt
    rows = sheet(cfg)
    assert set(rows) == {"Koperasi Gayo Mountain", "CV Unverified Kopi"}
    assert any("PT Putra Pulau" in line for line in ctx.report.sections["Skipped duplicates"])
    row = rows["Koperasi Gayo Mountain"]
    assert row["Email"] == "export@gayomountain.co.id" and row["State"] == "Aceh"
    assert row["Phone"] == "+62 643 555 0101 / WhatsApp +62 812 5550 0102"
    assert row["Comparison with current vendor"].startswith("Certified")
    assert row["Suggested approach"] and row["Fit score (0-100)"] == 82
    assert row["Status"] == "Enquiry drafted - to review" and row["Record type"] == "Supplier"
    # An address that isn't on the supplier's own pages is never used.
    assert rows["CV Unverified Kopi"]["Email"] is None
    assert rows["CV Unverified Kopi"]["Status"] == "No public email"

    [msg] = mailbox.drafts
    assert msg["To"] == "export@gayomountain.co.id"
    assert msg["List-Unsubscribe"] is None and "unsubscribe" not in msg.get_content()
    assert "Swami K" in msg.get_content()
    # The email writer sees the demand and the logistics line, but not your private comparison.
    email_prompt = next(p for kind, p in agent.calls if kind == "structured")
    assert "Soho Roasters: LCL trial of 600 kg" in email_prompt and "already arranged" in email_prompt
    assert "Certified and publishes" not in email_prompt and "Email the export desk" not in email_prompt

    # The next run finds nothing new and drafts nothing twice.
    agent.research_queue = [ResearchResult(data={"suppliers": [supplier()]})]
    ctx = make_ctx("suppliers", pages=PAGES)
    assert run_suppliers(ctx) == 0
    assert run_leads(ctx, only_kind="supplier") == 0
    assert len(mailbox.drafts) == 1


def test_blank_demand_asks_for_placeholders_and_low_fit_is_not_drafted(cfg, tmp_path, agent, mailbox, make_ctx):
    set_cfg(cfg, tmp_path)
    agent.research_queue = [ResearchResult(data={"suppliers": [supplier(score=35)]})]
    ctx = make_ctx("suppliers", pages=PAGES)
    assert run_suppliers(ctx) == 1
    assert "sourcing.demand is blank" in ctx.report.sections["Demand"][0]
    assert run_leads(ctx, only_kind="supplier") == 0
    assert sheet(cfg)["Koperasi Gayo Mountain"]["Status"] == "Researched - low fit"
    assert not mailbox.drafts


def test_regular_leads_run_ignores_suppliers_until_the_workbook_exists(cfg, tmp_path, make_ctx):
    set_cfg(cfg, tmp_path)
    ctx = make_ctx("leads")
    run_leads(ctx)
    assert "Errors" not in ctx.report.sections


def test_followup_after_three_days(cfg, tmp_path, agent, mailbox, make_ctx):
    set_cfg(cfg, tmp_path, demand="Soho Roasters: LCL trial of 600 kg, then one 20ft FCL per quarter")
    cfg.data["followup"]["workbooks"] = []
    agent.research_queue = [ResearchResult(data={"suppliers": [supplier()]})]
    ctx = make_ctx("suppliers", pages=PAGES)
    run_suppliers(ctx)
    run_leads(ctx, only_kind="supplier")
    draft = mailbox.drafts.pop()
    # You send it the same day.
    mailbox.sent.append(mail("Sent", utc(TODAY, 10), "hello@origincrest.com.au", "export@gayomountain.co.id",
                             draft["Subject"], "Dear team,\n\nWe are ready to import.", "<first@origincrest>"))

    def followup_on(day):
        c = make_ctx("followup", pages=PAGES)
        c.today = day
        return run_followup(c)

    assert followup_on(TODAY + timedelta(days=2)) == 0
    assert sheet(cfg)["Koperasi Gayo Mountain"]["Status"] == "Open"
    agent.structured_queue = [{"subject": "x", "body": "Dear team,\n\nFollowing up.\n\nKind regards,",
                               "call_points": []}]
    assert followup_on(TODAY + timedelta(days=3)) == 1
    [msg] = mailbox.drafts
    assert msg["Subject"].startswith("Re: ") and msg["In-Reply-To"] == "<first@origincrest>"
    prompt = agent.calls[-1][1]
    assert "(supplier)" in prompt and "Soho Roasters" in prompt and "Certified and publishes" not in prompt
    row = sheet(cfg)["Koperasi Gayo Mountain"]
    assert row["Follow-ups drafted"] == 1
    assert row["Next action date"].date() == TODAY + timedelta(days=6)


def test_dry_run_writes_nothing(cfg, tmp_path, agent, mailbox, make_ctx):
    set_cfg(cfg, tmp_path)
    agent.research_queue = [ResearchResult(data={"suppliers": [supplier()]})]
    ctx = make_ctx("suppliers", pages=PAGES, dry_run=True)
    assert run_suppliers(ctx) == 1
    assert not cfg.workbook_path("suppliers").exists() and not mailbox.drafts
    assert "Koperasi Gayo Mountain" in ctx.report.sections["Would add"][0]


def test_cli_has_suppliers_command(capsys):
    try:
        main(["suppliers", "--help"])
    except SystemExit as exc:
        assert exc.code == 0
    assert "--focus" in capsys.readouterr().out
