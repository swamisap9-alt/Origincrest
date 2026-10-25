import openpyxl

from origincrest_agent.discovery import run_discover
from origincrest_agent.llm import ResearchResult
from origincrest_agent.runtime import ChangeSet, commit


def company(name, category="Specialty roaster", **kw):
    base = {"company_name": name, "category": category, "website": "", "contact_name": "", "contact_role": "",
            "email": "", "phone": "", "city": "", "state": "", "adelaide_presence": False, "services": "",
            "coffee_evidence": "", "source_url": "", "notes": ""}
    base.update(kw)
    return base


PAGES = {
    "https://hillsroasting.com.au/contact": "<p>Wholesale enquiries: <a href='mailto:info@hillsroasting.com.au'>"
                                            "email us</a> or call (08) 8388 1234</p>",
    "https://quietroasters.com.au/contact": "<p>hello@quietroasters.com.au - no unsolicited emails please</p>",
}


def roaster_results():
    return ResearchResult(data={"companies": [
        company("Bean There Roasters Pty Ltd", website="beanthere.com.au", state="VIC"),  # in master already
        company("Hills Roasting Co", website="hillsroasting.com.au", email="info@hillsroasting.com.au",
                phone="08 8388 1234", city="Hahndorf", state="SA", adelaide_presence=True,
                services="Small-batch single origins", source_url="https://hillsroasting.com.au/contact"),
        company("Guessy Roasters", website="guessy.com.au", email="john@guessy.com.au", city="Melbourne",
                state="VIC", source_url="https://guessy.com.au/about"),
        company("Quiet Roasters", website="quietroasters.com.au", email="hello@quietroasters.com.au",
                city="Adelaide", state="SA", source_url="https://quietroasters.com.au/contact"),
    ]}, pages=[("https://hillsroasting.com.au/contact", PAGES["https://hillsroasting.com.au/contact"])])


def freight_results():
    return ResearchResult(data={"companies": [
        company("Indo Freight Lines", category="Freight forwarder", city="Port Adelaide", state="SA",
                adelaide_presence=True, coffee_evidence="Green coffee FCL from Belawan"),
        company("Sydney Bonded Warehousing", category="3PL / warehouse", city="Botany", state="NSW"),
    ]})


def rows_of(path, tab):
    ws = openpyxl.load_workbook(path)[tab]
    headers = [c.value for c in ws[1]]
    return [dict(zip(headers, (c.value for c in r))) for r in ws.iter_rows(min_row=2)]


def test_discovery_adds_rows_to_both_workbooks(cfg, agent, make_ctx):
    agent.research_queue = [roaster_results(), freight_results()]
    ctx = make_ctx("discover", pages=PAGES)
    added = run_discover(ctx, segments=["roasters_adelaide", "freight_australia"])
    assert added == 5

    for name in ("rfq", "master"):
        path = cfg.workbook_path(name)
        adelaide = rows_of(path, "Roasters - Adelaide")
        assert [r["Company"] for r in adelaide] == ["Hills Roasting Co", "Quiet Roasters"]
        hills = adelaide[0]
        assert hills["Status"] == "New - to review"
        assert hills["Email"] == "info@hillsroasting.com.au"  # verified on the source page
        assert hills["Phone"] == "08 8388 1234"
        assert hills["Source URL"] == "https://hillsroasting.com.au/contact"
        assert hills["Added by"] == "OriginCrest Agent"
        quiet = adelaide[1]
        assert quiet["Email"] is None and "no unsolicited" in quiet["Agent notes"]
        aus = rows_of(path, "Roasters - Australia")
        assert aus[0]["Company"] == "Guessy Roasters"
        assert aus[0]["Email"] is None and "Unverified email" in aus[0]["Agent notes"]
        assert [r["Company"] for r in rows_of(path, "Freight-3PL - Adelaide")] == ["Indo Freight Lines"]
        assert [r["Company"] for r in rows_of(path, "Freight-3PL - Australia")] == ["Sydney Bonded Warehousing"]
        assert rows_of(path, "Freight-3PL - Adelaide")[0]["Record type"] == "Quote"

    assert "Bean There" in " ".join(ctx.report.sections["Skipped duplicates"])

    # Second run with the same results: everything is a duplicate now.
    agent.research_queue = [roaster_results(), freight_results()]
    assert run_discover(make_ctx("discover", pages=PAGES), segments=["roasters_adelaide", "freight_australia"]) == 0


def test_discovery_caps(cfg, agent, make_ctx):
    cfg.data["discovery"]["max_new_rows_per_segment"] = 1
    agent.research_queue = [roaster_results()]
    ctx = make_ctx("discover", pages=PAGES)
    assert run_discover(ctx, segments=["roasters_adelaide"]) == 1
    assert any("cap of 1 reached" in line for line in ctx.report.sections["Caps"])


def test_rows_wait_while_workbook_is_open_in_excel(cfg, agent, make_ctx):
    rfq = cfg.workbook_path("rfq")
    owner = rfq.parent / ("~$" + rfq.name)
    owner.write_text("Excel")
    agent.research_queue = [roaster_results()]
    ctx = make_ctx("discover", pages=PAGES)
    run_discover(ctx, segments=["roasters_adelaide"])
    assert "Roasters - Adelaide" not in openpyxl.load_workbook(rfq).sheetnames
    assert "Roasters - Adelaide" in openpyxl.load_workbook(cfg.workbook_path("master")).sheetnames
    assert any("open in Excel" in line for line in ctx.report.sections["Not written"])

    owner.unlink()
    commit(make_ctx("flush"), ChangeSet())
    assert len(rows_of(rfq, "Roasters - Adelaide")) == 2
    commit(make_ctx("flush"), ChangeSet())  # nothing left queued: no duplicate rows
    assert len(rows_of(rfq, "Roasters - Adelaide")) == 2


def test_dry_run_writes_nothing(cfg, agent, make_ctx):
    agent.research_queue = [roaster_results()]
    ctx = make_ctx("discover", pages=PAGES, dry_run=True)
    run_discover(ctx, segments=["roasters_adelaide"])
    assert "Would add" in ctx.report.sections
    assert "Roasters - Adelaide" not in openpyxl.load_workbook(cfg.workbook_path("rfq")).sheetnames
