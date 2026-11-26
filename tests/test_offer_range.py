import zipfile
from datetime import datetime, timedelta

import openpyxl
import requests

from origincrest_agent.cli import _load_offer_range
from origincrest_agent.discovery import run_discover
from origincrest_agent.followup import run_followup
from origincrest_agent.leads import run_leads
from origincrest_agent.offer_range import html_text, load_range
from origincrest_agent.prompts import RANGE_SUMMARY, RESEARCH_SCHEMA, discovery_prompt, offer_text, research_prompt

from .test_followup import sent_rfq
from .test_leads import seed_freight

RANGE_PAGE = """<html><head><title>x</title><script>var a=1;</script></head><body><nav>Home | Shop</nav>
<h1>Our green coffee range</h1><table><tr><th>Coffee</th><th>Species</th><th>Process</th></tr>
<tr><td>Aceh Gayo G1</td><td>Arabica</td><td>Wet-hulled</td></tr>
<tr><td>Kalibaru G1 16/18</td><td>Robusta</td><td>Natural</td></tr></table><footer>(c) OriginCrest</footer></body></html>"""
URL = "https://origincrest.example/range"


def catalogue_files(tmp_path):
    xlsx = tmp_path / "Green coffee price list.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Range"
    ws.append(["Origin", "Species", "Process", "Grade"])
    ws.append(["Toraja Sapan", "Arabica", "Full wash", "G1"])
    ws.append(["Argopuro", "Robusta", "Natural", "G1"])
    wb.save(xlsx)
    csv_file = tmp_path / "extra.csv"
    csv_file.write_text("Origin,Process\nBali Kintamani,Natural\n", encoding="utf-8")
    docx = tmp_path / "offer sheet.docx"
    with zipfile.ZipFile(docx, "w") as z:
        z.writestr("word/document.xml", '<w:document><w:body><w:p><w:r><w:t>Flores Bajawa - full wash'
                                        '</w:t></w:r></w:p></w:body></w:document>')
    return xlsx, csv_file, docx


def test_html_page_becomes_readable_lines():
    text = html_text(RANGE_PAGE)
    assert "Aceh Gayo G1 | Arabica | Wet-hulled" in text.splitlines()
    assert "var a" not in text and "Home | Shop" not in text and "(c)" not in text


def test_range_read_from_website_and_files(cfg, tmp_path):
    xlsx, csv_file, docx = catalogue_files(tmp_path)
    cfg.data["offer"]["range_sources"] = [URL, str(xlsx), str(csv_file), str(docx), str(tmp_path / "gone.xlsx")]
    current = load_range(cfg, fetch=lambda url: html_text(RANGE_PAGE))
    for coffee in ("Kalibaru G1 16/18 | Robusta", "Toraja Sapan | Arabica | Full wash | G1", "Bali Kintamani",
                   "Flores Bajawa - full wash"):
        assert coffee in current.text
    assert len(current.loaded) == 4 and "gone.xlsx" in current.problems[0]

    cfg.offer_range = current.text
    cfg.data["offer"]["summary"] = ""  # no summary of your own: the range speaks for itself
    text = offer_text(cfg)
    assert RANGE_SUMMARY in text and "Current range" in text and "Argopuro | Robusta" in text
    assert "Argopuro | Robusta" in research_prompt({"Company": "Bean Co"}, cfg)
    assert "enum" not in RESEARCH_SCHEMA["properties"]["best_fit_origins"]["items"]  # any coffee in the range
    short = offer_text(cfg, range_chars=60)
    assert "more in the full range" in short and "Bali" not in short


def test_without_a_range_the_old_summary_is_used(cfg):
    from origincrest_agent.prompts import DEFAULT_SUMMARY

    cfg.data["offer"]["summary"] = ""
    assert load_range(cfg).text == ""
    assert DEFAULT_SUMMARY in offer_text(cfg) and "Current range" not in offer_text(cfg)


def test_range_is_cached_and_last_good_copy_survives_an_outage(cfg):
    cfg.data["offer"]["range_sources"] = [URL]
    now = datetime(2026, 11, 26, 6, 0)
    first = load_range(cfg, fetch=lambda url: html_text(RANGE_PAGE), now=now)

    def down(url):
        raise requests.ConnectionError("site down")

    again = load_range(cfg, fetch=down, now=now + timedelta(hours=2))  # within 12 h: not fetched again
    assert again.text == first.text and not again.from_cache and again.problems == []
    later = load_range(cfg, fetch=down, now=now + timedelta(hours=13))
    assert later.text == first.text and later.from_cache and "site down" in later.problems[0]
    assert "last good copy" in later.describe()


def test_long_range_is_trimmed(cfg):
    cfg.data["offer"].update(range_sources=[URL], range_max_chars=80)
    current = load_range(cfg, fetch=lambda url: "\n".join(f"Lot {i} | Arabica | Natural" for i in range(50)))
    assert len(current.text) < 140 and "list trimmed" in current.text


def test_run_loads_the_range_for_emails(cfg, tmp_path, make_ctx):
    _, csv_file, _ = catalogue_files(tmp_path)
    cfg.data["offer"]["range_sources"] = [str(csv_file)]
    ctx = make_ctx("leads")
    _load_offer_range(ctx)
    assert "Bali Kintamani" in ctx.cfg.offer_range
    assert "1 source(s)" in ctx.report.sections["Your range"][0]


def test_buyer_discovery_prompt_describes_the_range(cfg):
    cfg.offer_range = "Kalibaru G1 | Robusta | Natural"
    prompt = discovery_prompt(kind="buyer", category="Roasters", region="other", focus=[], n=5, known=[],
                              examples=[], offer=offer_text(cfg, range_chars=4000))
    assert "Kalibaru G1 | Robusta" in prompt
    freight = discovery_prompt(kind="freight", category="3PL", region="other", focus=[], n=5, known=[],
                               examples=[], offer=offer_text(cfg))
    assert "Kalibaru" not in freight


def test_new_outreach_off_stops_freight_searches_and_rfqs_but_not_follow_ups(cfg, agent, mailbox, make_ctx):
    cfg.data["discovery"]["profiles"]["rfq"]["new_outreach"] = False
    assert not cfg.new_outreach("rfq") and cfg.new_outreach("master")

    ctx = make_ctx("discover")
    run_discover(ctx, categories=["(nothing this run)"])
    assert any("rfq: new outreach is off" in n for n in ctx.report.sections["Categories"])
    assert "New Freight-3PL - Adelaide" not in openpyxl.load_workbook(cfg.workbook_path("rfq")).sheetnames

    seed_freight(cfg)
    run_leads(make_ctx("leads"))
    assert not [m for m in mailbox.drafts if "indofreight" in m["To"]]  # no new RFQ

    sent_rfq(mailbox)  # an RFQ you sent earlier is still chased...
    assert run_followup(make_ctx("followup")) == 1
    assert "seaway" in mailbox.drafts[-1]["To"]


def test_follow_ups_off_leaves_the_freight_workbook_alone(cfg, agent, mailbox, make_ctx):
    cfg.data["discovery"]["profiles"]["rfq"].update(new_outreach=False, follow_ups=False)
    sent_rfq(mailbox)  # ...unless follow-ups are off for that workbook too
    before = cfg.workbook_path("rfq").read_bytes()
    ctx = make_ctx("followup")
    assert run_followup(ctx) == 0
    assert mailbox.drafts == []
    assert cfg.workbook_path("rfq").read_bytes() == before  # not even reply checks written
    assert ctx.report.sections["Follow-ups off"] == ["rfq: follow_ups: false - not checked"]
    assert cfg.follow_ups("master")
