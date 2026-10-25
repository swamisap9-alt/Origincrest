"""Check that contact details really are published on the business's web pages.

A detail is kept only if it appears on a page Claude fetched or on the source
page / the company's own contact page (fetched here). Pages that say the
business doesn't want unsolicited email are respected.
"""

from __future__ import annotations

import html
import re
from dataclasses import dataclass, field
from urllib.parse import urljoin

import requests

from .dedupe import company_domain, host_of, norm_email, norm_phone
from .utils import log

USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) OriginCrestContactCheck/1.0"

NO_UNSOLICITED_PATTERNS = [
    r"no unsolicited",
    r"unsolicited (?:commercial |marketing )?(?:e-?mails?|messages?|solicitations?) (?:are|is|will) not",
    r"(?:do not|don't|please don't|please do not) (?:send|email|e-mail) (?:us )?(?:any )?(?:unsolicited|marketing|spam|sales)",
    r"we do not (?:accept|want|welcome) (?:any )?(?:unsolicited|marketing|sales) (?:e-?mails?|messages?|approaches|solicitations?)",
    r"not (?:to be )?used for (?:unsolicited|marketing|commercial) (?:purposes|e-?mails?|messages?)",
    r"no (?:spam|marketing e-?mails?|sales (?:calls|emails|enquiries))",
]
_NO_UNSOLICITED_RE = re.compile("|".join(NO_UNSOLICITED_PATTERNS), re.I)
_TAG_RE = re.compile(r"<(script|style)[^>]*>.*?</\1>|<[^>]+>", re.I | re.S)
_CF_EMAIL_RE = re.compile(r'data-cfemail="([0-9a-fA-F]+)"')


def _decode_cfemail(hexstr: str) -> str:
    try:
        key = int(hexstr[:2], 16)
        return "".join(chr(int(hexstr[i:i + 2], 16) ^ key) for i in range(2, len(hexstr), 2))
    except ValueError:
        return ""


def page_text(raw: str) -> str:
    """Visible-ish text of an HTML page plus decoded/obfuscated email forms."""
    extras = [_decode_cfemail(h) for h in _CF_EMAIL_RE.findall(raw)]
    text = html.unescape(_TAG_RE.sub(" ", raw))
    mailtos = re.findall(r'mailto:([^"\'?>\s]+)', raw, re.I)
    text = f"{text} {' '.join(extras)} {' '.join(mailtos)}"
    # "info [at] example [dot] com" style obfuscation
    text = re.sub(r"\s*[\[\(]\s*at\s*[\]\)]\s*", "@", text, flags=re.I)
    text = re.sub(r"\s*[\[\(]\s*dot\s*[\]\)]\s*", ".", text, flags=re.I)
    return re.sub(r"\s+", " ", text)


@dataclass
class Verification:
    email: str = ""
    phone: str = ""
    contact_name: str = ""
    notes: list[str] = field(default_factory=list)
    no_unsolicited: bool = False


class Verifier:
    def __init__(self, enabled: bool = True, timeout: float = 15):
        self.enabled = enabled
        self.timeout = timeout
        self._cache: dict[str, str | None] = {}
        self.session = requests.Session()
        self.session.headers["User-Agent"] = USER_AGENT

    def fetch(self, url: str) -> str | None:
        if url in self._cache:
            return self._cache[url]
        text = None
        try:
            resp = self.session.get(url, timeout=self.timeout, allow_redirects=True)
            if resp.ok and "html" in resp.headers.get("content-type", "html"):
                text = page_text(resp.text)
        except requests.RequestException as exc:
            log.debug("Fetch failed for %s: %s", url, exc)
        self._cache[url] = text
        return text

    def candidate_pages(self, source_url: str, website: str) -> list[str]:
        urls: list[str] = []
        if source_url.startswith("http"):
            urls.append(source_url)
        if website:
            base = website if website.startswith("http") else "https://" + website
            urls.append(base)
            for path in ("/contact", "/contact-us", "/pages/contact", "/contact-us/"):
                urls.append(urljoin(base if base.endswith("/") else base + "/", path.lstrip("/")))
        seen: list[str] = []
        for u in urls:
            if u not in seen:
                seen.append(u)
        return seen[:6]

    def verify(self, *, email: str, phone: str, contact_name: str, source_url: str, website: str,
               evidence_pages: list[tuple[str, str]]) -> Verification:
        email = norm_email(email)
        out = Verification()
        if not self.enabled:
            return Verification(email=email, phone=phone, contact_name=contact_name)
        pages: list[tuple[str, str]] = [(u, page_text(t)) for u, t in evidence_pages]
        site_domain = company_domain(website)

        def relevant(url: str) -> bool:
            return url == source_url or (site_domain and company_domain(url) == site_domain) or not site_domain

        pages = [(u, t) for u, t in pages if relevant(u)]

        def find(predicate) -> tuple[str, str] | None:
            for url, text in pages:
                if predicate(text):
                    return url, text
            for url in self.candidate_pages(source_url, website):
                text = self.fetch(url)
                if text is not None:
                    pages.append((url, text))
                    if predicate(text):
                        return url, text
            return None

        if email:
            hit = find(lambda t: _contains_email(t, email))
            if hit is None:
                out.notes.append(f"Unverified email not used ({email}) - not found on the listed pages")
            elif _NO_UNSOLICITED_RE.search(hit[1]):
                out.no_unsolicited = True
                out.notes.append(f"{host_of(hit[0])} says no unsolicited email - email left blank ({email})")
            else:
                out.email = email
        if phone:
            digits = norm_phone(phone)
            if digits and find(lambda t: digits in _phone_digits(t)):
                out.phone = phone
            else:
                out.notes.append(f"Unverified phone not used ({phone})")
        if contact_name:
            first = contact_name.split()[0].lower()
            if find(lambda t: contact_name.lower() in t.lower() or (len(first) > 2 and first in t.lower())):
                out.contact_name = contact_name
            else:
                out.notes.append(f"Unverified contact name not used ({contact_name})")
        return out


def _contains_email(text: str, email: str) -> bool:
    return re.search(r"(?<![\w.+-])" + re.escape(email) + r"(?![\w-]|\.\w)", text.lower()) is not None


def _phone_digits(text: str) -> str:
    """All phone-like digit runs on the page, normalised to local format."""
    runs = re.findall(r"\+?\d[\d\s().-]{6,}\d", text)
    return " ".join(norm_phone(r) for r in runs)
