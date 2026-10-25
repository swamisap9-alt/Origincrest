"""Duplicate detection across every tab of every workbook."""

from __future__ import annotations

import difflib
import re
import unicodedata
from urllib.parse import urlparse

FREE_MAIL_DOMAINS = {
    "gmail.com", "googlemail.com", "outlook.com", "outlook.com.au", "hotmail.com", "hotmail.com.au",
    "live.com", "live.com.au", "msn.com", "yahoo.com", "yahoo.com.au", "y7mail.com", "icloud.com", "me.com",
    "mac.com", "aol.com", "protonmail.com", "proton.me", "bigpond.com", "bigpond.net.au", "bigpond.com.au",
    "optusnet.com.au", "iinet.net.au", "tpg.com.au", "internode.on.net", "westnet.com.au", "dodo.com.au",
    "adam.com.au", "ozemail.com.au", "zoho.com", "gmx.com", "mail.com",
}

# Sites that host many businesses: a link to them says nothing about which company it is.
NON_COMPANY_DOMAINS = {
    "facebook.com", "instagram.com", "linkedin.com", "linktr.ee", "google.com", "goo.gl", "g.page",
    "yellowpages.com.au", "truelocal.com.au", "yelp.com", "yelp.com.au", "beanhunter.com", "tripadvisor.com",
    "tripadvisor.com.au", "twitter.com", "x.com", "tiktok.com", "youtube.com", "hotfrog.com.au",
    "localsearch.com.au", "startlocal.com.au", "abr.business.gov.au", "zomato.com", "ubereats.com",
    "doordash.com", "broadsheet.com.au", "timeout.com", "wordofmouth.com.au", "cylex.com.au",
    "aussieweb.com.au", "yably.com.au", "pinterest.com", "wikipedia.org",
}

# Site builders where the subdomain is the business (keep the full host).
HOSTED_SUBDOMAIN_DOMAINS = {"wixsite.com", "square.site", "business.site", "myshopify.com", "squarespace.com"}

_SECOND_LEVEL_SUFFIXES = {"com.au", "net.au", "org.au", "asn.au", "id.au", "edu.au", "gov.au", "co.nz",
                          "co.uk", "org.uk", "co.id", "com.sg", "com.my"}

_LEGAL_WORDS = {"pty", "ltd", "limited", "proprietary", "inc", "llc", "co", "company", "corp", "corporation",
                "holdings", "the", "and", "australia", "aust", "au", "group", "trading", "as", "t"}
_GENERIC_WORDS = {"coffee", "coffees", "roasters", "roaster", "roastery", "roasting", "roasted", "espresso",
                  "freight", "logistics", "forwarding", "forwarders", "international", "global", "services",
                  "transport", "shipping", "warehousing", "warehouse", "3pl", "customs", "brokers", "broker",
                  "brokerage", "supply", "chain", "solutions", "specialty", "speciality", "beans", "cafe", "co"}

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+'-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
URL_RE = re.compile(r"(?:https?://|www\.)[^\s,;<>\"')]+", re.I)


def _ascii(text: str) -> str:
    return unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode()


def name_tokens(name: str) -> list[str]:
    text = _ascii(str(name or "")).lower().replace("&", " and ")
    text = re.sub(r"\(.*?\)", " ", text)  # "(Adelaide branch)" etc.
    return [t for t in re.split(r"[^a-z0-9]+", text) if t]


def core_name(name: str) -> str:
    return " ".join(t for t in name_tokens(name) if t not in _LEGAL_WORDS)


def loose_name(name: str) -> str:
    key = " ".join(t for t in name_tokens(name) if t not in _LEGAL_WORDS and t not in _GENERIC_WORDS)
    return key if len(key.replace(" ", "")) >= 5 else ""


def host_of(url: str) -> str:
    url = str(url or "").strip()
    if not url:
        return ""
    if "@" in url and "/" not in url:
        return ""
    if not re.match(r"^[a-z]+://", url, re.I):
        url = "http://" + url
    try:
        host = (urlparse(url).hostname or "").lower()
    except ValueError:
        return ""
    return host[4:] if host.startswith("www.") else host


def registrable_domain(host: str) -> str:
    parts = [p for p in host.lower().split(".") if p]
    if len(parts) < 2:
        return ""
    if ".".join(parts[-2:]) in _SECOND_LEVEL_SUFFIXES and len(parts) >= 3:
        return ".".join(parts[-3:])
    return ".".join(parts[-2:])


def company_domain(url_or_host: str) -> str:
    """Domain that identifies the business, or '' for directories/social sites."""
    host = host_of(url_or_host)
    if not host or "." not in host:
        return ""
    reg = registrable_domain(host)
    if reg in HOSTED_SUBDOMAIN_DOMAINS:
        return host
    if reg in NON_COMPANY_DOMAINS:
        return ""
    return reg


def email_domain(email: str) -> str:
    email = str(email or "").strip().lower()
    if "@" not in email:
        return ""
    return email.rsplit("@", 1)[1]


def is_free_mail(email: str) -> bool:
    return email_domain(email) in FREE_MAIL_DOMAINS


def norm_email(email: str) -> str:
    email = str(email or "").strip().lower()
    if email.startswith("mailto:"):
        email = email[7:]
    return email.split("?")[0]


def norm_phone(phone: str) -> str:
    digits = re.sub(r"\D", "", str(phone or ""))
    if digits.startswith("61") and len(digits) >= 11:
        digits = "0" + digits[2:]
    return digits[-9:] if len(digits) >= 8 else ""


def entity_keys(company: str = "", website: str = "", email: str = "", phone: str = "",
                extra_text: str = "") -> set[str]:
    keys: set[str] = set()
    if core := core_name(company):
        keys.add("name:" + core)
    if loose := loose_name(company):
        keys.add("loose:" + loose)
    if dom := company_domain(website):
        keys.add("domain:" + dom)
    emails = {norm_email(email)} if email else set()
    emails |= {norm_email(e) for e in EMAIL_RE.findall(extra_text or "")}
    for addr in filter(None, emails):
        keys.add("email:" + addr)
        if not is_free_mail(addr) and (dom := company_domain(email_domain(addr))):
            keys.add("domain:" + dom)
    for url in URL_RE.findall(extra_text or ""):
        if dom := company_domain(url):
            keys.add("domain:" + dom)
    if ph := norm_phone(phone):
        keys.add("phone:" + ph)
    return keys


def primary_key(company: str = "", website: str = "", email: str = "") -> str:
    """Stable id for a company: domain, else email, else normalised name."""
    if dom := company_domain(website):
        return "domain:" + dom
    addr = norm_email(email)
    if addr and not is_free_mail(addr) and (dom := company_domain(email_domain(addr))):
        return "domain:" + dom
    if addr:
        return "email:" + addr
    return "name:" + core_name(company)


class DedupeIndex:
    def __init__(self) -> None:
        self._keys: dict[str, str] = {}
        self._names: list[tuple[str, str]] = []

    def __len__(self) -> int:
        return len(self._names)

    def add(self, label: str, company: str = "", website: str = "", email: str = "", phone: str = "",
            extra_text: str = "") -> None:
        for key in entity_keys(company, website, email, phone, extra_text):
            self._keys.setdefault(key, label)
        core = core_name(company)
        if core:
            self._names.append((core, label))

    def add_keys(self, label: str, keys: set[str] | list[str]) -> None:
        for key in keys:
            self._keys.setdefault(key, label)
            if key.startswith("name:"):
                self._names.append((key[5:], label))

    def match(self, company: str = "", website: str = "", email: str = "", phone: str = "") -> str | None:
        for key in entity_keys(company, website, email, phone):
            if key in self._keys:
                return self._keys[key]
        core = core_name(company)
        if len(core) >= 8:
            for other, label in self._names:
                if len(other) >= 8 and difflib.SequenceMatcher(None, core, other).ratio() >= 0.92:
                    return label
        return None

    def known_names(self, limit: int = 300) -> list[str]:
        seen: list[str] = []
        for core, _ in reversed(self._names):
            if core not in seen:
                seen.append(core)
            if len(seen) >= limit:
                break
        return seen
