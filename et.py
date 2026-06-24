import os
import re
import csv
import sys
import time
import json
import base64
import hashlib
import logging
from datetime import datetime, timedelta
from urllib.parse import urljoin, urlparse, urlsplit, urlunsplit, parse_qsl, urlencode

import requests
from bs4 import BeautifulSoup

# Optional: load secrets from a local .env file if python-dotenv is installed.
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# Optional heavy deps used for Excel export only.
try:
    import pandas as pd
    import openpyxl
    _XLSX_AVAILABLE = True
except ImportError:
    _XLSX_AVAILABLE = False

# Optional heavy deps used for paraphrase quality gating.
try:
    import language_tool_python
    from sentence_transformers import SentenceTransformer, util as st_util
    _NLP_AVAILABLE = True
except ImportError:
    _NLP_AVAILABLE = False

# =============================================================================
#  CONFIG
# =============================================================================
#
#  SOURCE
#  ------
#  https://ethiojobs.net/jobs  — the live job board.
#
#  ethiojobs.net is a Next.js single-page app: the HTML you get from a plain
#  GET request is just a loading screen (a <video> tag), with NO job content in
#  it. So there are no HTML job cards to parse — BeautifulSoup
#  on the page body would read zero jobs. Instead the page fetches all its data
#  in-browser from a JSON backend (api.ethiojobs.net). This scraper calls that
#  JSON API directly with plain `requests` — no browser, no Playwright, lighter
#  and faster than HTML scraping. The exact API base/endpoints are discovered
#  automatically on first run (see discover_api) by reading the site's own
#  JavaScript bundles, then cached to API_CACHE_FILE. You can override any of it
#  via the env vars below, or run `python <thisfile> --discover` to print what
#  was found.
#
#  APPLY RULE (hard, network-wide)
#  -------------------------------
#  A job only posts if it exposes a PUBLIC apply path: an email or an external
#  apply URL in its body / how-to-apply text. EthioJobs' own "Apply Now" button
#  is on-platform (login required), so the EthioJobs page is NOT treated as a
#  valid apply destination by default. Jobs without a public email/URL are
#  written to the flagged CSV. REQUIRE_PUBLIC_APPLY (default "1"/on) enforces
#  this; set to "0" to post everything. Set APPLY_VIA_SOURCE_URL=1 to instead
#  treat each EthioJobs job page as the external apply target (a per-source
#  policy opt-in, default off).
# =============================================================================

BASE_URL = "https://ethiojobs.net"

# Listing page(s) on the public site (used for the human-facing job_url + as the
# discovery seed). Comma-separated env override. Examples:
#   all jobs:            https://ethiojobs.net/jobs
#   by category:         https://ethiojobs.net/jobs/accounting-and-finance
#   by region:           https://ethiojobs.net/jobs/region/addis-ababa
JOBS_URL  = os.environ.get("ETHIOJOBS_JOBS_URL", "https://ethiojobs.net/jobs")
LISTING_URLS = [u.strip() for u in
                os.environ.get("ETHIOJOBS_LISTING_URLS", JOBS_URL).split(",") if u.strip()]

# JSON backend. Discovered automatically on first run; these are the defaults /
# overrides. API_BASE is the host; the list/detail PATHS are tried in order until
# one returns job JSON, and the winner is cached. {id} is filled with the job's id.
API_BASE          = os.environ.get("ETHIOJOBS_API_BASE", "https://api.ethiojobs.net")
API_LIST_PATHS    = [p.strip() for p in os.environ.get(
    "ETHIOJOBS_API_LIST_PATHS",
    "/job/search,/jobs/search,/job/list,/jobs,/job,/vacancy/search,/vacancies"
).split(",") if p.strip()]
API_DETAIL_PATHS  = [p.strip() for p in os.environ.get(
    "ETHIOJOBS_API_DETAIL_PATHS",
    "/job/{id},/jobs/{id},/vacancy/{id},/job/detail/{id},/job/get/{id}"
).split(",") if p.strip()]
API_CACHE_FILE    = os.environ.get("ETHIOJOBS_API_CACHE", "ethiojobs_api_cache.json")

# Enforce the public-apply-only rule (email or external URL required to post).
REQUIRE_PUBLIC_APPLY = os.environ.get("REQUIRE_PUBLIC_APPLY", "1") != "0"
# Treat the EthioJobs job page itself as the external apply target (on-platform apply).
APPLY_VIA_SOURCE_URL = os.environ.get("APPLY_VIA_SOURCE_URL", "0") == "1"

REQUEST_DELAY   = float(os.environ.get("REQUEST_DELAY", "1.0"))
MAX_JOBS        = int(os.environ.get("MAX_JOBS", "0"))     # 0 = unlimited
MAX_PAGES       = int(os.environ.get("MAX_PAGES", "50"))   # listing pagination cap
PAGE_SIZE       = int(os.environ.get("PAGE_SIZE", "20"))   # records requested per page
REQUEST_TIMEOUT = int(os.environ.get("REQUEST_TIMEOUT", "25"))

OUTPUT_FILE        = "ethiojobs_ethiopia_jobs.xlsx"
PROCESSED_IDS_FILE = "ethiojobs_ethiopia_processed.csv"
FLAGGED_FILE       = "ethiojobs_ethiopia_flagged.csv"

# CSV column names — defined once so _init_tracker, load, and upsert all agree.
_TRACKER_FIELDS = ["Job ID", "Job URL", "Job Title", "Company Name",
                   "Status", "Timestamp", "WP ID"]

_FLAGGED_FIELDS = ["Source", "Title", "Company", "Location", "Salary",
                   "Deadline", "Reason", "Apply Note", "Job URL", "Timestamp"]

# ── WordPress ────────────────────────────────────────────────────────────────
WP_URL      = os.environ.get("WP_BASE_URL", "")
WP_USER     = os.environ.get("WP_USERNAME", "")
WP_PASSWORD = os.environ.get("WP_APP_PASSWORD", "")
WP_BASE      = WP_URL.rstrip("/")
WP_JOBS_URL  = f"{WP_BASE}/job-listings"
WP_MEDIA_URL = f"{WP_BASE}/media"

# ── Mistral ──────────────────────────────────────────────────────────────────
MISTRAL_API_KEY = os.environ.get("MISTRAL_API_KEY", "")
MISTRAL_MODEL   = "mistral-small-latest"
MISTRAL_URL     = "https://api.mistral.ai/v1/chat/completions"

ENABLE_PARAPHRASE = True

# ── Startup warnings ─────────────────────────────────────────────────────────
for _var, _val, _feature in [
    ("MISTRAL_API_KEY", MISTRAL_API_KEY, "paraphrasing"),
    ("WP_USERNAME",     WP_USER,         "WordPress posting"),
    ("WP_APP_PASSWORD", WP_PASSWORD,     "WordPress posting"),
]:
    if not _val:
        logging.getLogger(__name__).warning(
            f"Environment variable {_var} is not set — {_feature} will be disabled/skipped."
        )

JOB_TYPE_MAPPING = {
    "full-time": "full-time", "full time": "full-time",
    "part-time": "part-time", "part time": "part-time",
    "contract":  "contract",  "temporary": "temporary",
    "internship":"internship","freelance": "freelance",
    "volunteer": "volunteer", "permanent": "full-time",
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml",
    "Accept-Charset": "utf-8",
}

SESSION = requests.Session()
SESSION.headers.update(HEADERS)

# Known Ethiopian towns/cities/regions, used as a fallback to pull a location
# from free text when the API location field is empty.
ETHIOPIA_LOCATIONS = [
    "Addis Ababa", "Adama", "Nazret", "Dire Dawa", "Mekelle", "Mekele",
    "Bahir Dar", "Hawassa", "Awassa", "Gondar", "Jimma", "Dessie", "Jijiga",
    "Shashamane", "Bishoftu", "Debre Birhan", "Sodo", "Arba Minch", "Hosaena",
    "Harar", "Dilla", "Nekemte", "Debre Markos", "Asella", "Ambo", "Woldia",
    "Gambela", "Assosa", "Semera", "Adigrat", "Axum", "Shire", "Sebeta",
    "Burayu", "Gelan", "Modjo", "Mojo", "Ziway", "Batu", "Bule Hora",
    "Oromia", "Amhara", "Tigray", "Sidama", "Afar", "Somali", "Benishangul",
    "SNNPR", "Gambella", "Dukem", "Holeta",
]
# EthioJobs uses "Ethiopia" as the country-level catch-all location.
DEFAULT_LOCATION = os.environ.get("ETHIOJOBS_DEFAULT_LOCATION", "Ethiopia")

# Hosts/paths that are never a real external apply destination.
_NON_APPLY_HOST_SUBSTR = (
    "ethiojobs.net", "ethiojobs.org", "ethiojobs.com", "dereja.com",
    "facebook.", "twitter.", "x.com", "linkedin.",
    "instagram.", "wa.me", "whatsapp", "t.me", "telegram",
    "plus.google", "pinterest.", "youtube.", "tiktok.",
)
_NON_APPLY_PATH_SUBSTR = (
    "/login", "/signin", "/sign-in", "/register", "/signup", "/sign-up",
    "action=login", "mode=register", "#share", "/share",
    "/cart", "/checkout", "/account",
)
# Emails belonging to the board itself are never a real apply address.
_NON_APPLY_EMAIL_DOMAINS = ("ethiojobs.net", "ethiojobs.org", "dereja.com")

def _is_real_apply_email(email: str) -> bool:
    if not email or "@" not in email:
        return False
    dom = email.rsplit("@", 1)[-1].lower()
    return not any(dom == d or dom.endswith("." + d) for d in _NON_APPLY_EMAIL_DOMAINS)

# =============================================================================
#  LOGGING / COLOUR
# =============================================================================

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log_ = logging.getLogger(__name__)

_USE_COLOUR = sys.stdout.isatty()

def _c(code, text):
    return f"\033[{code}m{text}\033[0m" if _USE_COLOUR else text

C_HEADER  = lambda t: _c("1;36",  t)
C_LABEL   = lambda t: _c("1;33",  t)
C_VALUE   = lambda t: _c("97",    t)
C_DIM     = lambda t: _c("2",     t)
C_GREEN   = lambda t: _c("1;32",  t)
C_RED     = lambda t: _c("1;31",  t)
C_BLUE    = lambda t: _c("1;34",  t)
C_DIVIDER = lambda: _c("2", "─" * 80)

def log(msg):
    print(msg, flush=True)

EMAIL_PATTERN = re.compile(r"[A-Za-z0-9.+_-]+@[A-Za-z0-9-]+\.[A-Za-z0-9.-]+")
URL_PATTERN   = re.compile(r"https?://[^\s)>\"']+", re.I)

TRACKING_PARAM_PREFIXES = ("utm_",)
TRACKING_PARAM_EXACT = {
    "fbclid", "gclid", "msclkid", "mc_cid", "mc_eid", "ref", "referrer",
}

MONTHS = {
    "january": 1, "february": 2, "march": 3, "april": 4, "may": 5, "june": 6,
    "july": 7, "august": 8, "september": 9, "october": 10, "november": 11,
    "december": 12, "jan": 1, "feb": 2, "mar": 3, "apr": 4, "jun": 6, "jul": 7,
    "aug": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dec": 12,
}

# Ordinal text date e.g. "30th June 2026" / "7th July, 2026".
TEXT_DATE_RE = re.compile(
    r"(\d{1,2})\s*(?:st|nd|rd|th)?\s+([A-Za-z]+)\s*[.,]?\s*(\d{4})", re.I
)
# Numeric DD/MM/YYYY or DD-MM-YYYY.
DMY_DATE_RE = re.compile(r"\b(\d{1,2})[/-](\d{1,2})[/-](\d{4})\b")
# Month-first text date — EthioJobs' native display form, e.g.
# "June 17th, 2026", "June 24/2026", "Sept 1 2026", "January 3, 2027".
MDY_TEXT_DATE_RE = re.compile(
    r"([A-Za-z]{3,9})\s+(\d{1,2})\s*(?:st|nd|rd|th)?\s*[,/]?\s*(\d{4})", re.I
)
# ISO 8601 from the JSON API, e.g. "2026-06-17" or "2026-06-17T09:00:00Z".
ISO_DATE_RE = re.compile(r"(?<!\d)(\d{4})-(\d{1,2})-(\d{1,2})(?!\d)")

# Labels inside the JobMonster "Job Overview" box.
DEADLINE_LABELS = ("application deadline", "closing date", "deadline",
                   "expiry date", "expires")

# Body headings that introduce the application instructions. Matched against a
# *stripped, short* line (see _is_apply_heading_line) so it never trips on
# 'Application Deadline:' / 'Application Format' or a mid-sentence 'to apply'.
_APPLY_HEAD_PHRASES = re.compile(
    r"^(?:how\s*(?:and|&)\s*deadline\s*to\s*apply|how\s*to\s*apply(?:\s*(?:and|&)\s*deadline)?|"
    r"how\s*to\s*submit|to\s*apply|application\s*(?:and|&)\s*deadline|"
    r"mode\s*of\s*application|method\s*of\s*application|"
    r"application\s*(?:procedure|process|instructions?|method|guidelines?)|"
    r"submission\s*of\s*applications?|deadline\s*(?:and|&)?\s*(?:how\s*)?to\s*apply)\b",
    re.I,
)

# Boilerplate that marks the end of usable post content on a detail page.
_BODY_CUT_MARKERS = [
    "related jobs", "leave your thoughts", "you must be logged in",
    "email me jobs like these", "send to a friend", "company information",
    "leave a reply", "post a comment",
]
# Standalone UI lines to drop from the description.
_BODY_DROP_LINES = {
    "apply for this job", "save", "share", "share:", "bookmark job",
    "quick view", "send to friend", "send to a friend", "clear all",
    "filter", "view more",
}

# =============================================================================
#  TEXT CLEANUP / SANITIZATION
# =============================================================================

_MOJIBAKE = [
    ("Â", ""), ("â€™", "'"), ("â€œ", '"'), ("â€\x9d", '"'), ("â€", '"'),
    ("â€¢", "•"), ("â„¢", "™"), ("\u00a0", " "), ("\u200b", ""), ("\ufeff", ""),
]

def _fix_mojibake(text: str) -> str:
    for pattern, replacement in _MOJIBAKE:
        text = text.replace(pattern, replacement)
    text = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", "", text)
    return text

def sanitize_text(text, is_url=False) -> str:
    if not isinstance(text, str):
        text = str(text) if (text is not None and str(text) not in ("nan", "None", "NaN")) else ""
    text = text.strip()
    if text in ("nan", "None", "NaN", "", "N/A", "n/a", "NA", "na"):
        return ""
    text = _fix_mojibake(text)
    if is_url:
        return re.sub(r"[ \t\r\n\f\v]+", " ", text).strip()
    text = re.sub(r"#+\s*", "", text)
    text = re.sub(r"\*\*", "", text)
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()

def clean_text(el):
    if el is None:
        return ""
    return re.sub(r"\s+", " ", el.get_text(" ", strip=True)).strip()

def extract_email(text):
    if not text:
        return ""
    m = EMAIL_PATTERN.search(text)
    return m.group(0) if m else ""

def strip_tracking_params(url):
    if not url:
        return url
    parts = urlsplit(url)
    if not parts.query:
        return url
    kept = []
    for key, value in parse_qsl(parts.query, keep_blank_values=True):
        key_lower = key.lower()
        if key_lower.startswith(TRACKING_PARAM_PREFIXES) or key_lower in TRACKING_PARAM_EXACT:
            continue
        kept.append((key, value))
    new_query = urlencode(kept)
    return urlunsplit((parts.scheme, parts.netloc, parts.path, new_query, parts.fragment))

# =============================================================================
#  BASIC HTTP / PARSING HELPERS
# =============================================================================

def get_soup(url):
    resp = SESSION.get(url, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()
    resp.encoding = resp.encoding or "utf-8"
    try:
        return BeautifulSoup(resp.text, "lxml")
    except Exception:
        return BeautifulSoup(resp.text, "html.parser")

def slugify(text, maxlen=80):
    s = re.sub(r"[^a-z0-9]+", "-", (text or "").lower()).strip("-")
    return s[:maxlen] or "job"

def html_block_to_text(el) -> str:
    """
    Convert a BeautifulSoup element to readable plain text, preserving line
    breaks for block-level tags and turning <li> into bullet lines. The block is
    mutated in place — only ever call this on a throwaway/per-job element.
    """
    if el is None:
        return ""
    for br in el.find_all("br"):
        br.replace_with("\n")
    for li in el.find_all("li"):
        txt = li.get_text(" ", strip=True)
        li.replace_with("\n• " + txt + "\n")
    for tag in el.find_all(["p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "tr"]):
        tag.insert_before("\n")
        tag.insert_after("\n")
    text = el.get_text("\n")
    text = _fix_mojibake(text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n[ \t]+", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()

# =============================================================================
#  DATE / FIELD EXTRACTORS
# =============================================================================

def dmy_dates(text: str) -> list:
    """Return ISO dates parsed from DD/MM/YYYY (or DD-MM-YYYY), in order."""
    out = []
    for d, m, y in DMY_DATE_RE.findall(text or ""):
        try:
            out.append(datetime(int(y), int(m), int(d)).strftime("%Y-%m-%d"))
        except ValueError:
            pass
    return out

def text_dates(text: str) -> list:
    """Return ISO dates parsed from ordinal text form ('30th June 2026'), in order."""
    out = []
    for d, mon, y in TEXT_DATE_RE.findall(text or ""):
        month = MONTHS.get(mon.lower())
        if not month:
            continue
        try:
            out.append(datetime(int(y), month, int(d)).strftime("%Y-%m-%d"))
        except ValueError:
            pass
    return out

def mdy_text_dates(text: str) -> list:
    """Return ISO dates parsed from month-first text ('June 17th, 2026'), in order."""
    out = []
    for mon, d, y in MDY_TEXT_DATE_RE.findall(text or ""):
        month = MONTHS.get(mon.lower())
        if not month:
            continue
        try:
            out.append(datetime(int(y), month, int(d)).strftime("%Y-%m-%d"))
        except ValueError:
            pass
    return out

def iso_dates(text: str) -> list:
    """Return ISO dates parsed from YYYY-MM-DD (API form), in order."""
    out = []
    for y, m, d in ISO_DATE_RE.findall(text or ""):
        try:
            out.append(datetime(int(y), int(m), int(d)).strftime("%Y-%m-%d"))
        except ValueError:
            pass
    return out

def parse_any_date(text: str) -> str:
    """
    Best single date from a value. EthioJobs returns ISO from the API and shows
    month-first text on the page ('June 17th, 2026'); legacy/free text may be
    day-first or numeric. Try ISO and month-first first (native forms), then fall
    back to day-first text and numeric DD/MM/YYYY.
    """
    for fn in (iso_dates, mdy_text_dates, text_dates, dmy_dates):
        ds = fn(text)
        if ds:
            return ds[-1]
    return ""

def clean_title(raw: str) -> str:
    """Strip the trailing ' 1442 views' / 'views' counter from a detail H1."""
    t = sanitize_text(raw)
    t = re.sub(r"\s*\d[\d,]*\s*views?\s*$", "", t, flags=re.I)
    t = re.sub(r"\s*views?\s*$", "", t, flags=re.I)
    return t.strip()

def map_job_type(raw: str) -> str:
    key = (raw or "").lower().strip()
    return JOB_TYPE_MAPPING.get(key, "full-time")

def pick_location(locations: list) -> str:
    """Prefer a specific town over the country-level 'Ethiopia' catch-all."""
    specific = [l for l in locations if l and l.strip().lower() not in ("ethiopia",)]
    if specific:
        return specific[0].strip()
    if locations:
        return locations[0].strip()
    return DEFAULT_LOCATION

def location_from_text(text: str) -> str:
    if text:
        for town in ETHIOPIA_LOCATIONS:
            if re.search(rf"\b{re.escape(town)}\b", text, re.I):
                if town.lower() == "mekele":   return "Mekelle"
                if town.lower() == "nazret":   return "Adama"
                if town.lower() == "awassa":   return "Hawassa"
                if town.lower() == "mojo":     return "Modjo"
                if town.lower() == "batu":     return "Ziway"
                return town
    return DEFAULT_LOCATION

def extract_experience(qual_text: str) -> str:
    if not qual_text:
        return ""
    m = re.search(r"(?:at least|minimum(?: of)?)\s+\d+\s+years?[^.\n;]*", qual_text, re.I)
    if m:
        return m.group(0).strip().rstrip(".")
    m = re.search(r"\b\d+\s+years?[^.\n;]*experience", qual_text, re.I)
    if m:
        return m.group(0).strip().rstrip(".")
    return ""

def extract_salary(text: str) -> str:
    """Best-effort salary. EthioJobs often omits a figure; returns '' if none."""
    if not text:
        return ""
    m = re.search(r"(?:GMD|D|GMD\s|D\s|₵)\s*([0-9]{1,3}(?:,\s?[0-9]{3})+(?:\.[0-9]+)?)", text)
    if m:
        amt = re.sub(r"\s+", "", m.group(1))
        return f"GMD {amt}"
    m = re.search(r"\b(?:salary|remuneration)\b[^.\n]{0,80}", text, re.I)
    if m and re.search(r"\d", m.group(0)):
        return m.group(0).strip().rstrip(".")
    return ""

# =============================================================================
#  CANONICAL NORMALISERS  (shared schema — qualification tier / experience band /
#  job field). These mirror the mappings used across the other country pipelines
#  so EthioJobs rows land in the same shape: a TIER label for qualification, a BAND
#  label for experience, and a single canonical FIELD rather than the site's raw
#  multi-category tag dump.
# =============================================================================

# Keyword matcher shared by the tier/field maps. Short or ambiguous tokens (<=3
# chars, e.g. "pa", "hr", "ma", "qa", "0 years") must match as whole words so they
# don't fire inside longer words ("diploma", "patrol", "40 years"); longer tokens
# keep prefix behaviour so "developer" still catches "developers".
def _kw_hit(text_low: str, keywords) -> bool:
    for k in keywords:
        kk = k.strip().lower()
        if not kk:
            continue
        esc = re.escape(kk)
        if len(kk) <= 3:
            # acronyms/codes -> exact whole token
            pat = r"(?<![a-z0-9])" + esc + r"(?![a-z0-9])"
        else:
            # whole token, tolerating a plural 's'/'es' (developer->developers)
            # but NOT an arbitrary suffix (quota !-> quotation, ma !-> diploma)
            pat = r"(?<![a-z0-9])" + esc + r"(?:es|s)?(?![a-z0-9])"
        if re.search(pat, text_low):
            return True
    return False

# --- Qualification: text -> single tier label -------------------------------
QUALIFICATION_TIERS = [
    ("PhD / Doctorate",          ["phd", "ph.d", "doctorate", "doctoral", "doctor of philosophy"]),
    ("Master's Degree",          ["master", "msc", "m.sc", "ma ", "m.a ", "mba", "m.b.a", "meng",
                                  "m.eng", "mphil", "postgraduate", "post-graduate", "post graduate"]),
    ("Bachelor's Degree",        ["bachelor", "bsc", "b.sc", "ba ", "b.a ", "beng", "b.eng", "bcom",
                                  "b.com", "bba", "llb", "degree in", "undergraduate degree",
                                  "honours degree", "hons"]),
    ("Higher National Diploma",  ["hnd", "hnc", "higher national diploma", "higher national certificate",
                                  "higher diploma", "advanced diploma"]),
    ("Diploma",                  ["diploma", "dip ", "dip.", "associate degree", "foundation degree"]),
    ("Professional Certification", ["acca", "cpa", "cfa", "cima", "pmp", "prince2", "cissp",
                                    "aws certified", "comptia", "cisco", "ccna", "ccnp", "shrm",
                                    "cipd", "chartered", "certified public", "certified financial",
                                    "certified project", "professional certification",
                                    "professional certificate"]),
    ("A-Levels / HSC",           ["a-level", "a level", "hsc", "higher school certificate", "ib diploma",
                                  "international baccalaureate", "gce advanced"]),
    ("O-Levels / School Certificate", ["o-level", "o level", "igcse", "gcse", "school certificate",
                                       "sc ", "cpe", "certificate of primary"]),
    ("No Formal Qualification Required", ["no qualification", "no degree", "no formal", "school leaver",
                                          "entry level", "no experience required", "training provided",
                                          "will train"]),
]

def extract_qualification(text: str) -> str:
    if not text:
        return ""
    # Skip school-admission notices that merely mention pupils' ages/levels.
    if re.search(r"nursery|primary years|ib pyp|aged between|boys and girls", text, re.I):
        return ""
    lower = text.lower()
    for label, keywords in QUALIFICATION_TIERS:
        if _kw_hit(lower, keywords):
            return label
    return ""

# --- Experience: text -> single band label ----------------------------------
NO_EXP_KW = ["no experience", "no prior experience", "fresh graduate", "freshers",
             "entry level", "entry-level", "0 years", "zero experience",
             "training provided", "will train", "no experience required"]
LESS1_KW  = ["less than 1 year", "under 1 year", "6 months", "less than a year",
             "some experience", "minimal experience"]

def years_to_band(n: int) -> str:
    if n <= 0:  return "No Experience Required"
    if n <= 2:  return "1 - 2 Years"
    if n <= 5:  return "3 - 5 Years"
    if n <= 10: return "6 - 10 Years"
    return "10+ Years"

# Only treat a number as a *requirement* when it sits in an experience context,
# so org-history phrasing ("established 40 years ago", "since 1982") is ignored.
# A real job requirement is also capped at a sane ceiling.
_EXP_CAP = 20
_EXP_REQ_RE = re.compile(
    r"(?:minimum|min\.?|at\s+least|atleast|least|over|more\s+than|not\s+less\s+than|"
    r"minimum\s+of|a\s+minimum\s+of)\s+(?:of\s+)?(\d{1,2})\s*\+?\s*years?", re.I)
# "N years of <work/experience/…>" is itself a tenure requirement, even without a
# 'minimum' prefix and even when 'experience' is far away in the sentence.
_EXP_YEARS_OF_RE = re.compile(r"(\d{1,2})\s*\+?\s*years?\s+of\b", re.I)
_EXP_ANY_YEARS_RE = re.compile(r"(\d{1,2})\s*\+?\s*years?", re.I)
_EXP_RANGE_RE = re.compile(r"(\d{1,2})\s*(?:-|–|to)\s*(\d{1,2})\s*years?", re.I)

def extract_experience_band(text: str) -> str:
    """Map free text to one of the canonical experience bands (or '')."""
    if not text:
        return ""
    low = text.lower()
    years = []
    # (a) explicit requirement phrasing: "minimum 3 years", "at least 5 years"
    for m in _EXP_REQ_RE.finditer(text):
        n = int(m.group(1))
        if 0 < n <= _EXP_CAP:
            years.append(n)
    # (b) "N years of ..." tenure construction (capped, so org-history "40 years
    #     of experience" is filtered out by _EXP_CAP).
    for m in _EXP_YEARS_OF_RE.finditer(low):
        n = int(m.group(1))
        if 0 < n <= _EXP_CAP:
            years.append(n)
    # (c) "N years ... experience" with 'experience' near the figure
    for m in _EXP_ANY_YEARS_RE.finditer(low):
        n = int(m.group(1))
        if 0 < n <= _EXP_CAP and "experien" in low[m.end():m.end() + 60]:
            years.append(n)
    # (d) ranges: "3-5 years", "3 to 5 years"
    for m in _EXP_RANGE_RE.finditer(text):
        a = int(m.group(1))
        if 0 < a <= _EXP_CAP:
            years.append(a)
    if years:
        return years_to_band(min(years))           # explicit figure wins
    if _kw_hit(low, NO_EXP_KW):
        return "No Experience Required"
    if _kw_hit(low, LESS1_KW):
        return "1 - 2 Years"                       # floor band (no sub-1yr bucket)
    return ""

# --- Job field: title+description -> single canonical field -----------------
# (field, strong keywords, weak keywords). Strong matches win over weak; the
# list order is the tie-break priority.
FIELD_KEYWORD_MAP = [
    ("Information Technology",
     ["software engineer", "developer", "devops", "frontend", "backend", "full stack", "fullstack",
      "sysadmin", "cloud", "cybersecurity", "data engineer", "machine learning", "artificial intelligence",
      "ai/ml", "it support", "network engineer", "database", "kubernetes", "docker", "aws", "azure",
      "react", "node.js", "python developer", "java developer"],
     ["programming", "coding", "api", "agile", "scrum", "git", "linux", "server", "infrastructure", "software"]),
    ("Finance & Accounting",
     ["accountant", "auditor", "finance manager", "financial analyst", "cfo", "treasurer", "tax",
      "bookkeeper", "payroll", "budget analyst", "credit analyst", "investment", "portfolio manager",
      "risk analyst", "forex", "actuary", "acca", "cfa", "cpa"],
     ["financial", "accounting", "balance sheet", "p&l", "reconciliation", "ifrs", "gaap", "ledger", "invoicing"]),
    ("Sales & Business Development",
     ["sales executive", "sales manager", "business development", "account manager",
      "sales representative", "bd manager", "regional sales", "key account", "sales director",
      "commercial manager", "sales officer"],
     ["revenue", "pipeline", "crm", "leads", "prospects", "quota", "target", "upsell", "cross-sell", "b2b", "b2c"]),
    ("Marketing & Communications",
     ["marketing manager", "digital marketing", "seo", "sem", "content marketer", "social media manager",
      "brand manager", "marketing executive", "communications manager", "pr manager", "copywriter",
      "growth hacker", "email marketing", "campaign manager"],
     ["marketing", "branding", "advertising", "social media", "content", "campaign", "analytics",
      "google ads", "facebook ads", "influencer"]),
    ("Human Resources",
     ["hr manager", "human resources", "recruiter", "talent acquisition", "hr business partner",
      "hrbp", "hr officer", "compensation", "benefits manager", "organisational development",
      "learning and development", "l&d", "hr generalist", "payroll manager"],
     ["recruitment", "onboarding", "performance management", "employee relations", "hr", "workforce"]),
    ("Engineering",
     ["mechanical engineer", "civil engineer", "electrical engineer", "structural engineer",
      "process engineer", "project engineer", "maintenance engineer", "production engineer",
      "quality engineer", "safety engineer", "site engineer", "design engineer"],
     ["engineering", "cad", "autocad", "solidworks", "manufacturing", "plant", "machinery", "commissioning"]),
    ("Healthcare & Medicine",
     ["doctor", "physician", "nurse", "pharmacist", "medical officer", "surgeon", "anaesthetist",
      "physiotherapist", "radiographer", "lab technician", "clinical", "healthcare manager",
      "occupational therapist", "dentist", "midwife"],
     ["hospital", "clinic", "patient", "medical", "health", "pharmaceutical", "diagnosis", "treatment"]),
    ("Education & Training",
     ["teacher", "lecturer", "professor", "trainer", "educator", "tutor", "school principal",
      "academic", "curriculum", "e-learning", "instructional designer", "teaching assistant"],
     ["school", "university", "college", "classroom", "students", "pedagogy", "curriculum", "education"]),
    ("Hospitality & Tourism",
     ["hotel manager", "front desk", "housekeeping", "chef", "sous chef", "food and beverage",
      "f&b manager", "restaurant manager", "bartender", "waiter", "concierge", "tour guide",
      "travel agent", "events coordinator", "catering"],
     ["hospitality", "hotel", "resort", "tourism", "guest", "accommodation", "restaurant", "kitchen"]),
    ("Logistics & Supply Chain",
     ["supply chain manager", "logistics coordinator", "warehouse manager", "fleet manager",
      "procurement manager", "purchasing manager", "import export", "freight", "shipping coordinator",
      "inventory manager", "demand planner"],
     ["logistics", "supply chain", "warehouse", "inventory", "freight", "procurement", "sourcing"]),
    ("Legal",
     ["lawyer", "attorney", "legal counsel", "paralegal", "compliance officer", "legal advisor",
      "solicitor", "barrister", "corporate counsel", "legal manager", "contract manager"],
     ["legal", "law", "contracts", "litigation", "regulatory", "compliance", "gdpr"]),
    ("Administration & Operations",
     ["office manager", "executive assistant", "administrative officer", "operations manager",
      "pa", "personal assistant", "receptionist", "data entry", "office administrator",
      "company secretary", "business analyst"],
     ["administration", "operations", "office", "coordination", "scheduling", "reporting", "clerical"]),
    ("Customer Service",
     ["customer service", "call centre", "customer success", "customer support", "help desk",
      "service advisor", "client relations", "customer experience", "contact centre"],
     ["customer", "support", "helpdesk", "tickets", "escalation", "satisfaction", "service level"]),
    ("Construction & Real Estate",
     ["quantity surveyor", "site supervisor", "project manager construction", "architect",
      "draughtsman", "property manager", "estate agent", "real estate", "building inspector",
      "land surveyor", "construction manager"],
     ["construction", "building", "property", "real estate", "site", "contractor", "tender"]),
    ("Manufacturing & Production",
     ["production manager", "quality control", "quality assurance", "qa", "qc", "factory manager",
      "plant manager", "production supervisor", "assembly", "cnc operator", "technician"],
     ["production", "manufacturing", "factory", "assembly", "quality", "lean", "six sigma"]),
    ("Design & Creative",
     ["graphic designer", "ui/ux", "product designer", "art director", "creative director",
      "animator", "illustrator", "photographer", "videographer", "motion designer", "web designer"],
     ["design", "creative", "adobe", "figma", "photoshop", "illustrator", "indesign", "sketch", "branding"]),
    ("Research & Science",
     ["research scientist", "data scientist", "lab researcher", "research analyst",
      "clinical researcher", "environmental scientist", "chemist", "biologist", "statistician"],
     ["research", "analysis", "data", "laboratory", "science", "experiment", "findings", "methodology"]),
    ("Security",
     ["security officer", "security guard", "security manager", "cctv", "loss prevention",
      "risk manager", "health and safety", "hse officer", "osh", "fire safety"],
     ["security", "safety", "risk", "surveillance", "patrol", "access control", "emergency"]),
    ("Media & Journalism",
     ["journalist", "editor", "reporter", "broadcast", "news anchor", "content creator",
      "media manager", "radio", "television", "producer", "scriptwriter"],
     ["media", "journalism", "broadcast", "news", "editorial", "publishing", "press"]),
    ("Non-Profit & Social Work",
     ["social worker", "ngo", "charity", "programme coordinator", "community development",
      "welfare officer", "case manager", "development officer", "fundraiser", "volunteer coordinator"],
     ["social", "ngo", "community", "welfare", "beneficiary", "donor", "impact", "charity"]),
]

# Procurement / notice markers in a title. Conservative: clear acronyms as whole
# tokens plus unambiguous phrases. Deliberately does NOT match a bare "call for
# applications" (often a volunteer/role advert, not a tender).
_TENDER_TITLE_RE = re.compile(
    r"\b(?:rfq|rfp|reoi|eoi|itb|itt|spn|rfb|rfa|gpn|ifb|rfi)\b"
    r"|invitation\s+to\s+(?:bid|tender)|invitation\s+for\s+bids?"
    r"|request\s+for\s+(?:quotation|proposal|proposals|expression|expressions|bids?)"
    r"|expressions?\s+of\s+interest"
    r"|\btenders?\b|procurement\s+notice|specific\s+procurement|general\s+procurement"
    r"|call\s+for\s+(?:bid|bids|tender|tenders|proposal|proposals|expression|expressions|quotation)"
    r"|matching\s+grant|terms\s+of\s+reference|prior\s+notice\s+of\s+procurement",
    re.I,
)
TENDER_FIELD = "Public Notices & Tenders"

def infer_field(title: str, description: str, fallback_categories: str = "") -> str:
    """
    Resolve a single canonical job field from the title + description. Procurement
    notices (common in the EthioJobs feed) are detected first and routed to
    "Public Notices & Tenders" — otherwise incidental keyword hits mislabel them
    (e.g. the word "tender" lives in Construction's weak list, and an SPN for a
    "Digital Marketing Campaign" would land in Marketing). After that, strong
    keywords win over weak (list order = tie-break). If nothing matches, fall back
    to the site's own category so the field is never empty.
    """
    title_l = (title or "").lower()
    if _TENDER_TITLE_RE.search(title_l):
        return TENDER_FIELD

    text = f"{title}\n{description}".lower()
    for field, strong, _weak in FIELD_KEYWORD_MAP:
        if _kw_hit(text, strong):
            return field
    for field, _strong, weak in FIELD_KEYWORD_MAP:
        if _kw_hit(text, weak):
            return field
    if fallback_categories:
        cats = [c.strip() for c in fallback_categories.split(",") if c.strip()]
        # Prefer the site's tender/notice category if it tagged one.
        for c in cats:
            if "tender" in c.lower() or "notice" in c.lower():
                return TENDER_FIELD
        if cats:
            return cats[0]
    return ""

# =============================================================================
#  NLP TOOLS (lazy init, optional)
# =============================================================================

_grammar_tool = None
_sim_model    = None

def _get_grammar_tool():
    global _grammar_tool
    if _grammar_tool is None and _NLP_AVAILABLE:
        try:
            _grammar_tool = language_tool_python.LanguageTool(
                "en-US", remote_server="https://api.languagetool.org")
        except Exception as e:
            log_.warning(f"LanguageTool init failed: {e}")
    return _grammar_tool

def _get_sim_model():
    global _sim_model
    if _sim_model is None and _NLP_AVAILABLE:
        try:
            _sim_model = SentenceTransformer("all-MiniLM-L6-v2", device="cpu")
        except Exception as e:
            log_.warning(f"SentenceTransformer init failed: {e}")
    return _sim_model

def grammar_correct(text: str) -> str:
    tool = _get_grammar_tool()
    if tool:
        try:
            return language_tool_python.utils.correct(text, tool.check(text))
        except Exception:
            pass
    return text

def similarity_score(a: str, b: str) -> float:
    model = _get_sim_model()
    if model:
        try:
            emb = model.encode([a, b], convert_to_tensor=True)
            return float(st_util.pytorch_cos_sim(emb[0], emb[1]))
        except Exception:
            pass
    def tokens(s):
        return set(re.sub(r"[^a-z0-9 ]", " ", s.lower()).split())
    ta, tb = tokens(a), tokens(b)
    if not ta or not tb:
        return 0.0
    return len(ta & tb) / max(len(ta), len(tb))

def clean_output(text: str) -> str:
    text = _fix_mojibake(text)
    for pat in [r"\[/?INST\]", r"</?s>",
                r"(?i)(rewritten?|rephrased?|output|paraphrase[d]?)[:\s]+",
                r"\*\*", r"###", r"---"]:
        text = re.sub(pat, "", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return grammar_correct(text.strip())

# =============================================================================
#  MISTRAL API
# =============================================================================

def mistral_generate(prompt: str, max_tokens: int = 400, temperature: float = 0.7) -> str:
    if not MISTRAL_API_KEY:
        log_.warning("MISTRAL_API_KEY not set — skipping paraphrase")
        return ""
    try:
        response = requests.post(
            MISTRAL_URL,
            headers={
                "Authorization": f"Bearer {MISTRAL_API_KEY}",
                "Content-Type": "application/json",
            },
            json={
                "model": MISTRAL_MODEL,
                "messages": [{"role": "user", "content": prompt}],
                "max_tokens": max_tokens,
                "temperature": temperature,
            },
            timeout=30,
        )
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"].strip()
    except Exception as e:
        log_.error(f"Mistral API error: {e}")
        return ""

# =============================================================================
#  PARAPHRASE FUNCTIONS
# =============================================================================

def _print_wrapped(text: str, prefix: str = "   ", width: int = 100):
    words = text.split()
    line  = []
    for w in words:
        line.append(w)
        if len(" ".join(line)) >= width:
            print(f"{prefix}{' '.join(line)}")
            line = []
    if line:
        print(f"{prefix}{' '.join(line)}")

def paraphrase_title(title: str) -> str:
    if not ENABLE_PARAPHRASE:
        return title
    clean = sanitize_text(title)
    if not clean:
        return title

    print(f"\n ┌─ TITLE PARAPHRASE {'─'*45}")
    print(f" │ Original : \"{clean}\"")
    print(f" │ {'─'*60}")

    best_result = None
    best_sim    = 0.0

    for attempt in range(4):
        temp = round(0.68 + attempt * 0.06, 2)
        print(f" │ Attempt {attempt+1} (temp={temp}):")

        prompt = (
            f"Rewrite this job title professionally using different words. "
            f"Output ONLY the rewritten title, nothing else. "
            f"Keep it between 4 and 12 words.\n\nJob title: {clean}"
        )

        raw    = mistral_generate(prompt, max_tokens=50, temperature=temp)
        result = clean_output(raw).split("\n")[0].strip().strip('"').strip("'")

        wc     = len(result.split()) if result else 0
        sim    = similarity_score(clean, result) if result else 0.0
        is_dup = result.lower().strip() == clean.lower().strip()

        print(f" │    Output  : \"{result}\"")
        print(f" │    Words   : {wc} | Similarity: {sim:.3f} | Duplicate: {'Yes' if is_dup else 'No'}")

        valid = bool(result) and 4 <= wc <= 14 and sim >= 0.55 and not is_dup

        if not valid:
            reasons = []
            if not result:  reasons.append("empty output")
            if wc < 4:      reasons.append(f"too short ({wc} words, min=4)")
            if wc > 14:     reasons.append(f"too long ({wc} words, max=14)")
            if sim < 0.55:  reasons.append(f"sim={sim:.3f} < 0.55")
            if is_dup:      reasons.append("identical to original")
            print(f" │    -> REJECTED — {', '.join(reasons)}")
        else:
            if sim > best_sim:
                best_sim    = sim
                best_result = result
                print(f" │    -> ACCEPTED — new best candidate (sim={sim:.3f})")
            else:
                print(f" │    -> VALID but not better than current best (best sim={best_sim:.3f})")

        print(f" │ {'─'*60}")
        time.sleep(1)

    if best_result:
        print(f" │ FINAL SELECTED : \"{best_result}\"")
        print(f" │    Similarity  : {best_sim:.3f}")
        print(f" └{'─'*65}")
        return best_result
    else:
        print(f" │ No valid paraphrase found -> Keeping original: \"{clean}\"")
        print(f" └{'─'*65}")
        return clean

def paraphrase_description(text: str) -> str:
    if not ENABLE_PARAPHRASE:
        return text
    clean = sanitize_text(text)
    if not clean:
        return text

    paragraphs  = [p.strip() for p in re.split(r"\n+", clean) if p.strip()]
    if not paragraphs:
        paragraphs = [clean]
    rewritten   = []
    success_count = 0

    print(f"\n ┌─ DESCRIPTION PARAPHRASE ({len(paragraphs)} paragraph(s)) {'─'*15}")

    for i, para in enumerate(paragraphs):
        orig_wc = len(para.split())

        print(f"\n │ ┌─ Paragraph {i+1}/{len(paragraphs)} {'─'*50}")
        print(f" │ │ ORIGINAL ({orig_wc} words):")
        _print_wrapped(para, prefix=" │ │    ")
        print(f" │ │ {'─'*60}")

        # Very short fragments (section labels, single bullets) — keep as-is.
        if orig_wc < 8:
            print(f" │ │ (kept — too short to paraphrase safely)")
            rewritten.append(para)
            print(f" │ └{'─'*62}")
            continue

        prompt = (
            f"Rewrite this job description paragraph professionally. "
            f"Keep ALL facts, requirements, and responsibilities. "
            f"Use different sentence structure and vocabulary. "
            f"Output ONLY the rewritten paragraph — no labels, no explanation.\n\n"
            f"Original:\n{para}"
        )

        best_result = None
        best_sim    = 0.0
        accepted_text = None

        for attempt in range(3):
            temp = round(0.65 + attempt * 0.08, 2)
            print(f" │ │ Attempt {attempt+1}/3 (temp={temp}):")

            raw    = mistral_generate(prompt, max_tokens=500, temperature=temp)
            result = clean_output(raw).strip()

            rw  = len(result.split()) if result else 0
            sim = similarity_score(para, result) if result and rw >= 5 else 0.0

            if result:
                print(f" │ │    Paraphrased ({rw} words, sim={sim:.3f}):")
                _print_wrapped(result, prefix=" │ │       ")
            else:
                print(f" │ │    Paraphrased : (no output from model)")

            valid = bool(result) and rw >= 8 and sim >= 0.48

            if not valid:
                reasons = []
                if not result: reasons.append("empty output")
                if rw < 8:     reasons.append(f"too short ({rw} words, min=8)")
                if sim < 0.48: reasons.append(f"sim={sim:.3f} < 0.48")
                print(f" │ │    -> REJECTED — {', '.join(reasons)}")
                if result and sim > best_sim:
                    best_sim    = sim
                    best_result = result
                    print(f" │ │       (stored as best fallback, sim={sim:.3f})")
            else:
                print(f" │ │    -> ACCEPTED on attempt {attempt+1}")
                rewritten.append(result)
                success_count += 1
                accepted_text = result
                break

            print(f" │ │ {'─'*60}")
            time.sleep(1)

        if accepted_text is None:
            print(f" │ │ {'─'*60}")
            if best_result and best_sim >= 0.40:
                print(f" │ │ FALLBACK — Using best attempt (sim={best_sim:.3f}):")
                _print_wrapped(best_result, prefix=" │ │    ")
                rewritten.append(best_result)
                success_count += 1
            else:
                print(f" │ │ KEPT ORIGINAL — no acceptable paraphrase (best sim={best_sim:.3f})")
                rewritten.append(para)

        print(f" │ └{'─'*62}")

    print(f"\n │ SUMMARY: {success_count}/{len(paragraphs)} paragraphs successfully paraphrased")
    print(f" └{'─'*80}\n")

    return "\n\n".join(rewritten)

# =============================================================================
#  DUPLICATE TRACKER — pure stdlib csv, NO pandas dependency
# =============================================================================

def _init_tracker():
    if not os.path.exists(PROCESSED_IDS_FILE):
        try:
            with open(PROCESSED_IDS_FILE, "w", newline="", encoding="utf-8") as f:
                csv.writer(f).writerow(_TRACKER_FIELDS)
            log_.info(f"Tracker file created: {PROCESSED_IDS_FILE}")
        except Exception as e:
            log_.error(f"Could not create tracker file {PROCESSED_IDS_FILE}: {e}")

def load_processed_ids() -> tuple:
    _init_tracker()
    ids, urls = set(), set()
    try:
        with open(PROCESSED_IDS_FILE, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                if row.get("Job ID"):
                    ids.add(row["Job ID"].strip())
                if row.get("Job URL"):
                    urls.add(row["Job URL"].strip())
    except Exception as e:
        log_.error(f"Could not read tracker file: {e}")
    return ids, urls

def _upsert_row(job_id: str, updates: dict):
    _init_tracker()
    rows = []
    try:
        with open(PROCESSED_IDS_FILE, newline="", encoding="utf-8") as f:
            rows = list(csv.DictReader(f))
    except Exception as e:
        log_.error(f"Tracker read error: {e}")
        rows = []

    found = False
    for row in rows:
        if row.get("Job ID", "").strip() == str(job_id):
            row.update(updates)
            row["Timestamp"] = datetime.now().isoformat()
            found = True
            break

    if not found:
        new_row = {k: "" for k in _TRACKER_FIELDS}
        new_row["Job ID"]    = str(job_id)
        new_row["Timestamp"] = datetime.now().isoformat()
        new_row.update(updates)
        rows.append(new_row)

    try:
        with open(PROCESSED_IDS_FILE, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=_TRACKER_FIELDS, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(rows)
    except Exception as e:
        log_.error(f"Tracker write error: {e}")

def make_job_id(job_url: str, title: str = "", company: str = "") -> str:
    if job_url:
        return hashlib.md5(job_url.encode()).hexdigest()[:16]
    seed = f"{title}{company}"
    return hashlib.md5(seed.encode()).hexdigest()[:16]

def mark_scraped(job_id, job_url, title, company):
    log_.info(f"Tracker -> scraped: {job_id} | {title}")
    _upsert_row(job_id, {
        "Job URL":      job_url,
        "Job Title":    title,
        "Company Name": company,
        "Status":       "scraped",
        "WP ID":        "",
    })

def mark_paraphrased(job_id):
    _upsert_row(job_id, {"Status": "paraphrased"})

def mark_posted(job_id, wp_id, wp_url):
    _upsert_row(job_id, {"Status": "posted", "WP ID": str(wp_id)})

def mark_failed(job_id, reason):
    _upsert_row(job_id, {"Status": f"failed|{reason}"})

# =============================================================================
#  FLAGGED CSV (non-qualifying / login-only apply)
# =============================================================================

def _init_flagged():
    if not os.path.exists(FLAGGED_FILE):
        try:
            with open(FLAGGED_FILE, "w", newline="", encoding="utf-8") as f:
                csv.writer(f).writerow(_FLAGGED_FIELDS)
        except Exception as e:
            log_.error(f"Could not create flagged file {FLAGGED_FILE}: {e}")

def write_flagged(raw_job: dict, reason: str, apply_note: str):
    _init_flagged()
    try:
        with open(FLAGGED_FILE, "a", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow([
                "EthioJobs",
                raw_job.get("title", ""),
                raw_job.get("company_name", ""),
                raw_job.get("location", ""),
                raw_job.get("salary", ""),
                raw_job.get("deadline", ""),
                reason,
                apply_note,
                raw_job.get("job_url", ""),
                datetime.now().isoformat(),
            ])
    except Exception as e:
        log_.error(f"Flagged write error: {e}")

# =============================================================================
#  WORDPRESS POSTING
# =============================================================================

def _wp_auth_headers() -> dict:
    token = base64.b64encode(f"{WP_USER}:{WP_PASSWORD}".encode()).decode()
    return {"Authorization": f"Basic {token}", "Content-Type": "application/json"}

def get_or_create_term(taxonomy_url: str, name: str):
    if not name or not name.strip():
        return None
    slug = re.sub(r"[^a-z0-9-]", "-", name.lower().strip())
    h = _wp_auth_headers()
    try:
        r = requests.get(f"{taxonomy_url}?slug={slug}", headers=h, timeout=10, verify=False)
        terms = r.json()
        if isinstance(terms, list) and terms:
            return terms[0]["id"]
    except Exception:
        pass
    try:
        r = requests.post(taxonomy_url, json={"name": name, "slug": slug},
                          headers=h, auth=(WP_USER, WP_PASSWORD), timeout=10, verify=False)
        return r.json().get("id")
    except Exception as e:
        log_.error(f"Term create error '{name}': {e}")
        return None

def post_job_to_wordpress(job: dict) -> tuple:
    if not WP_USER or not WP_PASSWORD:
        log_.warning("WP_USERNAME / WP_APP_PASSWORD not set — skipping WordPress post")
        return None, None

    h = _wp_auth_headers()

    title       = sanitize_text(job.get("jobTitle", ""))
    description = sanitize_text(job.get("jobDescription", ""))
    if not title or not description:
        return None, None

    slug = re.sub(r"[^a-z0-9-]", "-", title.lower())[:80]
    try:
        r = requests.get(f"{WP_JOBS_URL}?slug={slug}", headers=h, timeout=10, verify=False)
        posts = r.json()
        if isinstance(posts, list) and posts:
            log_.info(f"Job already on WP: {title}")
            return posts[0]["id"], posts[0].get("link")
    except Exception:
        pass

    logo_url    = sanitize_text(job.get("companyLogo", ""), is_url=True)
    location    = sanitize_text(job.get("jobLocation", ""))
    raw_type    = sanitize_text(job.get("jobType", "")) or "Full-time"
    job_type_s  = JOB_TYPE_MAPPING.get(raw_type.lower().strip(), "full-time")
    company     = sanitize_text(job.get("companyName", ""))
    application = sanitize_text(job.get("application", ""), is_url=True)
    company_url = sanitize_text(job.get("companyUrl", ""), is_url=True)
    deadline    = sanitize_text(job.get("deadline", ""))
    co_website  = sanitize_text(job.get("companyWebsite", ""), is_url=True)
    qualif      = sanitize_text(job.get("jobQualifications", ""))
    experience  = sanitize_text(job.get("jobExperience", ""))
    co_address  = sanitize_text(job.get("companyAddress", ""))
    job_field   = sanitize_text(job.get("jobField", ""))
    salary      = sanitize_text(job.get("salaryRange", ""))
    about       = sanitize_text(job.get("companyDetails", ""))

    is_email = bool(re.match(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", application))
    is_url_v = bool(re.match(r"^https?://[^\s]+$", application))
    if not (is_email or is_url_v):
        application = ""

    # Upload logo
    attachment_id = None
    if logo_url:
        try:
            img_r = requests.get(logo_url, headers={"User-Agent": HEADERS["User-Agent"]}, timeout=15)
            if img_r.status_code == 200:
                ct  = img_r.headers.get("Content-Type", "image/jpeg")
                ext = "png" if "png" in ct else "jpg"
                fn  = re.sub(r"[^a-z0-9]", "-", company.lower()) + "-logo." + ext
                up_h = dict(_wp_auth_headers())
                up_h["Content-Disposition"] = f"attachment; filename={fn}"
                up_h["Content-Type"] = ct
                up_r = requests.post(WP_MEDIA_URL, headers=up_h, data=img_r.content,
                                     auth=(WP_USER, WP_PASSWORD), timeout=20, verify=False)
                if up_r.status_code in (200, 201):
                    attachment_id = up_r.json().get("id")
        except Exception as e:
            log_.warning(f"Logo upload failed: {e}")

    region_term_id   = get_or_create_term(f"{WP_BASE}/job_listing_region", location)
    job_type_term_id = get_or_create_term(f"{WP_BASE}/job_listing_type",
                                           job_type_s.replace("-", " ").title())

    payload = {
        "title":          title,
        "content":        description,
        "status":         "publish",
        "featured_media": attachment_id or 0,
        "meta": {
            "_job_title":          title,
            "_job_location":       location,
            "_job_type":           job_type_s,
            "_job_description":    description,
            "_application":        application,
            "_company_url":        company_url,
            "_job_expires":        deadline,
            "_company_name":       company,
            "_company_website":    co_website,
            "_company_logo":       str(attachment_id) if attachment_id else "",
            "_company_address":    co_address,
            "_company_details":    about,
            "_job_qualifications": qualif,
            "_job_experiences":    experience,
            "_job_field":          job_field,
            "_job_salary":         salary,
        },
    }
    if region_term_id:   payload["job_listing_region"] = [region_term_id]
    if job_type_term_id: payload["job_listing_type"]   = [job_type_term_id]

    for attempt in range(3):
        try:
            r = requests.post(WP_JOBS_URL, json=payload, headers=h,
                              auth=(WP_USER, WP_PASSWORD), timeout=20, verify=False)
            r.raise_for_status()
            post = r.json()
            log_.info(f"Job posted: '{title}' -> WP ID {post.get('id')}")
            return post.get("id"), post.get("link")
        except Exception as e:
            log_.error(f"Job post attempt {attempt+1} failed: {e}")
            if attempt < 2:
                time.sleep(2 ** attempt)
    return None, None

# =============================================================================
#  ETHIOJOBS JSON-API SCRAPE LAYER  (pure requests — no browser)
# =============================================================================
#
#  ethiojobs.net is a Next.js SPA whose HTML is just a loading screen, so there
#  is nothing for BeautifulSoup to read on the page body. All job data is served
#  as JSON by the site's backend (api.ethiojobs.net), which the browser calls
#  client-side. We call that same JSON API directly with `requests`.
#
#  Because the exact backend routes are not published, the scraper discovers them
#  automatically on first run (discover_api): it downloads the site's own
#  JavaScript bundles and reads the API base URL + route fragments straight out
#  of them, then probes the candidate list/detail endpoints and caches the
#  working ones to API_CACHE_FILE. Subsequent runs read the cache. You can also
#  hard-set everything via the ETHIOJOBS_API_* env vars, or run with --discover
#  to print what was found.
#
#  FLOW
#  ----
#    discover_api()                  -> {"base","list","detail"} (cached)
#    collect_job_records(list_ep)    -> list of raw job dicts (paged JSON)
#    scrape_job_detail(rec_or_url)   -> normalised raw_job dict via _map_job_record
# =============================================================================

# A browser-like API session: many JSON backends behind a SPA check Origin/
# Referer and an Accept: application/json header.
API_SESSION = requests.Session()
API_SESSION.headers.update({
    "User-Agent": HEADERS["User-Agent"],
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
    "Origin": BASE_URL,
    "Referer": BASE_URL + "/",
})

# EthioJobs career-level label -> our standard experience band.
_CAREER_LEVEL_MAP = {
    "junior level(1-3 years)":  "1 - 2 Years",
    "junior level (1-3 years)": "1 - 2 Years",
    "junior":                   "1 - 2 Years",
    "entry level":              "1 - 2 Years",
    "entry-level":              "1 - 2 Years",
    "mid level(3-5 years)":     "3 - 5 Years",
    "mid level (3-5 years)":    "3 - 5 Years",
    "mid-level":                "3 - 5 Years",
    "mid level":                "3 - 5 Years",
    "senior(5-8 years)":        "6 - 10 Years",
    "senior (5-8 years)":       "6 - 10 Years",
    "senior level":             "6 - 10 Years",
    "senior":                   "6 - 10 Years",
    "managerial level":         "10+ Years",
    "managerial":               "10+ Years",
    "director":                 "10+ Years",
    "executive":                "10+ Years",
}

def _career_level_to_band(raw: str) -> str:
    """Convert an EthioJobs career-level string to our standard experience band."""
    if not raw:
        return ""
    key = raw.strip().lower()
    if key in _CAREER_LEVEL_MAP:
        return _CAREER_LEVEL_MAP[key]
    for label, band in _CAREER_LEVEL_MAP.items():
        if key.startswith(label):
            return band
    if "junior" in key or "entry" in key:                 return "1 - 2 Years"
    if "mid" in key:                                      return "3 - 5 Years"
    if "senior" in key:                                   return "6 - 10 Years"
    if "managerial" in key or "director" in key or "executive" in key:
        return "10+ Years"
    return ""

# Candidate API field names per logical field, tried in order. Covers snake_case
# and camelCase so minor backend renames don't silently drop data.
_JOB_KEY_CANDIDATES = {
    "id":           ["id", "_id", "jobId", "job_id", "slug", "token", "uuid"],
    "title":        ["title", "jobTitle", "job_title", "name", "position", "positionTitle"],
    "company":      ["company", "companyName", "company_name", "employer", "organization",
                     "organisation", "recruiter", "employerName"],
    "company_url":  ["companyUrl", "company_url", "companyLink", "profileUrl", "company_profile"],
    "company_logo": ["companyLogo", "company_logo", "logo", "logoUrl", "logo_url", "image", "companyImage"],
    "company_website": ["companyWebsite", "company_website", "website", "websiteUrl"],
    "company_address": ["companyAddress", "company_address", "address", "officeAddress"],
    "company_about":   ["companyAbout", "company_about", "companyDescription",
                        "company_description", "aboutCompany", "about"],
    "location":     ["location", "jobLocation", "job_location", "region", "city",
                     "workLocation", "work_location", "dutyStation", "duty_station"],
    "job_type":     ["jobType", "job_type", "employmentType", "employment_type", "type",
                     "contractType", "contract_type"],
    "category":     ["category", "categories", "jobCategory", "job_category", "sector",
                     "field", "department", "industry"],
    "career_level": ["careerLevel", "career_level", "level", "experienceLevel",
                     "experience_level", "seniority"],
    "deadline":     ["deadline", "closingDate", "closing_date", "applicationDeadline",
                     "application_deadline", "expiryDate", "expiry_date", "expiresAt",
                     "expires_at", "dueDate", "due_date", "endDate"],
    "date_posted":  ["createdAt", "created_at", "datePosted", "date_posted", "postedAt",
                     "posted_at", "publishedAt", "published_at", "startDate"],
    "description":  ["description", "jobDescription", "job_description", "details",
                     "jobDetails", "job_details", "body", "content", "summary"],
    "requirements": ["requirements", "jobRequirements", "job_requirements", "qualifications",
                     "jobQualifications", "job_qualifications", "requirement"],
    "how_to_apply": ["howToApply", "how_to_apply", "applicationInstructions",
                     "application_instructions", "applyInstructions", "apply_instructions",
                     "howToApplyText"],
    "salary":       ["salary", "salaryRange", "salary_range", "compensation", "pay",
                     "salaryInfo", "salary_info", "salaryScale"],
    "apply_email":  ["applicationEmail", "application_email", "applyEmail", "apply_email",
                     "contactEmail", "contact_email", "email"],
    "apply_url":    ["applicationUrl", "application_url", "applyUrl", "apply_url",
                     "applicationLink", "application_link", "applyLink", "apply_link",
                     "externalUrl", "external_url"],
    "num_vacancies":["numberOfVacancies", "number_of_vacancies", "vacancies",
                     "numberOfPositions", "positions", "noOfPositions"],
    "slug":         ["slug", "seoUrl", "url", "link", "jobUrl", "job_url"],
}

def _pick(data: dict, field: str, default=""):
    """Return the first non-empty candidate value for a logical field."""
    for k in _JOB_KEY_CANDIDATES.get(field, []):
        if k in data:
            v = data.get(k)
            if isinstance(v, dict):
                # nested object e.g. {"name": "..."} or {"title": "..."}
                v = v.get("name") or v.get("title") or v.get("value") or ""
            if isinstance(v, list):
                v = ", ".join(str(x.get("name", x) if isinstance(x, dict) else x) for x in v)
            if v is not None and str(v).strip():
                return str(v).strip()
    return default

_JOB_LIKE_KEYS = ("jobtitle", "job_title", "title", "position", "positiontitle")

def _is_job_api_response(data) -> bool:
    """Heuristic: is this JSON a job record, or a wrapper holding a list of jobs?"""
    if isinstance(data, dict):
        keys_low = {k.lower() for k in data}
        if any(k in keys_low for k in _JOB_LIKE_KEYS):
            return True
        for v in data.values():
            if isinstance(v, list) and v and isinstance(v[0], dict):
                sub = {kk.lower() for kk in v[0]}
                if any(k in sub for k in _JOB_LIKE_KEYS):
                    return True
            if isinstance(v, dict):
                for vv in v.values():
                    if isinstance(vv, list) and vv and isinstance(vv[0], dict):
                        sub = {kk.lower() for kk in vv[0]}
                        if any(k in sub for k in _JOB_LIKE_KEYS):
                            return True
    if isinstance(data, list) and data and isinstance(data[0], dict):
        sub = {k.lower() for k in data[0]}
        if any(k in sub for k in _JOB_LIKE_KEYS):
            return True
    return False

def _extract_job_list_from_api(data) -> list:
    """Given an API response blob, return a list of job-record dicts."""
    if isinstance(data, list):
        return [d for d in data if isinstance(d, dict)]
    if isinstance(data, dict):
        for wrap in ("data", "jobs", "results", "items", "listings", "vacancies",
                     "posts", "content", "docs", "rows", "records", "hits"):
            v = data.get(wrap)
            if isinstance(v, list) and v and isinstance(v[0], dict):
                return v
            # one level deeper: {"data": {"jobs": [...]}}
            if isinstance(v, dict):
                for wrap2 in ("jobs", "results", "items", "content", "docs", "data"):
                    vv = v.get(wrap2)
                    if isinstance(vv, list) and vv and isinstance(vv[0], dict):
                        return vv
        if _is_job_api_response(data) and any(
                k.lower() in _JOB_LIKE_KEYS for k in data):
            return [data]
    return []

def _norm_ethiojobs_url(href: str) -> str:
    """Canonicalise a detail URL to https://ethiojobs.net/job/<token>-<slug>."""
    if not href:
        return ""
    if href.startswith("http"):
        p = urlsplit(href)
        return urlunsplit(("https", "ethiojobs.net", p.path.rstrip("/"), "", ""))
    return urlunsplit(("https", "ethiojobs.net", "/" + href.lstrip("/").rstrip("/"), "", ""))

def _job_url_from_record(rec: dict) -> str:
    """Build the public ethiojobs.net/job/<...> URL for a record, for dedup + apply."""
    slug = _pick(rec, "slug")
    if slug:
        if slug.startswith("http") or slug.startswith("/job"):
            return _norm_ethiojobs_url(slug)
        if "/" not in slug:
            return _norm_ethiojobs_url("/job/" + slug)
        return _norm_ethiojobs_url(slug)
    jid = _pick(rec, "id")
    if jid:
        return _norm_ethiojobs_url("/job/" + jid)
    return ""

# ---------------------------------------------------------------------------
# API DISCOVERY — read the site's JS bundles to find base + routes, then probe
# ---------------------------------------------------------------------------

_JS_SRC_RE   = re.compile(r'(?:src|href)="(/_next/[^"]+\.js)"', re.I)
_API_BASE_RE = re.compile(r'https://[a-z0-9.-]*api[a-z0-9.-]*\.ethiojobs\.net', re.I)
_API_PATH_RE = re.compile(r'["\'`](/(?:api/)?(?:job|jobs|vacancy|vacancies)[a-z0-9/_\-{}.]*)["\'`]', re.I)

def _load_api_cache() -> dict:
    try:
        with open(API_CACHE_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def _save_api_cache(d: dict):
    try:
        with open(API_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(d, f, indent=2)
    except Exception as e:
        log_.warning(f"Could not write API cache: {e}")

def _scan_js_for_api(seed_url: str):
    """Download the SPA's JS bundles and pull out the API base + route fragments."""
    bases, paths = set(), set()
    try:
        r = SESSION.get(seed_url, timeout=REQUEST_TIMEOUT)
        html = r.text
    except Exception as e:
        log(C_DIM(f"  Could not load {seed_url} for discovery: {e}"))
        return bases, paths

    for m in _API_BASE_RE.findall(html):
        bases.add(m.rstrip("/"))

    js_urls = []
    for src in _JS_SRC_RE.findall(html):
        js_urls.append(urljoin(BASE_URL + "/", src))
    # De-dup, cap the number of bundles we fetch.
    seen = set()
    for ju in js_urls:
        if ju in seen:
            continue
        seen.add(ju)
        if len(seen) > 25:
            break
        try:
            jr = SESSION.get(ju, timeout=REQUEST_TIMEOUT)
            js = jr.text
        except Exception:
            continue
        for m in _API_BASE_RE.findall(js):
            bases.add(m.rstrip("/"))
        for m in _API_PATH_RE.findall(js):
            paths.add(m)
    return bases, paths

def _probe_list_endpoint(base: str, path: str) -> bool:
    """True if GET base+path (with common paging params) returns job JSON."""
    url = base.rstrip("/") + "/" + path.lstrip("/")
    for params in ({"page": 1, "limit": PAGE_SIZE},
                   {"page": 1, "pageSize": PAGE_SIZE},
                   {"offset": 0, "limit": PAGE_SIZE},
                   {}):
        try:
            r = API_SESSION.get(url, params=params, timeout=REQUEST_TIMEOUT)
            if r.status_code != 200 or "json" not in r.headers.get("content-type", ""):
                continue
            data = r.json()
        except Exception:
            continue
        if _extract_job_list_from_api(data):
            return True
    return False

def discover_api(force: bool = False) -> dict:
    """
    Resolve {"base","list","detail"} for the EthioJobs JSON API. Order:
      1) cached values (unless force)
      2) explicit env overrides (ETHIOJOBS_API_BASE + first probed list path)
      3) bases/paths scraped from the site's JS bundles, then probed
      4) the built-in candidate lists, probed
    The working combination is cached to API_CACHE_FILE.
    """
    if not force:
        cached = _load_api_cache()
        if cached.get("base") and cached.get("list"):
            return cached

    log(C_BLUE("\n  Discovering EthioJobs API endpoints …"))

    # Gather candidate bases.
    bases = []
    if API_BASE:
        bases.append(API_BASE.rstrip("/"))
    scanned_bases, scanned_paths = _scan_js_for_api(JOBS_URL)
    for b in scanned_bases:
        if b not in bases:
            bases.append(b)
    # Fallbacks if JS scan found nothing.
    for b in ("https://api.ethiojobs.net", "https://ethiojobs.net/api"):
        if b not in bases:
            bases.append(b)

    # Candidate list paths: scanned first, then configured.
    list_candidates = []
    for p in scanned_paths:
        if "{" not in p and p not in list_candidates:
            list_candidates.append(p)
    for p in API_LIST_PATHS:
        if p not in list_candidates:
            list_candidates.append(p)

    found = {}
    for base in bases:
        for path in list_candidates:
            if _probe_list_endpoint(base, path):
                found["base"] = base
                found["list"] = path
                log(C_GREEN(f"  ✓ list endpoint: {base}{path}"))
                break
        if found:
            break

    if not found:
        log(C_RED("  Could not auto-discover a working list endpoint."))
        log(C_DIM("  Set ETHIOJOBS_API_BASE / ETHIOJOBS_API_LIST_PATHS, or run --discover"))
        log(C_DIM(f"  JS-scanned bases: {sorted(scanned_bases) or '—'}"))
        log(C_DIM(f"  JS-scanned paths: {sorted(scanned_paths) or '—'}"))
        return {}

    # Detail endpoint: keep the configured candidates (filled per-job at runtime).
    found["detail"] = API_DETAIL_PATHS
    _save_api_cache(found)
    return found

# ---------------------------------------------------------------------------
# STEP 1 — collect job records from the listing API (paged)
# ---------------------------------------------------------------------------

def _fetch_list_page(base: str, path: str, page: int) -> list:
    """Fetch one listing page; return its job-record dicts (possibly empty)."""
    url = base.rstrip("/") + "/" + path.lstrip("/")
    for params in ({"page": page, "limit": PAGE_SIZE},
                   {"page": page, "pageSize": PAGE_SIZE},
                   {"offset": (page - 1) * PAGE_SIZE, "limit": PAGE_SIZE}):
        try:
            r = API_SESSION.get(url, params=params, timeout=REQUEST_TIMEOUT)
            if r.status_code != 200 or "json" not in r.headers.get("content-type", ""):
                continue
            recs = _extract_job_list_from_api(r.json())
            if recs:
                return recs
        except Exception:
            continue
    return []

def collect_job_records(api: dict, listing_label: str = "") -> list:
    """
    Page through the listing API and return raw job-record dicts, de-duplicated
    by their derived public job URL. MAX_PAGES caps the walk; we stop early when
    a page returns nothing or no new records.
    """
    base, path = api["base"], api["list"]
    print(C_BLUE(f"\n  Collecting jobs from API: {base}{path}"))
    seen, ordered = set(), []

    for page in range(1, MAX_PAGES + 1):
        recs = _fetch_list_page(base, path, page)
        if not recs:
            break
        new = 0
        for rec in recs:
            url = _job_url_from_record(rec) or _pick(rec, "id")
            key = url or json.dumps(rec, sort_keys=True)[:200]
            if key in seen:
                continue
            seen.add(key)
            ordered.append(rec)
            new += 1
        log(f"    page {page}: {new} new (total {len(ordered)})")
        if new == 0:
            break
        time.sleep(REQUEST_DELAY)

    return ordered

# ---------------------------------------------------------------------------
# STEP 2 — turn one API job record into a normalised raw_job dict
# ---------------------------------------------------------------------------

def _fetch_detail_record(api: dict, rec: dict) -> dict:
    """
    If the listing record lacks a body, fetch the per-job detail endpoint to get
    the full description / how-to-apply. Returns the richest dict available.
    """
    has_body = bool(_pick(rec, "description")) and len(_pick(rec, "description")) > 120
    jid = _pick(rec, "id")
    if has_body or not jid:
        return rec

    base = api["base"]
    for tmpl in api.get("detail", API_DETAIL_PATHS):
        path = tmpl.replace("{id}", jid)
        url  = base.rstrip("/") + "/" + path.lstrip("/")
        try:
            r = API_SESSION.get(url, timeout=REQUEST_TIMEOUT)
            if r.status_code != 200 or "json" not in r.headers.get("content-type", ""):
                continue
            data = r.json()
        except Exception:
            continue
        recs = _extract_job_list_from_api(data)
        if recs:
            merged = dict(rec)
            merged.update(recs[0])      # detail fields win
            return merged
        if isinstance(data, dict) and _is_job_api_response(data):
            merged = dict(rec)
            merged.update(data)
            return merged
    return rec

def _map_job_record(data: dict, job_url: str) -> dict:
    """Resolve an API job dict to our canonical raw_job output dict."""
    title        = clean_title(_pick(data, "title"))
    company      = sanitize_text(_pick(data, "company")) or "EthioJobs Employer"
    company_url  = sanitize_text(_pick(data, "company_url"),     is_url=True)
    company_logo = sanitize_text(_pick(data, "company_logo"),    is_url=True)
    company_website = sanitize_text(_pick(data, "company_website"), is_url=True)
    company_address = sanitize_text(_pick(data, "company_address"))
    company_about   = sanitize_text(_pick(data, "company_about"))
    salary          = sanitize_text(_pick(data, "salary"))

    location_raw = _pick(data, "location")
    location     = sanitize_text(location_raw) or DEFAULT_LOCATION

    job_type_raw = _pick(data, "job_type") or "Full-time"
    job_type     = JOB_TYPE_MAPPING.get(job_type_raw.lower().strip(), "full-time")

    deadline_raw    = _pick(data, "deadline")
    date_posted_raw = _pick(data, "date_posted")
    deadline    = parse_any_date(deadline_raw) if deadline_raw else ""
    date_posted = parse_any_date(date_posted_raw) if date_posted_raw else ""
    if not date_posted:
        date_posted = datetime.now().strftime("%Y-%m-%d")
    if not deadline:
        deadline = (datetime.now() + timedelta(days=30)).strftime("%Y-%m-%d")

    desc_raw  = _pick(data, "description")
    req_raw   = _pick(data, "requirements")
    apply_raw = _pick(data, "how_to_apply")

    # Description may be HTML — render to readable text.
    def _maybe_html(s):
        if s and ("<" in s and ">" in s):
            return html_block_to_text(BeautifulSoup(s, "lxml"))
        return s

    parts = []
    if desc_raw:
        parts.append(_maybe_html(desc_raw))
    if req_raw:
        parts.append("Requirements:\n" + _maybe_html(req_raw))
    if apply_raw:
        parts.append("How to Apply:\n" + _maybe_html(apply_raw))
    description = sanitize_text("\n\n".join(p for p in parts if p))

    cat_raw   = _pick(data, "category")
    job_field = infer_field(title, description, cat_raw)

    # Experience: API career level wins; else extract from requirements/body.
    career_level_raw = _pick(data, "career_level")
    experience_band  = _career_level_to_band(career_level_raw)
    if not experience_band:
        experience_band = extract_experience_band(_maybe_html(req_raw) or description)

    qualification = extract_qualification(_maybe_html(req_raw) or description)

    # Apply target: explicit API fields, then body scan, then policy opt-in.
    apply_email, apply_url = "", ""
    api_email = _pick(data, "apply_email")
    api_url   = _pick(data, "apply_url")
    if api_email and _is_real_apply_email(api_email):
        apply_email = api_email
    if api_url and _is_real_apply_url(api_url):
        apply_url = strip_tracking_params(api_url)

    scan = "\n".join([_maybe_html(apply_raw) or "", _maybe_html(req_raw) or "",
                      _maybe_html(desc_raw) or ""])
    if not apply_email:
        cand = extract_email(scan)
        if cand and _is_real_apply_email(cand):
            apply_email = cand
    if not apply_url:
        for u in URL_PATTERN.findall(scan):
            if _is_real_apply_url(u):
                apply_url = strip_tracking_params(u.rstrip(".,);"))
                break

    if APPLY_VIA_SOURCE_URL and not apply_email and not apply_url:
        apply_url = job_url

    if not salary:
        salary = extract_salary(scan)

    return {
        "title":           title,
        "company_name":    company,
        "company_url":     company_url,
        "company_website": company_website,
        "company_address": company_address,
        "company_logo":    company_logo,
        "company_about":   company_about,
        "job_type":        job_type,
        "location":        location,
        "job_field":       job_field,
        "job_categories":  cat_raw,
        "date_posted":     date_posted,
        "deadline":        deadline,
        "description":     description,
        "qualification":   qualification,
        "experience":      experience_band,
        "salary":          salary,
        "apply_email":     apply_email,
        "apply_url":       apply_url,
        "apply_text":      _maybe_html(apply_raw),
        "job_url":         job_url,
    }

def scrape_job_detail(rec, api: dict = None) -> dict:
    """
    Turn a listing API record into a normalised raw_job dict, fetching the detail
    endpoint first if the body is missing. `rec` is the listing record (dict).
    """
    if not isinstance(rec, dict):
        raise ValueError("scrape_job_detail expects an API job record dict")
    job_url = _job_url_from_record(rec)
    full    = _fetch_detail_record(api, rec) if api else rec
    return _map_job_record(full, job_url)

# --- one-shot --inspect: print the raw API record for the first job ----------
def inspect_first_job(api: dict):
    recs = _fetch_list_page(api["base"], api["list"], 1)
    if not recs:
        print(C_RED("  No records returned by the list endpoint."))
        return
    rec = recs[0]
    print(C_GREEN(f"\n  Listing record keys: {list(rec.keys())}"))
    for k, v in list(rec.items())[:25]:
        print(f"    {k!r:28s} -> {str(v)[:90]!r}")
    full = _fetch_detail_record(api, rec)
    if full is not rec:
        print(C_GREEN(f"\n  Detail record keys: {list(full.keys())}"))
    print(C_BLUE("\n  Mapped raw_job:"))
    mapped = _map_job_record(full, _job_url_from_record(rec))
    for k, v in mapped.items():
        print(f"    {k!r:18s} -> {str(v)[:90]!r}")

# =============================================================================
#  STEP 3 — DEDUPLICATE + PARAPHRASE + APPLY-RULE GATING
# =============================================================================

def process_job(raw_job: dict, processed_ids: set, processed_urls: set, seen_content: set):
    """
    Returns (status, job_dict_or_None):
        ("duplicate", None) — already processed / seen this run
        ("flagged",   None) — failed public-apply rule, written to flagged CSV
        ("ok",        dict) — ready to post to WordPress
    """
    job_url  = raw_job.get("job_url", "")
    title    = raw_job.get("title", "")
    company  = raw_job.get("company_name", "")
    location = raw_job.get("location", "")

    if not title:
        return "duplicate", None  # nothing usable

    job_id = make_job_id(job_url, title, company)

    if job_id in processed_ids or job_url in processed_urls:
        log(C_DIM(f"  Already processed (tracker) — skipped: {title}"))
        return "duplicate", None

    fingerprint = (title.lower().strip(), company.lower().strip(), location.lower().strip())
    if fingerprint in seen_content:
        log(C_DIM(f"  Duplicate content this run — skipped: {title}"))
        return "duplicate", None
    seen_content.add(fingerprint)

    # ---- Public-apply rule -------------------------------------------------
    apply_email = raw_job.get("apply_email", "")
    apply_url   = raw_job.get("apply_url", "")
    qualifies   = bool(apply_email) or bool(apply_url)

    if REQUIRE_PUBLIC_APPLY and not qualifies:
        write_flagged(raw_job,
                      "no public apply email or external URL (on-platform apply on "
                      "EthioJobs; set APPLY_VIA_SOURCE_URL=1 to post anyway)",
                      raw_job.get("apply_text", "")[:300])
        log(C_RED(f"  FLAGGED (no public apply) — {title}"))
        return "flagged", None

    # Record on scrape — before paraphrasing or posting.
    mark_scraped(job_id, job_url, title, company)
    processed_ids.add(job_id)
    processed_urls.add(job_url)

    description = raw_job.get("description", "")
    paraphrased_title = title
    paraphrased_desc  = description

    if ENABLE_PARAPHRASE and MISTRAL_API_KEY:
        print(C_BLUE(f"\n  Paraphrasing '{title}' ..."))
        paraphrased_title = paraphrase_title(title)
        paraphrased_desc  = paraphrase_description(description)
        mark_paraphrased(job_id)
    else:
        print(C_DIM("  Paraphrasing skipped (ENABLE_PARAPHRASE=False or MISTRAL_API_KEY not set)"))

    application = apply_email or apply_url
    apply_method = ("description_email" if apply_email
                    else "external_url" if apply_url else "not_found")

    company_link = raw_job.get("company_website") or raw_job.get("company_url", "")

    return "ok", {
        "jobTitle":          paraphrased_title,
        "jobDescription":    paraphrased_desc,
        "companyDetails":    "",
        "originalTitle":     title,
        "originalDesc":      description,
        "jobType":           raw_job.get("job_type", "full-time"),
        "jobQualifications": raw_job.get("qualification", ""),
        "jobExperience":     raw_job.get("experience", ""),
        "jobLocation":       location,
        "jobField":          raw_job.get("job_field", ""),
        "datePosted":        raw_job.get("date_posted", datetime.now().strftime("%Y-%m-%d")),
        "deadline":          raw_job.get("deadline", ""),
        "application":       application,
        "companyUrl":        company_link,
        "companyName":       company,
        "companyLogo":       raw_job.get("company_logo", ""),
        "companyWebsite":    raw_job.get("company_website", ""),
        "companyAddress":    raw_job.get("company_address", "") or location,
        "jobUrl":            job_url,
        "salaryRange":       raw_job.get("salary", ""),
        "_jobId":            job_id,
        "_apply_method":     apply_method,
        "_apply_raw":        raw_job.get("apply_text", "")[:160],
    }

# =============================================================================
#  VERBOSE PRINTER
# =============================================================================

def print_job_verbose(index, job):
    desc = job.get("jobDescription", "")
    desc_preview = (desc[:400] + " [...]") if len(desc) > 400 else desc

    print()
    print(C_DIVIDER())
    print(C_HEADER(f"  JOB #{index}"))
    print(C_DIVIDER())
    print(f"  {C_LABEL('Title (original)')}    : {C_VALUE(job.get('originalTitle',''))}")
    print(f"  {C_LABEL('Title (paraphrased)')} : {C_GREEN(job.get('jobTitle',''))}")
    print(f"  {C_LABEL('Job Type')}             : {job.get('jobType','') or C_DIM('—')}")
    print(f"  {C_LABEL('Qualification')}        : {(job.get('jobQualifications','')[:120] or C_DIM('—'))}")
    print(f"  {C_LABEL('Experience')}           : {job.get('jobExperience','') or C_DIM('—')}")
    print(f"  {C_LABEL('Location')}             : {job.get('jobLocation','') or C_DIM('—')}")
    print(f"  {C_LABEL('Category/Field')}       : {job.get('jobField','') or C_DIM('—')}")
    print(f"  {C_LABEL('Salary')}               : {job.get('salaryRange','') or C_DIM('—')}")
    print(f"  {C_LABEL('Posted')}               : {job.get('datePosted','') or C_DIM('—')}")
    print(f"  {C_LABEL('Deadline')}             : {job.get('deadline','') or C_DIM('—')}")

    application = job.get("application", "")
    print(f"  {C_LABEL('Apply')}                : {C_GREEN(application) if application else C_DIM('— not found —')}")
    print(f"  {C_LABEL('Apply Method')}         : {C_DIM(job.get('_apply_method',''))}")

    print()
    print(f"  {C_BLUE('── EMPLOYER ─────────────────────────────────────────')}")
    print(f"  {C_LABEL('Name')}      : {C_VALUE(job.get('companyName','') or C_DIM('—'))}")
    print(f"  {C_LABEL('Website')}   : {job.get('companyWebsite','') or C_DIM('—')}")
    print(f"  {C_LABEL('Source')}    : {job.get('companyUrl','') or C_DIM('—')}")
    print(f"  {C_LABEL('Logo')}      : {job.get('companyLogo','') or C_DIM('— none —')}")

    print()
    print(f"  {C_BLUE('── DESCRIPTION PREVIEW ─────────────────────────────')}")
    print(desc_preview if desc_preview else C_DIM("   — no description —"))
    print(f"  {C_LABEL('Job URL')}   : {job.get('jobUrl','')}")
    print(C_DIVIDER())

# =============================================================================
#  EXCEL SAVE (standardized column order)
# =============================================================================

EXCEL_HEADERS = [
    "Job Title", "Job Type", "Job Qualifications", "Job Experience",
    "Job Location", "Job Field", "Date Posted", "Deadline",
    "Job Description", "Application", "Company URL", "Company Name",
    "Company Logo", "Company Website", "Company Address",
    "Company Details", "Job URL", "Salary Range",
]

def _save_excel(jobs: list):
    if not _XLSX_AVAILABLE:
        log_.warning("pandas/openpyxl not installed — skipping Excel export")
        return
    if not jobs:
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(EXCEL_HEADERS)
    for job in jobs:
        ws.append([
            job["jobTitle"], job["jobType"], job["jobQualifications"], job["jobExperience"],
            job["jobLocation"], job["jobField"], job["datePosted"], job["deadline"],
            job["jobDescription"], job["application"], job["companyUrl"], job["companyName"],
            job["companyLogo"], job["companyWebsite"], job["companyAddress"],
            job["companyDetails"], job["jobUrl"], job["salaryRange"],
        ])
    wb.save(OUTPUT_FILE)
    log_.info(f"Saved {len(jobs)} rows -> {OUTPUT_FILE}")

# =============================================================================
#  MAIN
# =============================================================================

def main():
    import argparse
    parser = argparse.ArgumentParser(description="EthioJobs Ethiopia scraper (JSON API)")
    parser.add_argument("--discover", action="store_true",
                        help="Discover + print the API endpoints, then exit.")
    parser.add_argument("--inspect", action="store_true",
                        help="Print the raw + mapped record for the first job, then exit.")
    args = parser.parse_args()

    start_time = datetime.now()

    print()
    print(C_HEADER("=" * 80))
    print(C_HEADER("  ETHIOJOBS (ETHIOPIA) SCRAPER + MISTRAL PARAPHRASE + WORDPRESS POSTING"))
    print(C_HEADER("=" * 80))
    print(f"  Listing seed    : {JOBS_URL}")
    print(f"  API base (cfg)  : {API_BASE}")
    print(f"  Public-apply    : {'✅ enforced (flag others)' if REQUIRE_PUBLIC_APPLY else '❌ off (post all)'}")
    print(f"  Apply-via-URL   : {'✅ EthioJobs page = apply target' if APPLY_VIA_SOURCE_URL else '❌ off (strict)'}")
    print(f"  Max new jobs    : {'unlimited' if not MAX_JOBS else MAX_JOBS}")
    print(f"  Max pages       : {MAX_PAGES}  (page size {PAGE_SIZE})")
    print(f"  Paraphrase      : {'✅ enabled' if (ENABLE_PARAPHRASE and MISTRAL_API_KEY) else '❌ disabled'}")
    print(f"  WordPress post  : {'✅ enabled' if (WP_USER and WP_PASSWORD) else '❌ disabled'}")
    print(f"  Excel export    : {'✅ enabled' if _XLSX_AVAILABLE else '❌ disabled (pip install pandas openpyxl)'}")
    print(f"  NLP gating      : {'✅' if _NLP_AVAILABLE else '⚠️  no sentence-transformers / language-tool'}")
    print(f"  Started         : {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(C_HEADER("=" * 80))

    # Resolve the JSON API (cached after first success).
    api = discover_api(force=args.discover)
    if not api:
        log(C_RED("  Could not resolve the EthioJobs API — aborting."))
        sys.exit(2)

    if args.discover:
        print(C_GREEN(f"\n  base   : {api['base']}"))
        print(C_GREEN(f"  list   : {api['list']}"))
        print(C_GREEN(f"  detail : {api.get('detail')}"))
        print(C_DIM(f"  cached to {API_CACHE_FILE}"))
        return

    if args.inspect:
        inspect_first_job(api)
        return

    _init_tracker()
    _init_flagged()
    processed_ids, processed_urls = load_processed_ids()
    print(f"  Tracker loaded: {len(processed_ids)} previously processed job IDs")

    # Step 1: collect job records from the listing API (all configured listings).
    try:
        records = []
        seen_keys = set()
        for listing in LISTING_URLS:
            for rec in collect_job_records(api, listing):
                key = _job_url_from_record(rec) or json.dumps(rec, sort_keys=True)[:200]
                if key not in seen_keys:
                    seen_keys.add(key)
                    records.append(rec)
    except Exception as e:
        log(C_RED(f"  FATAL: could not collect job records: {e}"))
        return

    if not records:
        log(C_RED("  No job records found — nothing to do."))
        return
    print(C_GREEN(f"\n  Found {len(records)} job record(s) to process.\n"))

    jobs_out = []
    seen_content = set()
    posted_count = 0
    flagged_count = 0
    dup_count = 0
    errors = 0
    scraped = 0

    for rec in records:
        link = _job_url_from_record(rec)
        # Skip detail fetch entirely if URL already processed.
        if link and link in processed_urls:
            dup_count += 1
            log(C_DIM(f"  Already processed (tracker) — skipped: {link}"))
            continue

        try:
            raw_job = scrape_job_detail(rec, api)
            scraped += 1
        except Exception as e:
            errors += 1
            log(C_RED(f"  ERROR scraping {link or '<record>'} : {e}"))
            time.sleep(REQUEST_DELAY)
            continue

        try:
            status, job = process_job(raw_job, processed_ids, processed_urls, seen_content)
        except Exception as e:
            errors += 1
            log(C_RED(f"  ERROR processing '{raw_job.get('title','')}' : {e}"))
            continue

        if status == "duplicate":
            dup_count += 1
            time.sleep(REQUEST_DELAY)
            continue
        if status == "flagged":
            flagged_count += 1
            time.sleep(REQUEST_DELAY)
            continue

        jobs_out.append(job)
        print_job_verbose(len(jobs_out), job)

        print(C_BLUE("\n  Posting to WordPress …"))
        wp_id, wp_url = post_job_to_wordpress(job)
        if wp_id:
            mark_posted(job["_jobId"], wp_id, wp_url or "")
            posted_count += 1
            print(C_GREEN(f"  WP ID={wp_id}  {wp_url}"))
        else:
            mark_failed(job["_jobId"], "wp_post_failed_or_skipped")
            print(C_RED("  WordPress post failed / skipped"))

        if len(jobs_out) % 25 == 0:
            _save_excel(jobs_out)

        if MAX_JOBS and len(jobs_out) >= MAX_JOBS:
            log(f"\nMAX_JOBS limit ({MAX_JOBS}) reached, stopping.")
            break

        time.sleep(REQUEST_DELAY)

    _save_excel(jobs_out)

    end_time = datetime.now()
    duration = (end_time - start_time).total_seconds() / 60.0
    print()
    print(C_HEADER("=" * 80))
    print(C_HEADER("  SCRAPE COMPLETE"))
    print(C_HEADER("=" * 80))
    print(f"  {C_LABEL('Job records found')}         : {len(records)}")
    print(f"  {C_LABEL('Detail records scraped')}    : {scraped}")
    print(f"  {C_LABEL('New jobs processed')}        : {C_GREEN(str(len(jobs_out)))}")
    print(f"  {C_LABEL('Posted to WordPress')}       : {C_GREEN(str(posted_count))}")
    print(f"  {C_LABEL('Flagged (no public apply)')} : {flagged_count}")
    print(f"  {C_LABEL('Duplicates skipped')}        : {dup_count}")
    print(f"  {C_LABEL('Errors')}                    : {C_RED(str(errors)) if errors else '0'}")
    print(f"  {C_LABEL('Duration')}                  : ~{duration:.1f} min")
    print(f"  {C_LABEL('Output file')}               : {OUTPUT_FILE}")
    print(f"  {C_LABEL('Tracker file')}              : {PROCESSED_IDS_FILE}")
    print(f"  {C_LABEL('Flagged file')}              : {FLAGGED_FILE}")

    if jobs_out:
        with_apply = sum(1 for j in jobs_out if j.get("application"))
        with_email = sum(1 for j in jobs_out if "@" in (j.get("application") or ""))
        with_url   = with_apply - with_email
        print(f"\n  {C_LABEL('Application links:')}")
        print(f"    External URL : {with_url}")
        print(f"    Email found  : {with_email}")

        para_count = sum(1 for j in jobs_out if j.get("jobTitle") != j.get("originalTitle"))
        print(f"\n  {C_LABEL('Paraphrased titles')} : {para_count}/{len(jobs_out)}")

        with_deadline = sum(1 for j in jobs_out if j.get("deadline"))
        print(f"  {C_LABEL('Deadline captured')}  : {with_deadline}/{len(jobs_out)}")

    print(C_HEADER("=" * 80))


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(1)
