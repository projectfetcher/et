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

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

try:
    import pandas as pd
    import openpyxl
    _XLSX_AVAILABLE = True
except ImportError:
    _XLSX_AVAILABLE = False

try:
    import language_tool_python
    from sentence_transformers import SentenceTransformer, util as st_util
    _NLP_AVAILABLE = True
except ImportError:
    _NLP_AVAILABLE = False

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
    _PW_AVAILABLE = True
except ImportError:
    _PW_AVAILABLE = False

# =============================================================================
#  CONFIG
# =============================================================================

BASE_URL  = "https://ethiojobs.net"
JOBS_URL  = os.environ.get("ETHIOJOBS_JOBS_URL", "https://ethiojobs.net/jobs")

REQUIRE_PUBLIC_APPLY = os.environ.get("REQUIRE_PUBLIC_APPLY", "1") != "0"
REQUEST_DELAY        = float(os.environ.get("REQUEST_DELAY", "1.5"))
MAX_JOBS             = int(os.environ.get("MAX_JOBS", "0"))
MAX_PAGES            = int(os.environ.get("MAX_PAGES", "20"))
REQUEST_TIMEOUT      = int(os.environ.get("REQUEST_TIMEOUT", "30"))
HEADLESS             = os.environ.get("HEADLESS", "1") != "0"

OUTPUT_FILE        = "ethiojobs_ethiopia_jobs.xlsx"
PROCESSED_IDS_FILE = "ethiojobs_ethiopia_processed.csv"
FLAGGED_FILE       = "ethiojobs_ethiopia_flagged.csv"

_TRACKER_FIELDS = ["Job ID", "Job URL", "Job Title", "Company Name",
                   "Status", "Timestamp", "WP ID"]
_FLAGGED_FIELDS = ["Source", "Title", "Company", "Location", "Salary",
                   "Deadline", "Reason", "Apply Note", "Job URL", "Timestamp"]

# ── WordPress ─────────────────────────────────────────────────────────────────
WP_URL       = os.environ.get("WP_BASE_URL", "")
WP_USER      = os.environ.get("WP_USERNAME", "")
WP_PASSWORD  = os.environ.get("WP_APP_PASSWORD", "")
WP_BASE      = WP_URL.rstrip("/")
WP_JOBS_URL  = f"{WP_BASE}/job-listings"
WP_MEDIA_URL = f"{WP_BASE}/media"

# ── Mistral ───────────────────────────────────────────────────────────────────
MISTRAL_API_KEY = os.environ.get("MISTRAL_API_KEY", "")
MISTRAL_MODEL   = "mistral-small-latest"
MISTRAL_URL     = "https://api.mistral.ai/v1/chat/completions"
ENABLE_PARAPHRASE = True

# ── Startup warnings ──────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log_ = logging.getLogger(__name__)

for _var, _val, _feature in [
    ("MISTRAL_API_KEY", MISTRAL_API_KEY, "paraphrasing"),
    ("WP_USERNAME",     WP_USER,         "WordPress posting"),
    ("WP_APP_PASSWORD", WP_PASSWORD,     "WordPress posting"),
]:
    if not _val:
        log_.warning(f"{_var} not set — {_feature} disabled/skipped.")

if not _PW_AVAILABLE:
    log_.warning(
        "playwright not installed. Run: pip install playwright && playwright install chromium\n"
        "Without it the scraper cannot render EthioJobs' JavaScript pages."
    )

# =============================================================================
#  ETHIOPIA-SPECIFIC CONSTANTS
# =============================================================================

ETHIOPIA_LOCATIONS = [
    "Addis Ababa", "Dire Dawa", "Mekelle", "Gondar", "Hawassa", "Bahir Dar",
    "Dessie", "Jimma", "Jijiga", "Shashamane", "Bishoftu", "Sodo", "Arba Minch",
    "Hosaena", "Harar", "Dilla", "Nekemte", "Debre Birhan", "Asella", "Bale Robe",
    "Adama", "Nazret", "Gambela", "Assosa", "Semera", "Logia", "Gode",
    "Oromia", "Amhara", "Tigray", "SNNPR", "Somali", "Afar", "Benishangul",
    "Harari", "Sidama", "Central Ethiopia",
]
DEFAULT_LOCATION = os.environ.get("ETHIOJOBS_DEFAULT_LOCATION", "Ethiopia")

# EthioJobs category slugs → canonical job field mapping (site-specific).
# Used as fallback when keyword inference misses.
ETHIOJOBS_CATEGORY_MAP = {
    "accounting-and-finance":            "Finance & Accounting",
    "admin-and-secretarial":             "Administration & Operations",
    "banking-and-insurance":             "Finance & Accounting",
    "business-and-administration":       "Administration & Operations",
    "construction-and-engineering":      "Engineering",
    "customer-service":                  "Customer Service",
    "development-and-project-management":"Non-Profit & Social Work",
    "education":                         "Education & Training",
    "fmcg-and-manufacturing":            "Manufacturing & Production",
    "health-care":                       "Healthcare & Medicine",
    "hospitality-and-tourism":           "Hospitality & Tourism",
    "human-resource-and-recruitment":    "Human Resources",
    "ict":                               "Information Technology",
    "legal":                             "Legal",
    "logistics-and-supply-chain":        "Logistics & Supply Chain",
    "marketing-and-communication":       "Marketing & Communications",
    "media-and-journalism":              "Media & Journalism",
    "ngo-and-ingo":                      "Non-Profit & Social Work",
    "sales-and-marketing":               "Sales & Business Development",
    "social-sciences-and-community-service": "Non-Profit & Social Work",
}

JOB_TYPE_MAPPING = {
    "full-time": "full-time", "full time": "full-time",
    "part-time": "part-time", "part time": "part-time",
    "contract":  "contract",  "temporary": "temporary",
    "internship":"internship","freelance": "freelance",
    "volunteer": "volunteer", "permanent": "full-time",
}

_NON_APPLY_HOST_SUBSTR = (
    "ethiojobs.net", "facebook.", "twitter.", "x.com", "linkedin.",
    "instagram.", "wa.me", "whatsapp", "t.me", "telegram",
    "plus.google", "pinterest.", "youtube.",
)
_NON_APPLY_PATH_SUBSTR = (
    "/signIn", "/login", "/register", "/sign-up", "action=login",
    "#share", "/share", "/wp-login", "/cart", "/checkout",
)
_NON_APPLY_EMAIL_DOMAINS = ("ethiojobs.net", "ethiojobs.com")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml",
}
SESSION = requests.Session()
SESSION.headers.update(HEADERS)

# =============================================================================
#  COLOUR / LOGGING
# =============================================================================

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
TRACKING_PARAM_EXACT = {"fbclid", "gclid", "msclkid", "mc_cid", "mc_eid", "ref", "referrer"}

MONTHS = {
    "january":1,"february":2,"march":3,"april":4,"may":5,"june":6,
    "july":7,"august":8,"september":9,"october":10,"november":11,"december":12,
    "jan":1,"feb":2,"mar":3,"apr":4,"jun":6,"jul":7,
    "aug":8,"sep":9,"sept":9,"oct":10,"nov":11,"dec":12,
}

TEXT_DATE_RE = re.compile(
    r"(\d{1,2})\s*(?:st|nd|rd|th)?\s+([A-Za-z]+)\s*[.,]?\s*(\d{4})", re.I
)
DMY_DATE_RE  = re.compile(r"\b(\d{1,2})[/-](\d{1,2})[/-](\d{4})\b")
# EthioJobs uses "Jun 20, 2026" style  
MDY_DATE_RE  = re.compile(
    r"\b([A-Za-z]{3,9})\s+(\d{1,2})[,\s]+(\d{4})\b", re.I
)

DEADLINE_LABELS = (
    "application deadline", "closing date", "deadline", "apply by",
    "application closing date", "expiry date", "expires", "close date",
)

_APPLY_HEAD_PHRASES = re.compile(
    r"^(?:how\s*(?:and|&)\s*deadline\s*to\s*apply|how\s*to\s*apply(?:\s*(?:and|&)\s*deadline)?|"
    r"how\s*to\s*submit|to\s*apply|application\s*(?:and|&)\s*deadline|"
    r"mode\s*of\s*application|method\s*of\s*application|"
    r"application\s*(?:procedure|process|instructions?|method|guidelines?)|"
    r"submission\s*of\s*applications?|deadline\s*(?:and|&)?\s*(?:how\s*)?to\s*apply)\b",
    re.I,
)

_BODY_CUT_MARKERS = [
    "related jobs", "leave your thoughts", "you must be logged in",
    "email me jobs like these", "send to a friend", "leave a reply",
    "similar jobs", "other jobs", "share this job",
]
_BODY_DROP_LINES = {
    "apply now", "apply for this job", "save", "share", "share:", "bookmark job",
    "quick view", "send to friend", "send to a friend", "clear all",
    "filter", "view more", "login to apply", "sign in to apply",
}

# =============================================================================
#  TEXT CLEANUP
# =============================================================================

_MOJIBAKE = [
    ("Â", ""), ("â€™", "'"), ("â€œ", '"'), ("â€\x9d", '"'), ("â€", '"'),
    ("â€¢", "•"), ("â„¢", "™"), ("\u00a0", " "), ("\u200b", ""), ("\ufeff", ""),
]

def _fix_mojibake(text):
    for pattern, replacement in _MOJIBAKE:
        text = text.replace(pattern, replacement)
    text = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", "", text)
    return text

def sanitize_text(text, is_url=False):
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
    kept = [(k, v) for k, v in parse_qsl(parts.query, keep_blank_values=True)
            if not k.lower().startswith(TRACKING_PARAM_PREFIXES)
            and k.lower() not in TRACKING_PARAM_EXACT]
    return urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(kept), parts.fragment))

def slugify(text, maxlen=80):
    s = re.sub(r"[^a-z0-9]+", "-", (text or "").lower()).strip("-")
    return s[:maxlen] or "job"

def html_block_to_text(el):
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

def dmy_dates(text):
    out = []
    for d, m, y in DMY_DATE_RE.findall(text or ""):
        try:
            out.append(datetime(int(y), int(m), int(d)).strftime("%Y-%m-%d"))
        except ValueError:
            pass
    return out

def mdy_dates(text):
    """Parse 'Jun 20, 2026' style dates common on EthioJobs."""
    out = []
    for mon, d, y in MDY_DATE_RE.findall(text or ""):
        month = MONTHS.get(mon.lower())
        if not month:
            continue
        try:
            out.append(datetime(int(y), month, int(d)).strftime("%Y-%m-%d"))
        except ValueError:
            pass
    return out

def text_dates(text):
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

def parse_any_date(text):
    for fn in (dmy_dates, mdy_dates, text_dates):
        ds = fn(text or "")
        if ds:
            return ds[-1]
    return ""

def clean_title(raw):
    t = sanitize_text(raw)
    t = re.sub(r"\s*\d[\d,]*\s*views?\s*$", "", t, flags=re.I)
    return t.strip()

def map_job_type(raw):
    key = (raw or "").lower().strip()
    return JOB_TYPE_MAPPING.get(key, "full-time")

def pick_location(locations):
    specific = [l for l in locations
                if l and l.strip().lower() not in ("ethiopia", "all", "multiple")]
    if specific:
        return specific[0].strip()
    if locations:
        return locations[0].strip()
    return DEFAULT_LOCATION

def location_from_text(text):
    if text:
        for town in ETHIOPIA_LOCATIONS:
            if re.search(rf"\b{re.escape(town)}\b", text, re.I):
                return town
    return DEFAULT_LOCATION

def extract_salary(text):
    if not text:
        return ""
    m = re.search(r"(?:ETB|Birr|USD|\$)\s*([0-9]{1,3}(?:,\s?[0-9]{3})+(?:\.[0-9]+)?)", text, re.I)
    if m:
        amt = re.sub(r"\s+", "", m.group(1))
        prefix = re.search(r"ETB|Birr|USD|\$", m.group(0), re.I).group(0)
        return f"{prefix} {amt}"
    m = re.search(r"\b(?:salary|remuneration|compensation|pay)\b[^.\n]{0,80}", text, re.I)
    if m and re.search(r"\d", m.group(0)):
        return m.group(0).strip().rstrip(".")
    return ""

# =============================================================================
#  CANONICAL NORMALISERS (qualification / experience / job field)
# =============================================================================

def _kw_hit(text_low, keywords):
    for k in keywords:
        kk = k.strip().lower()
        if not kk:
            continue
        esc = re.escape(kk)
        if len(kk) <= 3:
            pat = r"(?<![a-z0-9])" + esc + r"(?![a-z0-9])"
        else:
            pat = r"(?<![a-z0-9])" + esc + r"(?:es|s)?(?![a-z0-9])"
        if re.search(pat, text_low):
            return True
    return False

QUALIFICATION_TIERS = [
    ("PhD / Doctorate",
     ["phd", "ph.d", "doctorate", "doctoral", "doctor of philosophy"]),
    ("Master's Degree",
     ["master", "msc", "m.sc", "ma ", "m.a ", "mba", "m.b.a", "meng",
      "m.eng", "mphil", "postgraduate", "post-graduate", "post graduate"]),
    ("Bachelor's Degree",
     ["bachelor", "bsc", "b.sc", "ba ", "b.a ", "beng", "b.eng", "bcom",
      "b.com", "bba", "llb", "degree in", "undergraduate degree",
      "honours degree", "hons"]),
    ("Higher National Diploma",
     ["hnd", "hnc", "higher national diploma", "higher diploma", "advanced diploma"]),
    ("Diploma",
     ["diploma", "dip ", "dip.", "associate degree", "foundation degree",
      "level iv", "level 4", "tvet", "tveta"]),   # TVET is common in Ethiopia
    ("Professional Certification",
     ["acca", "cpa", "cfa", "cima", "pmp", "prince2", "cissp",
      "aws certified", "comptia", "cisco", "ccna", "ccnp", "shrm",
      "cipd", "chartered", "certified", "professional certificate"]),
    ("A-Levels / HSC",
     ["a-level", "a level", "hsc", "higher school certificate"]),
    ("O-Levels / School Certificate",
     ["o-level", "o level", "igcse", "gcse", "school certificate",
      "grade 10", "grade 12"]),
    ("No Formal Qualification Required",
     ["no qualification", "no degree", "no formal", "entry level",
      "no experience required", "training provided", "will train"]),
]

def extract_qualification(text):
    if not text:
        return ""
    lower = text.lower()
    for label, keywords in QUALIFICATION_TIERS:
        if _kw_hit(lower, keywords):
            return label
    return ""

NO_EXP_KW  = ["no experience", "no prior experience", "fresh graduate", "freshers",
               "entry level", "entry-level", "0 years", "zero experience",
               "training provided", "will train", "no experience required"]
LESS1_KW   = ["less than 1 year", "under 1 year", "6 months",
               "less than a year", "some experience", "minimal experience"]
_EXP_CAP   = 20
_EXP_REQ_RE   = re.compile(
    r"(?:minimum|min\.?|at\s+least|atleast|least|over|more\s+than|not\s+less\s+than|"
    r"minimum\s+of|a\s+minimum\s+of)\s+(?:of\s+)?(\d{1,2})\s*\+?\s*years?", re.I)
_EXP_YEARS_OF_RE = re.compile(r"(\d{1,2})\s*\+?\s*years?\s+of\b", re.I)
_EXP_ANY_YEARS_RE = re.compile(r"(\d{1,2})\s*\+?\s*years?", re.I)
_EXP_RANGE_RE = re.compile(r"(\d{1,2})\s*(?:-|–|to)\s*(\d{1,2})\s*years?", re.I)

def years_to_band(n):
    if n <= 0:  return "No Experience Required"
    if n <= 2:  return "1 - 2 Years"
    if n <= 5:  return "3 - 5 Years"
    if n <= 10: return "6 - 10 Years"
    return "10+ Years"

def extract_experience_band(text):
    if not text:
        return ""
    low = text.lower()
    years = []
    for m in _EXP_REQ_RE.finditer(text):
        n = int(m.group(1))
        if 0 < n <= _EXP_CAP:
            years.append(n)
    for m in _EXP_YEARS_OF_RE.finditer(low):
        n = int(m.group(1))
        if 0 < n <= _EXP_CAP:
            years.append(n)
    for m in _EXP_ANY_YEARS_RE.finditer(low):
        n = int(m.group(1))
        if 0 < n <= _EXP_CAP and "experien" in low[m.end():m.end() + 60]:
            years.append(n)
    for m in _EXP_RANGE_RE.finditer(text):
        a = int(m.group(1))
        if 0 < a <= _EXP_CAP:
            years.append(a)
    if years:
        return years_to_band(min(years))
    if _kw_hit(low, NO_EXP_KW):
        return "No Experience Required"
    if _kw_hit(low, LESS1_KW):
        return "1 - 2 Years"
    return ""

# Field keyword map — identical to GamJobs template, works globally.
FIELD_KEYWORD_MAP = [
    ("Information Technology",
     ["software engineer", "developer", "devops", "frontend", "backend", "full stack", "fullstack",
      "sysadmin", "cloud", "cybersecurity", "data engineer", "machine learning", "artificial intelligence",
      "ai/ml", "it support", "network engineer", "database", "kubernetes", "docker", "aws", "azure",
      "react", "node.js", "python developer", "java developer"],
     ["programming", "coding", "api", "agile", "scrum", "git", "linux", "server", "software"]),
    ("Finance & Accounting",
     ["accountant", "auditor", "finance manager", "financial analyst", "cfo", "treasurer", "tax",
      "bookkeeper", "payroll", "credit analyst", "investment", "actuary", "acca", "cfa", "cpa"],
     ["financial", "accounting", "balance sheet", "reconciliation", "ifrs", "ledger", "invoicing"]),
    ("Sales & Business Development",
     ["sales executive", "sales manager", "business development", "account manager",
      "sales representative", "bd manager", "regional sales", "key account", "sales director"],
     ["revenue", "pipeline", "crm", "leads", "prospects", "quota", "target", "b2b", "b2c"]),
    ("Marketing & Communications",
     ["marketing manager", "digital marketing", "seo", "sem", "content marketer", "social media manager",
      "brand manager", "marketing executive", "communications manager", "pr manager", "copywriter"],
     ["marketing", "branding", "advertising", "social media", "content", "campaign"]),
    ("Human Resources",
     ["hr manager", "human resources", "recruiter", "talent acquisition", "hr business partner",
      "hrbp", "hr officer", "compensation", "learning and development", "l&d", "hr generalist"],
     ["recruitment", "onboarding", "performance management", "employee relations", "hr", "workforce"]),
    ("Engineering",
     ["mechanical engineer", "civil engineer", "electrical engineer", "structural engineer",
      "process engineer", "project engineer", "maintenance engineer", "production engineer",
      "quality engineer", "safety engineer", "site engineer", "design engineer"],
     ["engineering", "cad", "autocad", "manufacturing", "plant", "machinery"]),
    ("Healthcare & Medicine",
     ["doctor", "physician", "nurse", "pharmacist", "medical officer", "surgeon",
      "physiotherapist", "radiographer", "lab technician", "clinical", "dentist", "midwife",
      "health officer", "public health"],
     ["hospital", "clinic", "patient", "medical", "health", "pharmaceutical", "diagnosis"]),
    ("Education & Training",
     ["teacher", "lecturer", "professor", "trainer", "educator", "tutor", "school principal",
      "academic", "curriculum", "instructional designer", "teaching assistant"],
     ["school", "university", "college", "classroom", "students", "pedagogy", "education"]),
    ("Hospitality & Tourism",
     ["hotel manager", "front desk", "housekeeping", "chef", "sous chef", "food and beverage",
      "f&b manager", "restaurant manager", "bartender", "waiter", "concierge", "tour guide",
      "travel agent", "catering"],
     ["hospitality", "hotel", "resort", "tourism", "guest", "accommodation", "restaurant"]),
    ("Logistics & Supply Chain",
     ["supply chain manager", "logistics coordinator", "warehouse manager", "fleet manager",
      "procurement manager", "purchasing manager", "freight", "inventory manager"],
     ["logistics", "supply chain", "warehouse", "inventory", "procurement", "sourcing"]),
    ("Legal",
     ["lawyer", "attorney", "legal counsel", "paralegal", "compliance officer", "legal advisor",
      "solicitor", "barrister", "corporate counsel", "contract manager"],
     ["legal", "law", "contracts", "litigation", "regulatory", "compliance"]),
    ("Administration & Operations",
     ["office manager", "executive assistant", "administrative officer", "operations manager",
      "personal assistant", "receptionist", "data entry", "office administrator", "secretary"],
     ["administration", "operations", "office", "coordination", "scheduling", "reporting"]),
    ("Customer Service",
     ["customer service", "call centre", "customer success", "customer support", "help desk",
      "service advisor", "client relations", "customer experience"],
     ["customer", "support", "helpdesk", "tickets", "satisfaction", "service level"]),
    ("Construction & Real Estate",
     ["quantity surveyor", "site supervisor", "architect", "draughtsman", "property manager",
      "estate agent", "real estate", "building inspector", "land surveyor", "construction manager"],
     ["construction", "building", "property", "real estate", "site", "contractor"]),
    ("Manufacturing & Production",
     ["production manager", "quality control", "quality assurance", "qa", "qc", "factory manager",
      "plant manager", "production supervisor", "assembly", "cnc operator"],
     ["production", "manufacturing", "factory", "assembly", "quality", "lean"]),
    ("Design & Creative",
     ["graphic designer", "ui/ux", "product designer", "art director", "creative director",
      "animator", "illustrator", "photographer", "videographer", "web designer"],
     ["design", "creative", "adobe", "figma", "photoshop", "indesign", "branding"]),
    ("Research & Science",
     ["research scientist", "data scientist", "lab researcher", "research analyst",
      "clinical researcher", "environmental scientist", "chemist", "biologist", "statistician"],
     ["research", "analysis", "laboratory", "science", "experiment", "findings"]),
    ("Security",
     ["security officer", "security guard", "security manager", "loss prevention",
      "risk manager", "health and safety", "hse officer", "fire safety"],
     ["security", "safety", "risk", "surveillance", "patrol", "access control"]),
    ("Non-Profit & Social Work",
     ["social worker", "ngo", "charity", "programme coordinator", "community development",
      "welfare officer", "case manager", "development officer", "fundraiser"],
     ["social", "ngo", "community", "welfare", "beneficiary", "donor", "impact", "charity"]),
]

_TENDER_TITLE_RE = re.compile(
    r"\b(?:rfq|rfp|reoi|eoi|itb|itt|spn|rfb|rfa|gpn|ifb|rfi)\b"
    r"|invitation\s+to\s+(?:bid|tender)|request\s+for\s+(?:quotation|proposal|bids?)"
    r"|expressions?\s+of\s+interest|\btenders?\b|procurement\s+notice"
    r"|call\s+for\s+(?:bid|tenders?|proposals?|quotation)",
    re.I,
)
TENDER_FIELD = "Public Notices & Tenders"

def infer_field(title, description, fallback_categories=""):
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
    # Use the site's own category slug as fallback.
    for slug in (fallback_categories or "").split(","):
        slug = slug.strip().lower().replace(" ", "-")
        if slug in ETHIOJOBS_CATEGORY_MAP:
            return ETHIOJOBS_CATEGORY_MAP[slug]
    if fallback_categories:
        cats = [c.strip() for c in fallback_categories.split(",") if c.strip()]
        if cats:
            return cats[0]
    return ""

# =============================================================================
#  NLP TOOLS (optional)
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

def grammar_correct(text):
    tool = _get_grammar_tool()
    if tool:
        try:
            return language_tool_python.utils.correct(text, tool.check(text))
        except Exception:
            pass
    return text

def similarity_score(a, b):
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

def clean_output(text):
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

def mistral_generate(prompt, max_tokens=400, temperature=0.7):
    if not MISTRAL_API_KEY:
        return ""
    try:
        r = requests.post(
            MISTRAL_URL,
            headers={"Authorization": f"Bearer {MISTRAL_API_KEY}",
                     "Content-Type": "application/json"},
            json={"model": MISTRAL_MODEL,
                  "messages": [{"role": "user", "content": prompt}],
                  "max_tokens": max_tokens, "temperature": temperature},
            timeout=30,
        )
        r.raise_for_status()
        return r.json()["choices"][0]["message"]["content"].strip()
    except Exception as e:
        log_.error(f"Mistral API error: {e}")
        return ""

# =============================================================================
#  PARAPHRASE FUNCTIONS
# =============================================================================

def _print_wrapped(text, prefix="   ", width=100):
    words = text.split()
    line = []
    for w in words:
        line.append(w)
        if len(" ".join(line)) >= width:
            print(f"{prefix}{' '.join(line)}")
            line = []
    if line:
        print(f"{prefix}{' '.join(line)}")

def paraphrase_title(title):
    if not ENABLE_PARAPHRASE:
        return title
    clean = sanitize_text(title)
    if not clean:
        return title
    print(f"\n ┌─ TITLE PARAPHRASE {'─'*45}")
    print(f" │ Original : \"{clean}\"")
    best_result, best_sim = None, 0.0
    for attempt in range(4):
        temp = round(0.68 + attempt * 0.06, 2)
        prompt = (
            f"Rewrite this job title professionally using different words. "
            f"Output ONLY the rewritten title, nothing else. Keep it 4–12 words.\n\n"
            f"Job title: {clean}"
        )
        raw = mistral_generate(prompt, max_tokens=50, temperature=temp)
        result = clean_output(raw).split("\n")[0].strip().strip('"').strip("'")
        wc = len(result.split()) if result else 0
        sim = similarity_score(clean, result) if result else 0.0
        is_dup = result.lower().strip() == clean.lower().strip()
        valid = bool(result) and 4 <= wc <= 14 and sim >= 0.55 and not is_dup
        if valid and sim > best_sim:
            best_sim, best_result = sim, result
        time.sleep(1)
    if best_result:
        print(f" │ FINAL : \"{best_result}\" (sim={best_sim:.3f})")
        print(f" └{'─'*65}")
        return best_result
    print(f" │ No valid paraphrase → keeping original")
    print(f" └{'─'*65}")
    return clean

def paraphrase_description(text):
    if not ENABLE_PARAPHRASE:
        return text
    clean = sanitize_text(text)
    if not clean:
        return text
    paragraphs = [p.strip() for p in re.split(r"\n+", clean) if p.strip()] or [clean]
    rewritten = []
    for para in paragraphs:
        if len(para.split()) < 8:
            rewritten.append(para)
            continue
        prompt = (
            f"Rewrite this job description paragraph professionally. "
            f"Keep ALL facts, requirements, and responsibilities. "
            f"Use different sentence structure and vocabulary. "
            f"Output ONLY the rewritten paragraph.\n\nOriginal:\n{para}"
        )
        best_result, best_sim, accepted = None, 0.0, None
        for attempt in range(3):
            raw = mistral_generate(prompt, max_tokens=500, temperature=round(0.65 + attempt * 0.08, 2))
            result = clean_output(raw).strip()
            rw = len(result.split()) if result else 0
            sim = similarity_score(para, result) if result and rw >= 5 else 0.0
            if bool(result) and rw >= 8 and sim >= 0.48:
                accepted = result
                break
            if result and sim > best_sim:
                best_sim, best_result = sim, result
            time.sleep(1)
        rewritten.append(accepted or (best_result if best_result and best_sim >= 0.40 else para))
    return "\n\n".join(rewritten)

# =============================================================================
#  DUPLICATE TRACKER
# =============================================================================

def _init_tracker():
    if not os.path.exists(PROCESSED_IDS_FILE):
        try:
            with open(PROCESSED_IDS_FILE, "w", newline="", encoding="utf-8") as f:
                csv.writer(f).writerow(_TRACKER_FIELDS)
        except Exception as e:
            log_.error(f"Could not create tracker: {e}")

def load_processed_ids():
    _init_tracker()
    ids, urls = set(), set()
    try:
        with open(PROCESSED_IDS_FILE, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                if row.get("Job ID"):   ids.add(row["Job ID"].strip())
                if row.get("Job URL"):  urls.add(row["Job URL"].strip())
    except Exception as e:
        log_.error(f"Tracker read error: {e}")
    return ids, urls

def _upsert_row(job_id, updates):
    _init_tracker()
    rows = []
    try:
        with open(PROCESSED_IDS_FILE, newline="", encoding="utf-8") as f:
            rows = list(csv.DictReader(f))
    except Exception:
        pass
    found = False
    for row in rows:
        if row.get("Job ID", "").strip() == str(job_id):
            row.update(updates)
            row["Timestamp"] = datetime.now().isoformat()
            found = True
            break
    if not found:
        new_row = {k: "" for k in _TRACKER_FIELDS}
        new_row.update({"Job ID": str(job_id), "Timestamp": datetime.now().isoformat()})
        new_row.update(updates)
        rows.append(new_row)
    try:
        with open(PROCESSED_IDS_FILE, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=_TRACKER_FIELDS, extrasaction="ignore")
            w.writeheader(); w.writerows(rows)
    except Exception as e:
        log_.error(f"Tracker write error: {e}")

def make_job_id(job_url, title="", company=""):
    seed = job_url or f"{title}{company}"
    return hashlib.md5(seed.encode()).hexdigest()[:16]

def mark_scraped(job_id, job_url, title, company):
    _upsert_row(job_id, {"Job URL": job_url, "Job Title": title,
                          "Company Name": company, "Status": "scraped", "WP ID": ""})

def mark_paraphrased(job_id):
    _upsert_row(job_id, {"Status": "paraphrased"})

def mark_posted(job_id, wp_id, wp_url):
    _upsert_row(job_id, {"Status": "posted", "WP ID": str(wp_id)})

def mark_failed(job_id, reason):
    _upsert_row(job_id, {"Status": f"failed|{reason}"})

# =============================================================================
#  FLAGGED CSV
# =============================================================================

def _init_flagged():
    if not os.path.exists(FLAGGED_FILE):
        try:
            with open(FLAGGED_FILE, "w", newline="", encoding="utf-8") as f:
                csv.writer(f).writerow(_FLAGGED_FIELDS)
        except Exception as e:
            log_.error(f"Could not create flagged file: {e}")

def write_flagged(raw_job, reason, apply_note):
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
#  PLAYWRIGHT BROWSER MANAGER
# =============================================================================

class BrowserManager:
    """
    Wraps a single long-lived Playwright Chromium instance.
    Use as a context manager:
        with BrowserManager() as bm:
            html = bm.get_html("https://...")
            soup = bm.get_soup("https://...")
    """
    def __init__(self, headless=True):
        self._pw       = None
        self._browser  = None
        self._page     = None
        self._headless = headless

    def __enter__(self):
        if not _PW_AVAILABLE:
            raise RuntimeError(
                "playwright is not installed.\n"
                "Run: pip install playwright && playwright install chromium"
            )
        self._pw      = sync_playwright().__enter__()
        self._browser = self._pw.chromium.launch(headless=self._headless)
        ctx           = self._browser.new_context(
            user_agent=HEADERS["User-Agent"],
            extra_http_headers={"Accept-Language": "en-US,en;q=0.9"},
        )
        self._page = ctx.new_page()
        # Abort image/media requests to speed up scraping.
        self._page.route(
            "**/*.{png,jpg,jpeg,gif,webp,svg,woff,woff2,ttf,mp4,mp3,mpeg}",
            lambda route: route.abort()
        )
        return self

    def __exit__(self, *_):
        try:
            if self._browser:
                self._browser.close()
            if self._pw:
                self._pw.__exit__(None, None, None)
        except Exception:
            pass

    def get_html(self, url, wait_selector=None, timeout=30_000):
        """Navigate to url, optionally wait for a CSS selector, return HTML."""
        self._page.goto(url, wait_until="domcontentloaded", timeout=timeout)
        if wait_selector:
            try:
                self._page.wait_for_selector(wait_selector, timeout=timeout)
            except PWTimeout:
                pass
        else:
            # Give the SPA up to 5 s to render job cards.
            try:
                self._page.wait_for_timeout(3000)
            except Exception:
                pass
        return self._page.content()

    def get_soup(self, url, wait_selector=None, timeout=30_000):
        html = self.get_html(url, wait_selector=wait_selector, timeout=timeout)
        try:
            return BeautifulSoup(html, "lxml")
        except Exception:
            return BeautifulSoup(html, "html.parser")

# =============================================================================
#  STEP 1 — COLLECT JOB DETAIL URLs FROM ETHIOJOBS LISTING PAGES
# =============================================================================
#
#  EthioJobs URL patterns observed:
#    Listing  : https://ethiojobs.net/jobs                    (paginated via ?page=N)
#    Category : https://ethiojobs.net/jobs/accounting-and-finance
#    Detail   : https://ethiojobs.net/display-job/{ID}/{slug}
#               (old) https://www.ethiojobs.net/display-job/{ID}/{Slug}.html
#
#  The scraper walks the main /jobs listing and also accepts a list of category
#  URLs via ETHIOJOBS_EXTRA_URLS env var (comma-separated).

def _norm_job_url(href):
    if not href:
        return ""
    absu = urljoin(BASE_URL + "/", href)
    p = urlsplit(absu)
    # Normalise: always https, strip fragment.
    return urlunsplit(("https", p.netloc.lower(), p.path.rstrip("/"), p.query, ""))

def _is_job_detail_path(path):
    """
    True for /display-job/{id}/{slug} paths (the EthioJobs detail page pattern).
    Also accepts /jobs/{numeric-id} if the site ever uses that form.
    """
    parts = [s for s in path.split("/") if s]
    if len(parts) >= 2 and parts[0] == "display-job":
        return True
    return False

def _listing_page_url(base, page):
    if page <= 1:
        return base
    sep = "&" if "?" in base else "?"
    return f"{base}{sep}page={page}"

def collect_job_links(bm: BrowserManager, listing_url: str, max_pages: int = MAX_PAGES):
    """Walk paginated listing and return ordered, de-duplicated detail URLs."""
    print(C_BLUE(f"\n  Collecting job links from: {listing_url}"))

    # EthioJobs renders a job card grid; wait for at least one card.
    CARD_SEL = "a[href*='/display-job/']"

    seen, ordered = set(), []
    empty_streak = 0

    for page in range(1, max_pages + 1):
        url = _listing_page_url(listing_url, page)
        log(f"    Fetching listing page {page}: {url}")
        try:
            soup = bm.get_soup(url, wait_selector=CARD_SEL)
        except Exception as e:
            log(C_DIM(f"  Page {page}: error ({e}) — stopping."))
            break

        page_new = 0
        for a in soup.find_all("a", href=True):
            href = a["href"].strip()
            path = urlparse(urljoin(BASE_URL + "/", href)).path
            if not _is_job_detail_path(path):
                continue
            norm = _norm_job_url(href)
            if norm and norm not in seen:
                seen.add(norm)
                ordered.append(norm)
                page_new += 1

        log(f"      → {page_new} new link(s) (total {len(ordered)})")

        if page_new == 0:
            empty_streak += 1
            if empty_streak >= 2:
                log(C_DIM("  Two consecutive empty pages — stopping pagination."))
                break
        else:
            empty_streak = 0

        time.sleep(REQUEST_DELAY)

    return ordered

# =============================================================================
#  STEP 2 — PARSE ONE ETHIOJOBS DETAIL PAGE
# =============================================================================

# Selectors tried in order for the main job body.
_CONTENT_SELECTORS = [
    "div.job-description-content",
    "div.job-detail-description",
    "div.description",
    "div.job-description",
    "section.job-description",
    "div[class*='job-detail']",
    "div[class*='jobDetail']",
    "div[class*='job-content']",
    "article .entry-content",
    "div.entry-content",
    "main section",
    "main",
]

# EthioJobs detail page typically shows structured meta in a sidebar/infobox.
# These are the label strings we look for.
_META_LABELS = {
    "career level":     "career_level",
    "employment type":  "job_type",
    "salary":           "salary",
    "deadline":         "deadline",
    "closing date":     "deadline",
    "application deadline": "deadline",
    "location":         "location",
    "place of work":    "location",
    "number of positions": "positions",
    "education":        "qualification",
    "experience":       "experience_raw",
}

def _is_real_apply_email(email):
    if not email or "@" not in email:
        return False
    dom = email.rsplit("@", 1)[-1].lower()
    return not any(dom == d or dom.endswith("." + d) for d in _NON_APPLY_EMAIL_DOMAINS)

def _is_real_apply_url(href):
    if not href:
        return False
    low = href.lower()
    if low.startswith(("mailto:", "#", "javascript:")):
        return False
    if not low.startswith("http"):
        return False
    if any(s in low for s in _NON_APPLY_HOST_SUBSTR):
        return False
    if any(s in low for s in _NON_APPLY_PATH_SUBSTR):
        return False
    return True

def _is_apply_heading_line(line):
    s = line.strip().lstrip("•*-–—#:. ").strip()
    if not s or len(s.split()) > 9:
        return False
    return bool(_APPLY_HEAD_PHRASES.match(s))

def _split_description_and_apply(content_text):
    if not content_text:
        return "", ""
    lines = content_text.split("\n")
    kept = []
    for ln in lines:
        low = ln.strip().lower()
        if low in _BODY_DROP_LINES:
            continue
        if any(low.startswith(m) for m in _BODY_CUT_MARKERS):
            break
        kept.append(ln)
    apply_idx = None
    for i, ln in enumerate(kept):
        if _is_apply_heading_line(ln):
            apply_idx = i
            break
    if apply_idx is None:
        return "\n".join(kept).strip(), ""
    description = "\n".join(kept[:apply_idx]).strip()
    apply_text  = "\n".join(kept[apply_idx:]).strip()
    if not description:
        return "\n".join(kept).strip(), ""
    return description, apply_text

def _extract_meta_table(soup):
    """
    Parse the structured metadata sidebar/table on EthioJobs detail pages.
    Returns a dict of field_key -> raw_value_string.
    """
    meta = {}
    # Strategy 1: look for a <table> or dl/ul with labelled rows.
    for row in soup.find_all(["tr", "li", "div"]):
        cells = row.find_all(["td", "th", "dt", "dd", "span"])
        if len(cells) >= 2:
            label = cells[0].get_text(" ", strip=True).lower().rstrip(":").strip()
            value = cells[1].get_text(" ", strip=True)
            key = _META_LABELS.get(label)
            if key and value:
                meta.setdefault(key, value)

    # Strategy 2: look for bold/strong label : value patterns in text nodes.
    page_text = soup.get_text("\n")
    for label, key in _META_LABELS.items():
        if key in meta:
            continue
        m = re.search(rf"{re.escape(label)}\s*[:\-]\s*([^\n<]{{1,120}})", page_text, re.I)
        if m:
            meta[key] = m.group(1).strip()

    return meta

def scrape_job_detail(url: str, bm: BrowserManager) -> dict:
    """Parse a single EthioJobs detail page into a raw_job dict."""
    soup = bm.get_soup(url, wait_selector="h1")
    page_text = soup.get_text("\n")

    # --- Title ---------------------------------------------------------------
    h1 = (soup.find("h1") or soup.select_one("h2.job-title, h2.title"))
    title = clean_title(h1.get_text(" ", strip=True) if h1 else "")

    # --- Company logo --------------------------------------------------------
    logo = ""
    og = soup.find("meta", attrs={"property": "og:image"})
    if og and og.get("content") and "seo/" not in og["content"]:
        logo = og["content"].strip()
    # Fallback: company logo img
    if not logo:
        for img in soup.find_all("img"):
            src = img.get("src", "")
            if re.search(r"logo|employer|company", src, re.I) and src.startswith("http"):
                logo = src
                break

    # --- Company name --------------------------------------------------------
    company_name = ""
    # EthioJobs usually shows the employer in a link or a labelled field.
    for selector in ["a.company-name", "span.company-name", "div.company-name",
                     "a[href*='/company/']", "a[href*='/employer/']",
                     "div.employer-name", "span.employer"]:
        el = soup.select_one(selector)
        if el:
            company_name = el.get_text(" ", strip=True)
            break
    if not company_name:
        m = re.search(r"(?:company|employer|organization)[:\s]+([A-Z][^\n]{2,80})", page_text, re.I)
        if m:
            company_name = m.group(1).strip()
    company_name = company_name or "EthioJobs Employer"

    # Company website (often in an <a title="Website"> or labelled row)
    company_website = ""
    for a in soup.find_all("a", href=True):
        href = a["href"]
        label = (a.get("title", "") + " " + a.get_text(" ", strip=True)).lower()
        if "website" in label and href.startswith("http") and "ethiojobs.net" not in href:
            company_website = href.strip()
            break

    # --- Meta table ----------------------------------------------------------
    meta = _extract_meta_table(soup)

    # --- Job type ------------------------------------------------------------
    job_type_raw = meta.get("job_type", "")
    # Also check for explicit type badges / filter links
    if not job_type_raw:
        for badge in soup.select("span.job-type, a.job-type, div.job-type, span.employment-type"):
            job_type_raw = badge.get_text(" ", strip=True)
            if job_type_raw:
                break
    job_type = map_job_type(job_type_raw)

    # --- Location ------------------------------------------------------------
    location_raw = meta.get("location", "")
    if not location_raw:
        for sel in ["span.location", "div.location", "span.job-location", "a.location"]:
            el = soup.select_one(sel)
            if el:
                location_raw = el.get_text(" ", strip=True)
                break
    # Normalise to a known Ethiopian location if possible.
    if location_raw:
        found = [l for l in ETHIOPIA_LOCATIONS
                 if re.search(rf"\b{re.escape(l)}\b", location_raw, re.I)]
        location = found[0] if found else location_raw.strip()
    else:
        location = location_from_text(page_text)

    # --- Category slug (from URL) --------------------------------------------
    url_parts = [p for p in urlparse(url).path.split("/") if p]
    category_slug = ""
    if len(url_parts) >= 3:
        # /display-job/{id}/{slug} — slug often has category hints
        category_slug = url_parts[-1].lower()

    # Also check breadcrumb / category links on the page.
    cat_links = []
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if "/jobs/" in href and href != "/jobs" and "region" not in href:
            slug_part = href.rstrip("/").split("/jobs/")[-1]
            if slug_part and "/" not in slug_part:
                cat_links.append(slug_part)

    job_field_raw = ", ".join(dict.fromkeys(cat_links)) if cat_links else category_slug

    # --- Dates ---------------------------------------------------------------
    date_posted = ""
    deadline    = ""

    # Prefer explicit meta deadline
    if meta.get("deadline"):
        deadline = parse_any_date(meta["deadline"])

    # Scan page text for deadline labels
    if not deadline:
        for lab in DEADLINE_LABELS:
            m = re.search(rf"{lab}\s*[:\-]?\s*([^\n<]{{1,80}})", page_text, re.I)
            if m:
                d = parse_any_date(m.group(1))
                if d:
                    deadline = d
                    break

    # Try og:article:published_time or a posted-date element
    pub_meta = soup.find("meta", attrs={"property": "article:published_time"})
    if pub_meta and pub_meta.get("content"):
        try:
            date_posted = pub_meta["content"][:10]   # ISO date prefix
        except Exception:
            pass
    if not date_posted:
        for sel in ["span.date-posted", "span.posted-date", "time", "span.created"]:
            el = soup.select_one(sel)
            if el:
                d = parse_any_date(el.get_text())
                if d:
                    date_posted = d
                    break

    if not date_posted:
        date_posted = datetime.now().strftime("%Y-%m-%d")
    if not deadline:
        deadline = (datetime.now() + timedelta(days=30)).strftime("%Y-%m-%d")

    # --- Body text -----------------------------------------------------------
    # Find the richest content container.
    best_el, best_len = None, 0
    for sel in _CONTENT_SELECTORS:
        el = soup.select_one(sel)
        if el:
            txt = el.get_text(" ", strip=True)
            if len(txt) > best_len:
                best_el, best_len = el, len(txt)
        if best_el and best_len > 400:
            break
    if not best_el:
        best_el = soup.find("main") or soup.find("article") or soup.body or soup

    content_copy = BeautifulSoup(str(best_el), "lxml")
    content_text = html_block_to_text(content_copy)
    description, apply_text = _split_description_and_apply(content_text)
    if not description:
        description = content_text

    # --- Qualification + experience ------------------------------------------
    qual_block = ""
    qm = re.search(
        r"(?:^|\n)[ \t]*qualifications?(?:\s*(?:&|and)\s*experience)?[^:\n]{0,30}:?[ \t]*\n"
        r"(.*?)"
        r"(?:\n[ \t]*(?:how\s*to\s*apply|what\s+we\s+offer|application\s+procedure|"
        r"deadline|salary|about\s+the\s+company)\b|\Z)",
        description, re.I | re.S)
    if qm:
        qual_block = qm.group(1).strip()[:1500]

    # Use meta table values first; fall back to body text.
    qual_text   = meta.get("qualification", "") or qual_block or description
    exp_text    = meta.get("experience_raw", "") or qual_block or description
    qualification = extract_qualification(qual_text)
    experience    = extract_experience_band(exp_text)

    # --- Job field -----------------------------------------------------------
    job_field = infer_field(title, description, job_field_raw)

    # --- Apply target --------------------------------------------------------
    apply_email = ""
    apply_url   = ""

    for a in best_el.find_all("a", href=True):
        href = a["href"].strip()
        if href.lower().startswith("mailto:"):
            cand = extract_email(href[7:])
            if cand and _is_real_apply_email(cand):
                apply_email = apply_email or cand
        elif _is_real_apply_url(href):
            apply_url = apply_url or strip_tracking_params(href)

    scan = apply_text or description
    if not apply_email:
        cand = extract_email(scan)
        if cand and _is_real_apply_email(cand):
            apply_email = cand
    if not apply_url:
        for u in URL_PATTERN.findall(scan):
            if _is_real_apply_url(u):
                apply_url = strip_tracking_params(u.rstrip(".,);"))
                break

    salary = meta.get("salary", "") or extract_salary(description)

    return {
        "title":          title,
        "company_name":   company_name,
        "company_url":    "",
        "company_website":company_website,
        "company_address": location,
        "company_logo":   logo,
        "job_type":       job_type,
        "location":       location,
        "job_field":      job_field,
        "job_categories": job_field_raw,
        "date_posted":    date_posted,
        "deadline":       deadline,
        "description":    description,
        "qualification":  qualification,
        "experience":     experience,
        "salary":         salary,
        "apply_email":    apply_email,
        "apply_url":      apply_url,
        "apply_text":     apply_text,
        "job_url":        _norm_job_url(url),
    }

# =============================================================================
#  STEP 3 — DEDUPLICATE + PARAPHRASE + APPLY-RULE GATING
# =============================================================================

def process_job(raw_job, processed_ids, processed_urls, seen_content):
    job_url  = raw_job.get("job_url", "")
    title    = raw_job.get("title", "")
    company  = raw_job.get("company_name", "")
    location = raw_job.get("location", "")

    if not title:
        return "duplicate", None

    job_id = make_job_id(job_url, title, company)

    if job_id in processed_ids or job_url in processed_urls:
        log(C_DIM(f"  Already processed — skipped: {title}"))
        return "duplicate", None

    fingerprint = (title.lower().strip(), company.lower().strip(), location.lower().strip())
    if fingerprint in seen_content:
        log(C_DIM(f"  Duplicate content this run — skipped: {title}"))
        return "duplicate", None
    seen_content.add(fingerprint)

    apply_email = raw_job.get("apply_email", "")
    apply_url   = raw_job.get("apply_url", "")
    qualifies   = bool(apply_email) or bool(apply_url)

    if REQUIRE_PUBLIC_APPLY and not qualifies:
        write_flagged(raw_job,
                      "no public apply email or external URL",
                      raw_job.get("apply_text", "")[:300])
        log(C_RED(f"  FLAGGED (no public apply) — {title}"))
        return "flagged", None

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
        print(C_DIM("  Paraphrasing skipped"))

    application  = apply_email or apply_url
    apply_method = ("description_email" if apply_email
                    else "external_url" if apply_url else "not_found")

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
        "companyUrl":        raw_job.get("company_website", ""),
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
    print(f"  {C_LABEL('Name')}    : {C_VALUE(job.get('companyName','') or C_DIM('—'))}")
    print(f"  {C_LABEL('Website')} : {job.get('companyWebsite','') or C_DIM('—')}")
    print(f"  {C_LABEL('Logo')}    : {job.get('companyLogo','') or C_DIM('— none —')}")
    print()
    print(f"  {C_BLUE('── DESCRIPTION PREVIEW ─────────────────────────────')}")
    print(desc_preview if desc_preview else C_DIM("   — no description —"))
    print(f"  {C_LABEL('Job URL')} : {job.get('jobUrl','')}")
    print(C_DIVIDER())

# =============================================================================
#  WORDPRESS POSTING
# =============================================================================

def _wp_auth_headers():
    token = base64.b64encode(f"{WP_USER}:{WP_PASSWORD}".encode()).decode()
    return {"Authorization": f"Basic {token}", "Content-Type": "application/json"}

def get_or_create_term(taxonomy_url, name):
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

def post_job_to_wordpress(job):
    if not WP_USER or not WP_PASSWORD:
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

    is_email = bool(re.match(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", application))
    is_url_v = bool(re.match(r"^https?://[^\s]+$", application))
    if not (is_email or is_url_v):
        application = ""

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
            "_company_details":    "",
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
            log_.info(f"Posted: '{title}' → WP ID {post.get('id')}")
            return post.get("id"), post.get("link")
        except Exception as e:
            log_.error(f"WP post attempt {attempt+1} failed: {e}")
            if attempt < 2:
                time.sleep(2 ** attempt)
    return None, None

# =============================================================================
#  EXCEL SAVE
# =============================================================================

EXCEL_HEADERS = [
    "Job Title", "Job Type", "Job Qualifications", "Job Experience",
    "Job Location", "Job Field", "Date Posted", "Deadline",
    "Job Description", "Application", "Company URL", "Company Name",
    "Company Logo", "Company Website", "Company Address",
    "Company Details", "Job URL", "Salary Range",
]

def _save_excel(jobs):
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
    log_.info(f"Saved {len(jobs)} rows → {OUTPUT_FILE}")

# =============================================================================
#  MAIN
# =============================================================================

def main():
    start_time = datetime.now()

    # Optional extra listing URLs (category pages, region pages, etc.)
    extra_urls = [u.strip() for u in
                  os.environ.get("ETHIOJOBS_EXTRA_URLS", "").split(",") if u.strip()]

    print()
    print(C_HEADER("=" * 80))
    print(C_HEADER("  ETHIOJOBS (ETHIOPIA) SCRAPER + MISTRAL PARAPHRASE + WORDPRESS POSTING"))
    print(C_HEADER("=" * 80))
    print(f"  Source          : {JOBS_URL}")
    print(f"  Extra URLs      : {extra_urls or 'none'}")
    print(f"  Public-apply    : {'✅ enforced' if REQUIRE_PUBLIC_APPLY else '❌ off'}")
    print(f"  Max new jobs    : {'unlimited' if not MAX_JOBS else MAX_JOBS}")
    print(f"  Max pages       : {MAX_PAGES}")
    print(f"  Paraphrase      : {'✅ enabled' if (ENABLE_PARAPHRASE and MISTRAL_API_KEY) else '❌ disabled'}")
    print(f"  WordPress post  : {'✅ enabled' if (WP_USER and WP_PASSWORD) else '❌ disabled'}")
    print(f"  Excel export    : {'✅ enabled' if _XLSX_AVAILABLE else '❌ (pip install pandas openpyxl)'}")
    print(f"  Playwright      : {'✅ available' if _PW_AVAILABLE else '❌ not installed'}")
    print(f"  Headless mode   : {'yes' if HEADLESS else 'no (browser visible)'}")
    print(f"  Started         : {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(C_HEADER("=" * 80))

    if not _PW_AVAILABLE:
        print(C_RED(
            "\n  ERROR: playwright is not installed.\n"
            "  Run: pip install playwright && playwright install chromium\n"
        ))
        sys.exit(1)

    _init_tracker()
    _init_flagged()
    processed_ids, processed_urls = load_processed_ids()
    print(f"  Tracker loaded: {len(processed_ids)} previously processed job IDs")

    # Build the full list of listing URLs to crawl.
    listing_urls = [JOBS_URL] + extra_urls

    jobs_out    = []
    seen_content = set()
    posted_count = 0
    flagged_count = 0
    dup_count    = 0
    errors       = 0
    scraped      = 0
    all_links    = []

    with BrowserManager(headless=HEADLESS) as bm:
        # --- Collect all job links -------------------------------------------
        for listing_url in listing_urls:
            try:
                links = collect_job_links(bm, listing_url, max_pages=MAX_PAGES)
                for l in links:
                    if l not in all_links:
                        all_links.append(l)
            except Exception as e:
                log(C_RED(f"  ERROR collecting links from {listing_url}: {e}"))

        if not all_links:
            log(C_RED("  No job links found — nothing to do."))
            return
        print(C_GREEN(f"\n  Found {len(all_links)} job detail page(s) to process.\n"))

        # --- Scrape + process each detail page --------------------------------
        for link in all_links:
            if link in processed_urls:
                dup_count += 1
                log(C_DIM(f"  Already processed — skipped: {link}"))
                continue

            try:
                raw_job = scrape_job_detail(link, bm)
                scraped += 1
            except Exception as e:
                errors += 1
                log(C_RED(f"  ERROR scraping {link}: {e}"))
                time.sleep(REQUEST_DELAY)
                continue

            try:
                status, job = process_job(raw_job, processed_ids, processed_urls, seen_content)
            except Exception as e:
                errors += 1
                log(C_RED(f"  ERROR processing '{raw_job.get('title','')}': {e}"))
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
                log(f"\nMAX_JOBS limit ({MAX_JOBS}) reached — stopping.")
                break

            time.sleep(REQUEST_DELAY)

    _save_excel(jobs_out)

    end_time = datetime.now()
    duration = (end_time - start_time).total_seconds() / 60.0
    print()
    print(C_HEADER("=" * 80))
    print(C_HEADER("  SCRAPE COMPLETE"))
    print(C_HEADER("=" * 80))
    print(f"  {C_LABEL('Job links found')}           : {len(all_links)}")
    print(f"  {C_LABEL('Detail pages scraped')}      : {scraped}")
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
        print(f"\n  {C_LABEL('Application links:')}")
        print(f"    External URL : {with_apply - with_email}")
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
