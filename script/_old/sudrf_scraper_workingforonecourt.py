#!/usr/bin/env python3
"""
sudrf.ru case-list scraper
===========================

Scrapes case number / case-page link / date of receipt / judge from the
search-results pages of a sudrf.ru-style court site (the "GAS Pravosudie"
template used by essentially all Russian court websites), across all
paginated result pages, and writes them to CSV/XLSX.

Because these sites share one template but courts sometimes reorder or
rename columns slightly, the parser does NOT hard-code column positions.
Instead, for each page it:
  1. Finds the results table.
  2. Reads the header row to figure out which column holds the date and
     which holds the judge (by matching Russian keywords).
  3. Extracts the case number + link from the anchor tag in the case-number
     column.

This makes the same script reusable across the "several dozen" court sites
you mentioned -- you just change BASE_URL (and, if needed, tweak the
KEYWORD lists below if a court uses unusual header wording).

------------------------------------------------------------------------
USAGE
------------------------------------------------------------------------
1. Edit BASE_URL below to the first results page (page=1) for a given
   court/search query.
2. Run in DEBUG mode first (default): it will only crawl 2 result pages
   and print out what it found, plus save the raw HTML of page 1 to
   debug_page1.html so you can visually confirm the scraper is reading
   the right table/columns.
3. Once it looks correct, set DEBUG = False (or pass --full on the CLI)
   to crawl every page. The script auto-computes the number of pages from
   the "results found" counter on the page (falls back to a manual value
   if that can't be found).

------------------------------------------------------------------------
BEING A GOOD NETWORK CITIZEN
------------------------------------------------------------------------
- Randomized delay (default 4-9s) between every request, plus a longer
  "coffee break" pause every ~15 requests, so traffic doesn't look like a
  metronomic bot.
- A single persistent `requests.Session` (like a real browser tab) with a
  realistic header set (Accept-Language ru-RU, Accept, Accept-Encoding,
  Connection: keep-alive) rather than a bare python-requests UA.
- Referer is set to the previous page, mimicking normal click-through
  navigation instead of jumping to arbitrary URLs cold.
- Exponential backoff + retry on transient errors / 403 / 429, instead of
  hammering the server.
- Sequential, single-threaded requests only -- no concurrency/parallel
  hammering.
- Respects robots.txt (checked once at startup; the script will warn, not
  silently ignore, if disallowed).

This is deliberately conservative. If you're running this across dozens of
sites, please also consider spacing out *which* site you hit and when
(e.g. don't run 30 of these back to back with no break), and check each
site's robots.txt / terms yourself -- this script checks but you should
still use judgment.
"""

import csv
import random
import re
import sys
import time
import urllib.robotparser
from dataclasses import dataclass, asdict
from typing import List, Optional, Tuple
from urllib.parse import urljoin, urlparse, parse_qs, urlencode, urlunparse

import requests
from bs4 import BeautifulSoup

# =========================================================================
# CONFIG -- edit these per court site
# =========================================================================

# The FIRST results page (page=1) URL. Paste the full URL you copy from the
# browser after running a search on that court's site.
BASE_URL = (
    "https://amv--dnr.sudrf.ru/modules.php?name=sud_delo&srv_num=1&name_op=r&page=1&vnkod=93RS0012&srv_num=1&name_op=r&vnkod=93RS0012&delo_id=1540005&case_type=0&new=0&G1_PARTS__NAMESS=&g1_case__CASE_NUMBERSS=&g1_case__JUDICIAL_UIDSS=&delo_table=g1_case&g1_case__ENTRY_DATE1D=&g1_case__ENTRY_DATE2D=&lawbookarticles%5B0%5D=%CE+%EF%F0%E8%E7%ED%E0%ED%E8%E8+%E4%E2%E8%E6%E8%EC%EE%E9+%E2%E5%F9%E8+%E1%E5%E7%F5%EE%E7%FF%E9%ED%EE%E9+%E8+%EF%F0%E8%E7%ED%E0%ED%E8%E8+%EF%F0%E0%E2%E0+%EC%F3%ED%E8%F6%E8%EF%E0%EB%FC%ED%EE%E9+%F1%EE%E1%F1%F2%E2%E5%ED%ED%EE%F1%F2%E8+%ED%E0+%E1%E5%E7%F5%EE%E7%FF%E9%ED%F3%FE+%ED%E5%E4%E2%E8%E6%E8%EC%F3%FE+%E2%E5%F9%FC&G1_CASE__JUDGE=&g1_case__RESULT_DATE1D=&g1_case__RESULT_DATE2D=&G1_CASE__RESULT=&G1_CASE__BUILDING_ID=&G1_CASE__COURT_STRUCT=&G1_EVENT__EVENT_NAME=&G1_EVENT__EVENT_DATEDD=&G1_PARTS__PARTS_TYPE=&G1_PARTS__INN_STRSS=&G1_PARTS__KPP_STRSS=&G1_PARTS__OGRN_STRSS=&G1_PARTS__OGRNIP_STRSS=&G1_RKN_ACCESS_RESTRICTION__RKN_REASON=&g1_rkn_access_restriction__RKN_RESTRICT_URLSS=&g1_requirement__ACCESSION_DATE1D=&g1_requirement__ACCESSION_DATE2D=&G1_REQUIREMENT__CATEGORY=&g1_requirement__ESSENCESS=&g1_requirement__JOIN_END_DATE1D=&g1_requirement__JOIN_END_DATE2D=&G1_REQUIREMENT__PUBLICATION_ID=&G1_DOCUMENT__PUBL_DATE1D=&G1_DOCUMENT__PUBL_DATE2D=&G1_CASE__VALIDITY_DATE1D=&G1_CASE__VALIDITY_DATE2D=&G1_ORDER_INFO__ORDER_DATE1D=&G1_ORDER_INFO__ORDER_DATE2D=&G1_ORDER_INFO__ORDER_NUMSS=&G1_ORDER_INFO__EXTERNALKEYSS=&G1_ORDER_INFO__STATE_ID=&G1_ORDER_INFO__RECIP_ID=&Submit=Find"
)

RESULTS_PER_PAGE = 25          # used to compute total pages from result count
DEBUG = True                   # True -> only crawl DEBUG_PAGES pages
DEBUG_PAGES = 2
MANUAL_PAGE_COUNT_FALLBACK = 10  # used only if auto-detection of total results fails

OUTPUT_PREFIX = "../output/cases_output"   # final files will be {OUTPUT_PREFIX}_{CourtName}.csv/.xlsx

# If True, the auto-detected court name (pulled from the page's <title>) is
# transliterated to Latin characters for the filename (e.g. "Амвросиевский"
# -> "Amvrosievsky"). If False, the original Cyrillic is kept (modern
# filesystems handle Unicode filenames fine).
TRANSLITERATE_COURT_NAME = True

# Basic Russian Cyrillic -> Latin transliteration table (a common
# BGN/PCGN-style mapping). Good enough for filenames; not meant to be a
# precise/official transliteration standard.
_CYRILLIC_TO_LATIN = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "e",
    "ж": "zh", "з": "z", "и": "i", "й": "i", "к": "k", "л": "l", "м": "m",
    "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t", "у": "u",
    "ф": "f", "х": "kh", "ц": "ts", "ч": "ch", "ш": "sh", "щ": "shch",
    "ъ": "", "ы": "y", "ь": "", "э": "e", "ю": "yu", "я": "ya",
}


def transliterate(text: str) -> str:
    out = []
    for ch in text:
        lower = ch.lower()
        if lower in _CYRILLIC_TO_LATIN:
            piece = _CYRILLIC_TO_LATIN[lower]
            if ch.isupper() and piece:
                piece = piece[0].upper() + piece[1:]
            out.append(piece)
        else:
            out.append(ch)
    return "".join(out)


def extract_court_name(html: str) -> str:
    """
    Pull a short, filename-friendly court name from the page's <title>,
    which on sudrf.ru sites looks like:
        "Амвросиевский районный суд Донецкой Народной Республики"
    We take just the first word (the court's own adjectival name).
    """
    soup = BeautifulSoup(html, "lxml")
    title_tag = soup.find("title")
    if not title_tag:
        return "UnknownCourt"
    title = title_tag.get_text(strip=True)
    m = re.match(r"[A-Za-zА-Яа-яЁё\-]+", title)
    name = m.group(0) if m else title
    if TRANSLITERATE_COURT_NAME:
        name = transliterate(name)
    name = re.sub(r"[^\w\-]", "_", name).strip("_")
    return name or "UnknownCourt"

# Delay ranges (seconds) -- tweak if you want to be even more conservative
DELAY_MIN, DELAY_MAX = 4.0, 9.0
LONG_BREAK_EVERY = 15           # take a longer pause every N requests
LONG_BREAK_MIN, LONG_BREAK_MAX = 20.0, 40.0

# Keywords used to auto-detect which column is which (case-insensitive,
# matched against the <th> header text). Add synonyms here if a particular
# court's site uses different wording.
DATE_KEYWORDS = ["дата поступ", "поступлен", "дата регистрации"]
JUDGE_KEYWORDS = ["судья", "судья-докладчик"]
CASE_NUM_KEYWORDS = ["номер дела", "№ дела", "дело"]
DECISION_DATE_KEYWORDS = ["дата решения"]
SOLUTION_KEYWORDS = ["решение"]  # note: doesn't match "дата решениЯ" (genitive) above, only "решение" (nominative)
ENTRY_INTO_FORCE_KEYWORDS = ["вступления в законную силу", "законную силу"]
JUDICIAL_ACTS_KEYWORDS = ["судебные акты", "судебный акт"]

# A small pool of realistic desktop User-Agents to rotate between runs
# (kept stable *within* a run -- real browsers don't change UA mid-session).
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 "
    "(KHTML, like Gecko) Version/17.4 Safari/605.1.15",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0",
]

MAX_RETRIES = 4

# sudrf.ru sites are encoded in Windows-1251 (Cyrillic), and many of them
# don't declare charset in the HTTP Content-Type header (only in a <meta>
# tag), so requests' auto-detection can't be trusted. We sniff it properly
# below, falling back to this default.
DEFAULT_ENCODING = "windows-1251"

# Known result-table id used by the standard sudrf.ru template. Checked
# first (fast path); if a particular court's page doesn't have it, we fall
# back to the header-keyword heuristic further down.
KNOWN_TABLE_IDS = ["tablcont"]


# =========================================================================
# Data model
# =========================================================================

@dataclass
class CaseRecord:
    case_number: str
    case_link: str
    date_of_receipt: str
    judge: str
    decision_date: str
    solution: str
    entry_into_force_date: str
    judicial_acts_link: str
    source_page: int


# =========================================================================
# Networking helpers
# =========================================================================

def build_session() -> requests.Session:
    s = requests.Session()
    ua = random.choice(USER_AGENTS)
    s.headers.update({
        "User-Agent": ua,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,"
                  "image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection": "keep-alive",
        "Upgrade-Insecure-Requests": "1",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "same-origin",
        "Sec-Fetch-User": "?1",
        "Cache-Control": "max-age=0",
    })
    return s


def check_robots(base_url: str) -> bool:
    """Return True if crawling this path is allowed (or unknown), else False."""
    parsed = urlparse(base_url)
    robots_url = f"{parsed.scheme}://{parsed.netloc}/robots.txt"
    rp = urllib.robotparser.RobotFileParser()
    try:
        rp.set_url(robots_url)
        rp.read()
        allowed = rp.can_fetch("*", base_url)
        if not allowed:
            print(f"[!] robots.txt at {robots_url} appears to DISALLOW this path.")
        return allowed
    except Exception as e:
        print(f"[i] Could not read robots.txt ({e}); proceeding cautiously.")
        return True


def polite_sleep(request_count: int):
    if request_count % LONG_BREAK_EVERY == 0:
        pause = random.uniform(LONG_BREAK_MIN, LONG_BREAK_MAX)
        print(f"    ...taking a longer break ({pause:.1f}s) to look human...")
    else:
        pause = random.uniform(DELAY_MIN, DELAY_MAX)
    time.sleep(pause)


def decode_response(resp: requests.Response) -> str:
    """
    sudrf.ru pages are Windows-1251. Some don't set charset in the HTTP
    header at all, so we check (in order): HTTP Content-Type header,
    <meta charset=...> sniffed from the raw bytes, then fall back to
    DEFAULT_ENCODING.
    """
    encoding = None
    content_type = resp.headers.get("Content-Type", "")
    m = re.search(r"charset=([\w-]+)", content_type, re.IGNORECASE)
    if m:
        encoding = m.group(1)
    else:
        m = re.search(rb'charset=[\'"]?([\w-]+)', resp.content[:2000], re.IGNORECASE)
        if m:
            encoding = m.group(1).decode("ascii", errors="ignore")
    if not encoding:
        encoding = DEFAULT_ENCODING
    try:
        return resp.content.decode(encoding, errors="replace")
    except LookupError:
        return resp.content.decode(DEFAULT_ENCODING, errors="replace")


def fetch(session: requests.Session, url: str, referer: Optional[str]) -> Optional[str]:
    headers = {}
    if referer:
        headers["Referer"] = referer
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = session.get(url, headers=headers, timeout=25)
            if resp.status_code == 200:
                return decode_response(resp)
            elif resp.status_code in (403, 429):
                wait = (2 ** attempt) + random.uniform(1, 4)
                print(f"    [!] HTTP {resp.status_code}. Backing off {wait:.1f}s "
                      f"(attempt {attempt}/{MAX_RETRIES})...")
                time.sleep(wait)
            else:
                print(f"    [!] HTTP {resp.status_code} for {url}")
                return None
        except requests.RequestException as e:
            wait = (2 ** attempt) + random.uniform(1, 4)
            print(f"    [!] Request error: {e}. Retrying in {wait:.1f}s "
                  f"(attempt {attempt}/{MAX_RETRIES})...")
            time.sleep(wait)
    print(f"    [x] Giving up on {url} after {MAX_RETRIES} attempts.")
    return None


def set_page_param(url: str, page: int) -> str:
    """
    Replace/add the page= query param regardless of its current value.

    IMPORTANT: this does a plain string/regex substitution rather than
    routing through urllib.parse.parse_qs()/urlencode(). Those functions
    decode percent-encoded bytes as UTF-8 and re-encode them, which
    silently corrupts any parameter that was percent-encoded in a
    different charset -- and sudrf.ru search filters (e.g. lawbookarticles,
    which encodes a Cyrillic law-article description) are percent-encoded
    in Windows-1251, not UTF-8. Round-tripping through parse_qs/urlencode
    turns every such byte into a mangled "%EF%BF%BD" replacement-character
    sequence, which the server then fails to recognize as valid search
    criteria at all. A regex substitution never decodes anything, so every
    other byte in the URL is passed through byte-for-byte untouched.
    """
    if re.search(r"[?&]page=\d*", url):
        return re.sub(r"([?&])page=\d*", rf"\g<1>page={page}", url, count=1)
    sep = "&" if "?" in url else "?"
    return f"{url}{sep}page={page}"


# =========================================================================
# Parsing helpers
# =========================================================================

def warm_up_session(session: requests.Session, base_url: str) -> Optional[str]:
    """
    Visit the site's root and the bare search module (no deep query params)
    before jumping straight into the deep results URL. Many sudrf.ru courts
    reset/bounce a cold direct request to page=N&vnkod=...&delo_id=... back
    to a blank search form if there's no prior session/cookie -- this
    mimics a normal user landing on the site and running a search, rather
    than teleporting straight to page 3 of results.
    Returns the referer URL to use for the first real request.
    """
    parsed = urlparse(base_url)
    root_url = f"{parsed.scheme}://{parsed.netloc}/"
    search_form_url = f"{parsed.scheme}://{parsed.netloc}/modules.php?name=sud_delo&srv_num=1"

    print(f"[i] Warming up session: visiting {root_url}")
    fetch(session, root_url, referer=None)
    time.sleep(random.uniform(2.0, 4.0))

    print(f"[i] Warming up session: visiting {search_form_url}")
    fetch(session, search_form_url, referer=root_url)
    time.sleep(random.uniform(2.0, 4.0))

    return search_form_url


def looks_like_blank_search_form(table) -> bool:
    """
    Detect the 'bounced back to blank form' failure mode: a results-shaped
    table with only 1-2 rows, none of which contain a case-number-like
    link (href with case_id=), usually means the session wasn't
    established and the site served the default search page instead of
    real results.
    """
    if table is None:
        return True
    rows = table.find_all("tr")
    if len(rows) > 2:
        return False
    for row in rows:
        a = row.find("a", href=True)
        if a and "case_id=" in a["href"]:
            return False
    return True


def find_results_table(soup: BeautifulSoup):
    """
    First try the known table id(s) used by the standard sudrf.ru template
    (fast, exact path). If that's not present on a given court's page,
    fall back to a heuristic: the <table> whose header row best matches
    date/judge/case-number keywords.
    """
    for table_id in KNOWN_TABLE_IDS:
        t = soup.find("table", id=table_id)
        if t is not None and len(t.find_all("tr")) > 1:
            return t

    candidate_tables = soup.find_all("table")
    best = None
    best_score = -1
    for t in candidate_tables:
        header_cells = t.find_all("th")
        header_text = " ".join(c.get_text(" ", strip=True).lower() for c in header_cells)
        rows = t.find_all("tr")
        score = 0
        if any(k in header_text for k in DATE_KEYWORDS):
            score += 2
        if any(k in header_text for k in JUDGE_KEYWORDS):
            score += 2
        if any(k in header_text for k in CASE_NUM_KEYWORDS):
            score += 2
        score += min(len(rows), 30) * 0.1  # prefer tables with more rows
        if score > best_score:
            best_score = score
            best = t
    return best


def map_columns(table) -> dict:
    """
    Return a dict of column indices keyed by field name, based on matching
    the header row's text against the KEYWORDS lists above.
    """
    header_row = table.find("tr")
    mapping = {
        "date": None,
        "judge": None,
        "case_num": None,
        "decision_date": None,
        "solution": None,
        "entry_into_force": None,
        "judicial_acts": None,
    }
    if not header_row:
        return mapping
    cells = header_row.find_all(["th", "td"])
    for idx, cell in enumerate(cells):
        text = cell.get_text(" ", strip=True).lower()
        if mapping["date"] is None and any(k in text for k in DATE_KEYWORDS):
            mapping["date"] = idx
        if mapping["judge"] is None and any(k in text for k in JUDGE_KEYWORDS):
            mapping["judge"] = idx
        if mapping["case_num"] is None and any(k in text for k in CASE_NUM_KEYWORDS):
            mapping["case_num"] = idx
        if mapping["decision_date"] is None and any(k in text for k in DECISION_DATE_KEYWORDS):
            mapping["decision_date"] = idx
        if mapping["solution"] is None and any(k in text for k in SOLUTION_KEYWORDS):
            mapping["solution"] = idx
        if mapping["entry_into_force"] is None and any(k in text for k in ENTRY_INTO_FORCE_KEYWORDS):
            mapping["entry_into_force"] = idx
        if mapping["judicial_acts"] is None and any(k in text for k in JUDICIAL_ACTS_KEYWORDS):
            mapping["judicial_acts"] = idx
    return mapping


def parse_results_page(html: str, page_url: str, page_num: int) -> List[CaseRecord]:
    soup = BeautifulSoup(html, "lxml")
    table = find_results_table(soup)
    records: List[CaseRecord] = []

    if table is None:
        print(f"    [!] Could not locate a results table on page {page_num}.")
        return records

    if looks_like_blank_search_form(table):
        print(f"    [!!!] Page {page_num} looks like a BLANK/RESET search form, "
              f"not real results (no case_id links found in the table). "
              f"This usually means the session wasn't established correctly, "
              f"or the site blocked/redirected this request. Check the saved "
              f"debug HTML and verify the page actually shows a case list "
              f"when opened in a normal browser at this exact URL.")
        return records

    col_map = map_columns(table)
    rows = table.find_all("tr")[1:]  # skip header row

    for row in rows:
        cells = row.find_all("td")
        if not cells:
            continue

        # Case number + link: look for the first <a> in the row (or in the
        # designated case_num column if we found one).
        link_tag = None
        if col_map["case_num"] is not None and col_map["case_num"] < len(cells):
            link_tag = cells[col_map["case_num"]].find("a")
        if link_tag is None:
            link_tag = row.find("a")
        if link_tag is None:
            continue  # not a data row (e.g. a spacer/nav row)

        case_number = link_tag.get_text(" ", strip=True)
        case_link = urljoin(page_url, link_tag.get("href", ""))

        date_val = ""
        if col_map["date"] is not None and col_map["date"] < len(cells):
            date_val = cells[col_map["date"]].get_text(" ", strip=True)

        judge_val = ""
        if col_map["judge"] is not None and col_map["judge"] < len(cells):
            judge_val = cells[col_map["judge"]].get_text(" ", strip=True)

        decision_date_val = ""
        if col_map["decision_date"] is not None and col_map["decision_date"] < len(cells):
            decision_date_val = cells[col_map["decision_date"]].get_text(" ", strip=True)

        solution_val = ""
        if col_map["solution"] is not None and col_map["solution"] < len(cells):
            solution_val = cells[col_map["solution"]].get_text(" ", strip=True)

        entry_into_force_val = ""
        if col_map["entry_into_force"] is not None and col_map["entry_into_force"] < len(cells):
            entry_into_force_val = cells[col_map["entry_into_force"]].get_text(" ", strip=True)

        # "Судебные акты" (judicial decisions) column: only save the link
        # (if one is present), not any surrounding text.
        judicial_acts_link_val = ""
        if col_map["judicial_acts"] is not None and col_map["judicial_acts"] < len(cells):
            acts_link_tag = cells[col_map["judicial_acts"]].find("a", href=True)
            if acts_link_tag is not None:
                judicial_acts_link_val = urljoin(page_url, acts_link_tag["href"])

        # Skip obvious non-data rows (e.g. empty case number)
        if not case_number:
            continue

        records.append(CaseRecord(
            case_number=case_number,
            case_link=case_link,
            date_of_receipt=date_val,
            judge=judge_val,
            decision_date=decision_date_val,
            solution=solution_val,
            entry_into_force_date=entry_into_force_val,
            judicial_acts_link=judicial_acts_link_val,
            source_page=page_num,
        ))

    return records


def detect_total_pages(html: str) -> Optional[int]:
    """
    Try to find a 'N results found' style counter on the page and compute
    ceil(N / RESULTS_PER_PAGE). Returns None if not found.
    """
    text = BeautifulSoup(html, "lxml").get_text(" ", strip=True)
    # Observed real phrasing: "Всего по запросу найдено — 228. На странице
    # записи с 1 по 25." Also tolerate other variants like "Найдено дел: 123".
    m = re.search(r"найдено\D{0,15}(\d+)", text, re.IGNORECASE)
    if m:
        total = int(m.group(1))
        pages = -(-total // RESULTS_PER_PAGE)  # ceil division
        return max(pages, 1)
    return None


# =========================================================================
# Main crawl
# =========================================================================

def crawl(base_url: str, debug: bool) -> Tuple[List[CaseRecord], str]:
    session = build_session()
    check_robots(base_url)

    all_records: List[CaseRecord] = []
    request_count = 0

    referer = warm_up_session(session, base_url)
    request_count += 2

    page1_url = set_page_param(base_url, 1)
    print(f"[1] Fetching page 1: {page1_url}")
    html = fetch(session, page1_url, referer)
    request_count += 1
    if html is None:
        print("[x] Failed to fetch page 1 -- aborting.")
        return all_records, "UnknownCourt"

    court_name = extract_court_name(html)
    print(f"    [i] Detected court name: {court_name}")

    with open("debug_page1.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("    [i] Saved raw HTML of page 1 to debug_page1.html for inspection.")

    total_pages = detect_total_pages(html)
    if total_pages is None:
        total_pages = MANUAL_PAGE_COUNT_FALLBACK
        print(f"    [i] Could not auto-detect result count; "
              f"falling back to MANUAL_PAGE_COUNT_FALLBACK = {total_pages}.")
    else:
        print(f"    [i] Auto-detected {total_pages} total page(s) "
              f"(at {RESULTS_PER_PAGE}/page).")

    if debug:
        total_pages = min(total_pages, DEBUG_PAGES)
        print(f"    [i] DEBUG mode: limiting crawl to {total_pages} page(s).")

    page1_records = parse_results_page(html, page1_url, 1)
    print(f"    [+] Parsed {len(page1_records)} case(s) from page 1.")
    all_records.extend(page1_records)
    referer = page1_url

    for page_num in range(2, total_pages + 1):
        polite_sleep(request_count)
        page_url = set_page_param(base_url, page_num)
        print(f"[{page_num}] Fetching page {page_num}: {page_url}")
        html = fetch(session, page_url, referer)
        request_count += 1
        if html is None:
            print(f"    [!] Skipping page {page_num} (fetch failed).")
            continue
        records = parse_results_page(html, page_url, page_num)
        print(f"    [+] Parsed {len(records)} case(s) from page {page_num}.")
        all_records.extend(records)
        referer = page_url

    return all_records, court_name


def save_output(records: List[CaseRecord], csv_path: str, xlsx_path: str):
    if not records:
        print("[!] No records to save.")
        return

    fieldnames = [
        "case_number", "case_link", "date_of_receipt", "judge",
        "decision_date", "solution", "entry_into_force_date",
        "judicial_acts_link", "source_page",
    ]

    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for r in records:
            writer.writerow(asdict(r))
    print(f"[✓] Saved {len(records)} records to {csv_path}")

    try:
        import pandas as pd
        df = pd.DataFrame([asdict(r) for r in records])
        df.to_excel(xlsx_path, index=False)
        print(f"[✓] Saved {len(records)} records to {xlsx_path}")
    except ImportError:
        print("[i] pandas/openpyxl not installed -- skipped .xlsx output "
              "(CSV was still saved). Run: pip install pandas openpyxl")


if __name__ == "__main__":
    debug_mode = DEBUG
    if "--full" in sys.argv:
        debug_mode = False
    if "--debug" in sys.argv:
        debug_mode = True

    print(f"=== sudrf case scraper === (debug={debug_mode})")
    results, court_name = crawl(BASE_URL, debug=debug_mode)
    csv_path = f"{OUTPUT_PREFIX}_{court_name}.csv"
    xlsx_path = f"{OUTPUT_PREFIX}_{court_name}.xlsx"
    save_output(results, csv_path, xlsx_path)
    print("Done.")
