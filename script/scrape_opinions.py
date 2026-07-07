#!/usr/bin/env python3
"""
sudrf.ru judicial-decision (opinion text) scraper
==================================================

Reads the deduplicated combined case workbook (ENGLISH_deduped.xlsx),
filters to rows where "judicial_acts_link" is present, fetches each
decision page, extracts the full text of the ruling, and writes
everything back out to court_cases_and_opinions.xlsx with two new columns:
  - opinion_text: the decision text inline (truncated with a marker if
    over Excel's 32,767-char cell limit)
  - opinion_text_file: path to a .txt file containing the FULL text,
    saved for every fetched decision (named by case number + short hash)
    in opinion_fulltexts/.

------------------------------------------------------------------------
CHECKPOINTING / RESUME (safe to interrupt at any time)
------------------------------------------------------------------------
Every successfully fetched decision is immediately saved as its own small
JSON file in CHECKPOINT_DIR (default: ./opinion_checkpoints/), keyed by a
hash of the decision URL. This means:

  - You can Ctrl+C, close your laptop, or drop off wi-fi at ANY moment
    and lose at most the single request that was in flight.
  - On the next run, links that already have a checkpoint file are
    skipped automatically -- the script just picks up where it left off.
  - Duplicate links (same decision referenced by multiple rows) are only
    fetched once.
  - Permanent failures are also checkpointed (with an error marker) so a
    repeatedly-failing link doesn't stall every future run. Delete its
    checkpoint file (or run with --retry-failures) to try it again.

To rebuild the final Excel from existing checkpoints WITHOUT fetching
anything new (e.g. offline), run:  python scrape_opinions.py --assemble-only

------------------------------------------------------------------------
USAGE
------------------------------------------------------------------------
    python scrape_opinions.py               # test mode: first 10 links only
    python scrape_opinions.py --full        # all links
    python scrape_opinions.py --assemble-only   # just build the Excel from checkpoints
    python scrape_opinions.py --retry-failures  # re-attempt links that errored before

The final Excel is (re)assembled from checkpoints at the end of EVERY run,
so even a test run produces a valid court_cases_and_opinions.xlsx with
whatever has been fetched so far.

------------------------------------------------------------------------
POLITENESS (same approach as sudrf_scraper.py)
------------------------------------------------------------------------
- Randomized 4-9s delays between requests, longer break every ~15.
- Persistent session, realistic browser headers, Referer chaining.
- Session warm-up per domain (decisions may span many court domains).
- Exponential backoff on 403/429/errors; sequential requests only.
"""

import csv
import hashlib
import json
import random
import re
import sys
import time
from pathlib import Path
from typing import Optional
from urllib.parse import urlparse

import requests
from bs4 import BeautifulSoup

# =========================================================================
# CONFIG
# =========================================================================

INPUT_XLSX = "../temp/ENGLISH_deduped.xlsx"
OUTPUT_XLSX = "../output/court_cases_and_opinions.xlsx"
LINK_COLUMN = "judicial_acts_link"
OPINION_COLUMN = "opinion_text"
OPINION_FILE_COLUMN = "opinion_text_file"

CHECKPOINT_DIR = Path("opinion_checkpoints")

TEST_MODE_DEFAULT = True     # True -> only the first TEST_LINKS links
TEST_LINKS = 10

# EVERY fetched opinion is also saved as its own .txt file (full,
# untruncated text) in LONG_TEXT_DIR, named after the case number plus a
# short link-hash suffix, with its path recorded in the
# opinion_text_file column. The opinion_text cell in Excel additionally
# holds the text inline for convenience -- truncated with a marker if it
# exceeds Excel's 32,767-character cell limit.
EXCEL_CELL_LIMIT = 32000
LONG_TEXT_DIR = Path("opinion_fulltexts")

# Delays (same philosophy as the case-list scraper)
DELAY_MIN, DELAY_MAX = 4.0, 9.0
LONG_BREAK_EVERY = 15
LONG_BREAK_MIN, LONG_BREAK_MAX = 20.0, 40.0
MAX_RETRIES = 4

DEFAULT_ENCODING = "windows-1251"

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 "
    "(KHTML, like Gecko) Version/17.4 Safari/605.1.15",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0",
]


# =========================================================================
# Networking (mirrors sudrf_scraper.py)
# =========================================================================

def build_session() -> requests.Session:
    s = requests.Session()
    s.headers.update({
        "User-Agent": random.choice(USER_AGENTS),
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


def decode_response(resp: requests.Response) -> str:
    encoding = None
    m = re.search(r"charset=([\w-]+)", resp.headers.get("Content-Type", ""), re.IGNORECASE)
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
            resp = session.get(url, headers=headers, timeout=30)
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
    return None


def polite_sleep(request_count: int):
    if request_count > 0 and request_count % LONG_BREAK_EVERY == 0:
        pause = random.uniform(LONG_BREAK_MIN, LONG_BREAK_MAX)
        print(f"    ...taking a longer break ({pause:.1f}s)...")
    else:
        pause = random.uniform(DELAY_MIN, DELAY_MAX)
    time.sleep(pause)


def warm_up_domain(session: requests.Session, url: str,
                   warmed: set) -> Optional[str]:
    """
    Visit the domain root + bare search module once per domain before the
    first deep request to it (decision links can span many court domains).
    Returns a referer URL to use. Mirrors the warm-up in sudrf_scraper.py,
    which proved necessary to avoid being bounced by these sites.
    """
    parsed = urlparse(url)
    domain = parsed.netloc
    referer = f"{parsed.scheme}://{domain}/modules.php?name=sud_delo&srv_num=1"
    if domain in warmed:
        return referer
    root = f"{parsed.scheme}://{domain}/"
    print(f"    [i] Warming up new domain: {domain}")
    fetch(session, root, referer=None)
    time.sleep(random.uniform(2.0, 4.0))
    fetch(session, referer, referer=root)
    time.sleep(random.uniform(2.0, 4.0))
    warmed.add(domain)
    return referer


# =========================================================================
# Decision-text extraction
# =========================================================================

def extract_opinion_text(html: str) -> Optional[str]:
    """
    On sudrf.ru decision pages (name_op=doc), the ruling text lives inside
    a <span> within <div id="content">, structured as <p> paragraphs.
    Verified against a real decision page. Falls back progressively if a
    court's template differs slightly.
    """
    soup = BeautifulSoup(html, "lxml")

    content = soup.find(id="content")
    if content is not None:
        span = content.find("span")
        if span is not None:
            text = span.get_text("\n", strip=True)
            if len(text) > 100:  # sanity check: a real ruling, not a stub
                return text
        # Fallback within #content: take all of it minus nav titles
        text = content.get_text("\n", strip=True)
        if len(text) > 100:
            return text

    # Last-resort fallback: largest run of <p> tags on the page
    best_text = ""
    for candidate in soup.find_all(["span", "div"]):
        paras = candidate.find_all("p", recursive=False)
        if len(paras) >= 5:
            text = candidate.get_text("\n", strip=True)
            if len(text) > len(best_text):
                best_text = text
    return best_text if len(best_text) > 100 else None


def looks_like_error_page(html: str) -> bool:
    """Detect the 'specify search criteria' bounce or an empty stub."""
    return "Необходимо задать критерии поиска" in html


# =========================================================================
# Checkpointing
# =========================================================================

def checkpoint_path(url: str) -> Path:
    key = hashlib.sha1(url.encode("utf-8")).hexdigest()[:16]
    return CHECKPOINT_DIR / f"{key}.json"


def load_checkpoint(url: str) -> Optional[dict]:
    p = checkpoint_path(url)
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            return None
    return None


def save_checkpoint(url: str, opinion_text: Optional[str], error: Optional[str]):
    CHECKPOINT_DIR.mkdir(exist_ok=True)
    data = {
        "url": url,
        "opinion_text": opinion_text,
        "error": error,
        "fetched_at": time.strftime("%Y-%m-%d %H:%M:%S"),
    }
    checkpoint_path(url).write_text(
        json.dumps(data, ensure_ascii=False), encoding="utf-8"
    )


# =========================================================================
# Main workflow
# =========================================================================

def load_cases() -> "tuple":
    import pandas as pd
    df = pd.read_excel(INPUT_XLSX, dtype=str)
    if LINK_COLUMN not in df.columns:
        print(f"[x] Column '{LINK_COLUMN}' not found in {INPUT_XLSX}. "
              f"Columns present: {list(df.columns)}")
        sys.exit(1)
    mask = df[LINK_COLUMN].notna() & (df[LINK_COLUMN].str.strip() != "")
    with_links = df[mask]
    print(f"[i] Loaded {len(df)} rows from {INPUT_XLSX}; "
          f"{len(with_links)} have a {LINK_COLUMN}.")
    return df, with_links


def scrape(links: list, retry_failures: bool):
    session = build_session()
    warmed_domains = set()
    request_count = 0
    fetched, skipped, failed = 0, 0, 0

    for i, url in enumerate(links, 1):
        existing = load_checkpoint(url)
        if existing is not None:
            if existing.get("error") and retry_failures:
                pass  # fall through and re-fetch
            else:
                skipped += 1
                continue

        print(f"[{i}/{len(links)}] Fetching decision: {url}")
        referer = warm_up_domain(session, url, warmed_domains)
        request_count += 1
        html = fetch(session, url, referer)

        if html is None:
            save_checkpoint(url, None, error="fetch_failed")
            failed += 1
        elif looks_like_error_page(html):
            save_checkpoint(url, None, error="bounced_to_search_form")
            print("    [!] Bounced to search form -- checkpointed as error.")
            failed += 1
        else:
            text = extract_opinion_text(html)
            if text:
                save_checkpoint(url, text, error=None)
                print(f"    [+] Extracted {len(text)} chars of opinion text.")
                fetched += 1
            else:
                save_checkpoint(url, None, error="no_text_found")
                print("    [!] Page fetched but no opinion text found -- "
                      "checkpointed as error. (The decision may not be "
                      "published yet, or the page layout differs.)")
                failed += 1

        if i < len(links):
            polite_sleep(request_count)

    print(f"\n[i] Scrape pass complete: {fetched} fetched, "
          f"{skipped} already checkpointed (skipped), {failed} failed.")


def sanitize_filename(text: str, max_len: int = 60) -> str:
    """Turn a case number like '2-595/2026 ~ М-480/2026' into a safe filename stem."""
    stem = re.sub(r"[^\w\-]+", "_", str(text)).strip("_")
    return stem[:max_len] or "case"


def assemble(df_all, df_links):
    """
    Build the final Excel: all original rows/columns, plus:
      - opinion_text: the decision text (truncated with a marker if it
        exceeds Excel's ~32k cell character limit)
      - opinion_text_file: relative path to a .txt file holding the FULL,
        untruncated text -- written for EVERY fetched opinion, named after
        the row's case number (plus a short link-hash suffix so two cases
        with similar numbers, or one case with multiple decision documents,
        never collide/overwrite).
    Rows without a link, or whose link hasn't been fetched yet, get empty
    values in both columns.
    """
    import pandas as pd

    opinions = {}
    n_errors = 0
    if CHECKPOINT_DIR.exists():
        for p in CHECKPOINT_DIR.glob("*.json"):
            try:
                data = json.loads(p.read_text(encoding="utf-8"))
            except Exception:
                continue
            if data.get("opinion_text"):
                opinions[data["url"]] = data["opinion_text"]
            elif data.get("error"):
                n_errors += 1

    df_out = df_all.copy()

    # A link can appear on multiple rows; write its .txt only once, named
    # after the FIRST row that references it.
    link_to_txt_path = {}

    def process_row(row):
        link = row.get(LINK_COLUMN)
        if not isinstance(link, str) or not link.strip():
            return pd.Series({OPINION_COLUMN: "", OPINION_FILE_COLUMN: ""})
        link = link.strip()
        text = opinions.get(link, "")
        if not text:
            return pd.Series({OPINION_COLUMN: "", OPINION_FILE_COLUMN: ""})

        if link not in link_to_txt_path:
            LONG_TEXT_DIR.mkdir(exist_ok=True)
            case_stem = sanitize_filename(row.get("case_number", "case"))
            link_hash = hashlib.sha1(link.encode("utf-8")).hexdigest()[:8]
            txt_path = LONG_TEXT_DIR / f"{case_stem}_{link_hash}.txt"
            txt_path.write_text(text, encoding="utf-8")
            link_to_txt_path[link] = str(txt_path)
        txt_path_str = link_to_txt_path[link]

        if len(text) > EXCEL_CELL_LIMIT:
            cell_text = (f"[TRUNCATED -- full text in {txt_path_str}]\n"
                         + text[:EXCEL_CELL_LIMIT - 100])
        else:
            cell_text = text
        return pd.Series({OPINION_COLUMN: cell_text,
                          OPINION_FILE_COLUMN: txt_path_str})

    df_out[[OPINION_COLUMN, OPINION_FILE_COLUMN]] = df_out.apply(
        process_row, axis=1
    )
    df_out.to_excel(OUTPUT_XLSX, index=False)

    n_filled = (df_out[OPINION_COLUMN].str.len() > 0).sum()
    print(f"[✓] Wrote {OUTPUT_XLSX}: {len(df_out)} rows total, "
          f"{n_filled} with opinion text "
          f"({len(df_links)} rows have links; {n_errors} checkpointed errors).")
    print(f"[✓] Full opinion texts saved as .txt files in {LONG_TEXT_DIR}/ "
          f"({len(link_to_txt_path)} unique decisions).")


if __name__ == "__main__":
    test_mode = TEST_MODE_DEFAULT
    if "--full" in sys.argv:
        test_mode = False
    if "--test" in sys.argv:
        test_mode = True
    retry_failures = "--retry-failures" in sys.argv
    assemble_only = "--assemble-only" in sys.argv

    df_all, df_links = load_cases()

    # Unique links, preserving order of first appearance
    seen = set()
    links = []
    for link in df_links[LINK_COLUMN]:
        link = link.strip()
        if link and link not in seen:
            seen.add(link)
            links.append(link)
    print(f"[i] {len(links)} unique decision links.")

    if not assemble_only:
        if test_mode:
            links_to_fetch = links[:TEST_LINKS]
            print(f"[i] TEST MODE: only fetching the first "
                  f"{len(links_to_fetch)} links. Run with --full for all.")
        else:
            links_to_fetch = links
        scrape(links_to_fetch, retry_failures=retry_failures)
    else:
        print("[i] --assemble-only: skipping all fetching.")

    assemble(df_all, df_links)
    print("Done.")
