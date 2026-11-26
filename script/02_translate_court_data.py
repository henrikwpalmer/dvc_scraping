"""
translate_court_data.py
-----------------------
Reads the combined ALL_COURTS workbook produced by sudrf_scraper.py
(cases_output_ALL_COURTS.xlsx), translates the Russian free-text columns
to English using deep-translator (free, no API key required), and writes
the results to a second "ENGLISH" tab in the same workbook -- the original
sheet (with the untranslated Russian data) is left untouched.

Rate-limit handling (Google's free endpoint returns TooManyRequests / HTTP 429
if it sees too many calls from one IP):
  - Bundles many short strings into a single request (newline-joined, up to
    MAX_CHARS), so a few thousand judge names / outcomes become a few dozen
    requests instead of thousands.
  - Retries rate-limited requests with exponential backoff.
  - Stops after MAX_CONSECUTIVE_FAILURES failed batches instead of hammering
    the endpoint (which only prolongs the block).
  - Saves successful translations to a JSON cache next to the workbook after
    every batch, so a re-run resumes where it left off.

Usage:
    pip install deep-translator openpyxl
    python translate_court_data.py

    # Or specify a different file:
    python translate_court_data.py my_data.xlsx
"""

import sys
import os
import re
import json
import time
import random

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from deep_translator import GoogleTranslator

# ── Configuration ──────────────────────────────────────────────────────────────

INPUT_FILE  = sys.argv[1] if len(sys.argv) > 1 else "../output/cases_output_ALL_COURTS_luhansk.xlsx"
SOURCE_LANG = "ru"
TARGET_LANG = "en"

# Translations are persisted here so interrupted / rate-limited runs resume.
CACHE_FILE = os.path.splitext(INPUT_FILE)[0] + "_translation_cache.json"

# pandas' DataFrame.to_excel() (used by sudrf_scraper.py) writes to a sheet
# named "Sheet1" by default. If that's not found, we fall back to whichever
# sheet is first in the workbook, so this still works if you've renamed it.
SOURCE_SHEET = "Sheet1"
OUTPUT_SHEET = "ENGLISH"

# Columns to translate, matched against the exact header text written by
# sudrf_scraper.py.
TRANSLATE_COLUMNS = {
    "judge",
    "solution",
}

# Delay between requests (seconds). Requests are now bundled, so there are far
# fewer of them -- a slower, steadier pace avoids tripping the limit.
DELAY = (1.5, 3.0)

# Max characters per translate() call (Google's limit is ~5000)
MAX_CHARS = 4500

# Retry / backoff settings for TooManyRequests
MAX_RETRIES = 5
BACKOFF_BASE = 15            # seconds; doubles each retry: 15, 30, 60, 120, 240
MAX_CONSECUTIVE_FAILURES = 3 # give up (progress saved) after this many dead batches

SEP = "\n"


# ── Styling ────────────────────────────────────────────────────────────────────

HDR_FILL   = PatternFill("solid", start_color="1F4E79")
HDR_FONT   = Font(bold=True, color="FFFFFF", name="Arial", size=10)
DATA_FONT  = Font(name="Arial", size=10)
ALT_FILL   = PatternFill("solid", start_color="EBF3FB")
THIN = Border(
    left=Side(style="thin"),  right=Side(style="thin"),
    top=Side(style="thin"),   bottom=Side(style="thin"),
)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)


# ── Cache ──────────────────────────────────────────────────────────────────────

def load_cache() -> dict:
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, encoding="utf-8") as f:
                cache = json.load(f)
            print(f"Loaded {len(cache)} cached translations from {CACHE_FILE}")
            return cache
        except Exception as exc:
            print(f"[WARN] Could not read cache file ({exc!r}); starting fresh.")
    return {}


def save_cache(cache: dict) -> None:
    tmp = CACHE_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=0)
    os.replace(tmp, CACHE_FILE)


# ── Translation ────────────────────────────────────────────────────────────────

def needs_translation(text: str) -> bool:
    """Skip blanks, dates, URLs, pure numbers, and text with no Cyrillic."""
    if not text or not text.strip():
        return False
    t = text.strip()
    if t.startswith("http"):
        return False
    if re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", t):
        return False
    if re.fullmatch(r"\d{2}:\d{2}", t):
        return False
    if re.fullmatch(r"[\d\-]+", t):
        return False
    if not re.search(r"[А-Яа-яЁё]", t):
        return False
    return True


def is_rate_limit(exc: Exception) -> bool:
    return "TooManyRequests" in type(exc).__name__ or "429" in str(exc)


def translate_with_retry(translator, text: str) -> str:
    """Translate one request, backing off exponentially on rate limits."""
    for attempt in range(MAX_RETRIES):
        try:
            return translator.translate(text)
        except Exception as exc:
            if not is_rate_limit(exc) or attempt == MAX_RETRIES - 1:
                raise
            wait = BACKOFF_BASE * (2 ** attempt) + random.uniform(0, 5)
            print(f"  [429] Rate-limited -- waiting {wait:.0f}s "
                  f"(retry {attempt + 1}/{MAX_RETRIES - 1})…")
            time.sleep(wait)


def make_batches(strings: list) -> list:
    """
    Group strings into newline-joined requests of up to MAX_CHARS.
    Strings containing newlines (which would break the split) or that are
    too long on their own get a request to themselves.
    """
    batches, cur, cur_len = [], [], 0
    for s in strings:
        if SEP in s or len(s) >= MAX_CHARS:
            batches.append([s])
            continue
        if cur and cur_len + len(s) + 1 > MAX_CHARS:
            batches.append(cur)
            cur, cur_len = [], 0
        cur.append(s)
        cur_len += len(s) + 1
    if cur:
        batches.append(cur)
    return batches


def translate_batch(translator, batch: list) -> list:
    if len(batch) == 1:
        s = batch[0]
        if len(s) > MAX_CHARS:
            print(f"  [WARN] String of {len(s)} chars truncated to {MAX_CHARS}.")
        return [translate_with_retry(translator, s[:MAX_CHARS])]

    out = translate_with_retry(translator, SEP.join(batch)) or ""
    results = out.split(SEP)
    if len(results) == len(batch):
        return results

    # Google merged or split lines -- fall back to one request per string.
    print(f"  [i] Line count mismatch ({len(results)} vs {len(batch)}); "
          f"translating this batch one by one.")
    results = []
    for s in batch:
        results.append(translate_with_retry(translator, s))
        time.sleep(random.uniform(*DELAY))
    return results


def build_translation_cache(unique_strings: set) -> dict:
    """
    Translate all unique strings not already cached.
    Returns {original: translated}. Strings that couldn't be translated are
    absent and fall back to the original text when the sheet is written.
    """
    translator = GoogleTranslator(source=SOURCE_LANG, target=TARGET_LANG)
    cache = load_cache()

    candidates = [s for s in unique_strings if needs_translation(s)]
    pending = sorted(s for s in candidates if s not in cache)
    batches = make_batches(pending)

    print(f"\n{len(candidates)} strings need translation: "
          f"{len(candidates) - len(pending)} already cached, "
          f"{len(pending)} to go in {len(batches)} request(s).")

    consecutive_failures = 0
    done = 0
    for b_i, batch in enumerate(batches, 1):
        try:
            results = translate_batch(translator, batch)
            for s, r in zip(batch, results):
                if r and r.strip():
                    cache[s] = r.strip()
            consecutive_failures = 0
            save_cache(cache)
        except Exception as exc:
            consecutive_failures += 1
            print(f"  [WARN] Batch {b_i} ({len(batch)} strings) failed: {exc!r}")
            if consecutive_failures >= MAX_CONSECUTIVE_FAILURES:
                save_cache(cache)
                print(f"\n[!] {consecutive_failures} batches in a row failed. "
                      f"Google is blocking this IP for now.\n"
                      f"    Progress is saved to {CACHE_FILE}.\n"
                      f"    Wait a few hours (or switch network/VPN) and "
                      f"re-run -- it will resume where it stopped.")
                sys.exit(1)

        done += len(batch)
        print(f"  request {b_i}/{len(batches)} -- {done}/{len(pending)} strings")
        if b_i < len(batches):
            time.sleep(random.uniform(*DELAY))

    return cache


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    print(f"Loading: {INPUT_FILE}")
    wb = load_workbook(INPUT_FILE)

    source_sheet = SOURCE_SHEET
    if source_sheet not in wb.sheetnames:
        source_sheet = wb.sheetnames[0]
        print(f"[i] Sheet '{SOURCE_SHEET}' not found -- using first sheet "
              f"'{source_sheet}' instead. Available sheets: {wb.sheetnames}")

    ws_src = wb[source_sheet]

    headers = {}
    for cell in ws_src[1]:
        if cell.value:
            headers[cell.column] = str(cell.value)

    translate_col_indices = {
        col_idx for col_idx, label in headers.items()
        if label in TRANSLATE_COLUMNS
    }

    if not translate_col_indices:
        print(f"[!] None of {sorted(TRANSLATE_COLUMNS)} were found in the "
              f"header row ({list(headers.values())}). Nothing will be "
              f"translated -- the ENGLISH tab will just be a copy. Check "
              f"that TRANSLATE_COLUMNS matches your actual column headers.")
    else:
        print(f"Columns to translate: "
              f"{[headers[i] for i in sorted(translate_col_indices)]}")

    # ── Collect all unique strings that need translation ───────────────────────
    unique_strings = set()
    data_rows = list(ws_src.iter_rows(min_row=2, values_only=True))

    for row in data_rows:
        for col_idx in translate_col_indices:
            val = row[col_idx - 1]
            if val and str(val).strip():
                unique_strings.add(str(val))

    print(f"Unique non-empty values in translatable columns: {len(unique_strings)}")

    # ── Translate ──────────────────────────────────────────────────────────────
    cache = build_translation_cache(unique_strings)

    # ── Remove existing ENGLISH sheet if re-running ───────────────────────────
    if OUTPUT_SHEET in wb.sheetnames:
        del wb[OUTPUT_SHEET]

    ws_out = wb.create_sheet(OUTPUT_SHEET)

    for cell in ws_src[1]:
        new_cell = ws_out.cell(row=1, column=cell.column, value=cell.value)
        new_cell.fill      = HDR_FILL
        new_cell.font      = HDR_FONT
        new_cell.border    = THIN
        new_cell.alignment = CENTER

    ws_out.freeze_panes = "C2"

    for row_idx, row in enumerate(data_rows, 2):
        is_alt = row_idx % 2 == 0
        for col_idx, label in headers.items():
            original = row[col_idx - 1]
            val = str(original) if original is not None else ""

            if col_idx in translate_col_indices:
                val = cache.get(val, val)

            cell = ws_out.cell(row=row_idx, column=col_idx, value=val)
            cell.font      = DATA_FONT
            cell.border    = THIN
            cell.alignment = LEFT
            if is_alt:
                cell.fill = ALT_FILL

    for col_idx, label in headers.items():
        col_letter = get_column_letter(col_idx)
        max_len = max(
            len(str(ws_out.cell(row=r, column=col_idx).value or ""))
            for r in range(1, len(data_rows) + 2)
        )
        ws_out.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 55)

    ws_out.auto_filter.ref = ws_out.dimensions

    wb.save(INPUT_FILE)
    untranslated = sum(1 for s in unique_strings
                       if needs_translation(s) and s not in cache)
    print(f"\n✓ Done — '{OUTPUT_SHEET}' tab written to {INPUT_FILE}")
    print(f"  {len(data_rows)} rows across {len(translate_col_indices)} columns"
          + (f" ({untranslated} strings left in Russian)" if untranslated else "."))


if __name__ == "__main__":
    main()
