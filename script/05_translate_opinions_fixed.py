"""
translate_opinions.py
---------------------
Translates the scraped court decision texts (opinions) to English.

Reads court_cases_and_opinions.xlsx (produced by scrape_opinions.py) and
writes a SEPARATE file, court_cases_and_opinions_translated.xlsx (never
overwrites the input -- see OUTPUT_FILE below for why), with all the same
columns plus two new ones:
  - opinion_text_ENGLISH:      translated text (truncated with a marker
                               if over Excel's 32,767-char cell limit)
  - opinion_text_file_ENGLISH: path to a .txt file with the FULL
                               translated text (never truncated), saved
                               next to the Russian original as
                               {same_name}_EN.txt

The FULL Russian .txt files (from the opinion_text_file column) are used
as the translation source -- NOT the Excel cells -- so decisions that were
truncated in Excel still get translated in full.

------------------------------------------------------------------------
EFFICIENCY / ROBUSTNESS FEATURES (translation is the slow step)
------------------------------------------------------------------------
1. DISK-CACHED, RESUMABLE: every translated opinion is checkpointed to
   translation_checkpoints/ the moment it finishes. Ctrl+C / wi-fi drop
   loses at most the opinion in flight; re-running skips everything
   already translated. (Same pattern as scrape_opinions.py.)
2. DEDUPLICATED: opinions shared by multiple rows (same source .txt) are
   translated once.
3. PARAGRAPH-AWARE CHUNKING: Google's free endpoint caps at ~5000 chars
   per call. Long decisions are split into chunks at PARAGRAPH boundaries
   (never mid-sentence), translated chunk by chunk, and re-joined --
   so nothing is silently cut off (the old script simply truncated at
   4500 chars!) and sentence context is preserved for quality.
4. The Excel is (re)assembled from checkpoints at the end of every run,
   so even a partial/test run produces a valid updated workbook.

Usage:
    pip install deep-translator openpyxl pandas
    python translate_opinions.py             # test mode: first 10 opinions
    python translate_opinions.py --full      # all opinions
    python translate_opinions.py --assemble-only   # rebuild Excel from cache only
    python translate_opinions.py my_file.xlsx --full   # different input file
"""

import hashlib
import json
import random
import re
import sys
import time
from pathlib import Path

from deep_translator import GoogleTranslator

# ── Configuration ──────────────────────────────────────────────────────────────

_cli_files = [a for a in sys.argv[1:] if not a.startswith("--")]
INPUT_FILE = _cli_files[0] if _cli_files else "../output/court_cases_and_opinions.xlsx"

# Written to a SEPARATE file from INPUT_FILE (never overwrites it). This
# matters because scrape_opinions.py rebuilds court_cases_and_opinions.xlsx
# from scratch on every run (assemble() re-reads checkpoints + the input
# workbook and rewrites the whole file) -- if this script saved back into
# that same file, re-running/continuing the scraper later would silently
# wipe out the *_ENGLISH columns with no warning. Reading from one file
# and writing to another means you can safely keep scraping more decisions
# at any time and just re-run this script afterward -- already-translated
# opinions are still served from translation_checkpoints/, so nothing gets
# re-translated.
OUTPUT_FILE = INPUT_FILE.rsplit(".", 1)[0] + "_translated.xlsx"

SOURCE_LANG = "ru"
TARGET_LANG = "en"

TEXT_COLUMN = "opinion_text"                    # source cell column (fallback only)
FILE_COLUMN = "opinion_text_file"               # source .txt path column (preferred)
OUT_TEXT_COLUMN = "opinion_text_ENGLISH"
OUT_FILE_COLUMN = "opinion_text_file_ENGLISH"

CHECKPOINT_DIR = Path("translation_checkpoints")

TEST_MODE_DEFAULT = True      # True -> only translate the first TEST_OPINIONS
TEST_OPINIONS = 10

EXCEL_CELL_LIMIT = 32000      # Excel hard cap is 32,767 chars per cell

# Max characters per translate() call. deep-translator's advertised limit
# is 5000 chars, BUT it sends the text as a GET-request URL parameter and
# Cyrillic URL-encodes to ~5.5 bytes per character ("П" -> "%D0%9F").
# A 4500-char Russian chunk becomes a ~24 KB URL, which Google's servers
# reject outright (URL limit is ~8 KB) -- so EVERY chunk errored and the
# script fell back to keeping the Russian. 1200 Cyrillic chars encode to
# ~6.5 KB, safely under the limit.
MAX_CHARS = 1200

# Retries per chunk before giving up on it (with exponential backoff).
MAX_RETRIES = 3

# Polite delay range between API calls (seconds). Long decisions need
# multiple calls (one per chunk), each separated by this delay.
DELAY = (0.15, 0.35)


# ── Chunking ───────────────────────────────────────────────────────────────────

def split_into_chunks(text: str, max_chars: int = MAX_CHARS) -> list:
    """
    Split text into chunks of at most max_chars, breaking ONLY at
    paragraph (newline) boundaries so sentences are never cut mid-way.
    A single paragraph longer than max_chars (rare -- real decisions
    average ~350 chars/paragraph) is split at sentence boundaries, and as
    an absolute last resort hard-split.
    """
    paragraphs = text.split("\n")
    chunks, current = [], ""

    def flush():
        nonlocal current
        if current:
            chunks.append(current)
            current = ""

    for para in paragraphs:
        if len(para) > max_chars:
            # Oversized paragraph: split at sentence ends.
            flush()
            sentences = re.split(r"(?<=[.!?])\s+", para)
            piece = ""
            for s in sentences:
                while len(s) > max_chars:          # pathological: hard split
                    chunks.append(s[:max_chars])
                    s = s[max_chars:]
                if len(piece) + len(s) + 1 > max_chars:
                    chunks.append(piece)
                    piece = s
                else:
                    piece = f"{piece} {s}".strip()
            if piece:
                chunks.append(piece)
        elif len(current) + len(para) + 1 > max_chars:
            flush()
            current = para
        else:
            current = f"{current}\n{para}" if current else para
    flush()
    return chunks


# ── Translation with disk cache ────────────────────────────────────────────────

def checkpoint_path(key: str) -> Path:
    return CHECKPOINT_DIR / f"{key}.json"


def opinion_key(source_id: str) -> str:
    """Stable cache key for an opinion (hash of its .txt path or text)."""
    return hashlib.sha1(source_id.encode("utf-8")).hexdigest()[:16]


def load_cached_translation(key: str):
    p = checkpoint_path(key)
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8")).get("translated")
        except Exception:
            return None
    return None


def save_cached_translation(key: str, source_id: str, translated: str):
    CHECKPOINT_DIR.mkdir(exist_ok=True)
    checkpoint_path(key).write_text(
        json.dumps({"source": source_id, "translated": translated,
                    "translated_at": time.strftime("%Y-%m-%d %H:%M:%S")},
                   ensure_ascii=False),
        encoding="utf-8",
    )


def translate_opinion(translator: GoogleTranslator, text: str):
    """
    Translate one full opinion, chunk by chunk, preserving paragraphs.
    Returns (translated_text, ok) -- ok is False if ANY chunk failed
    (after retries), so the caller can decide NOT to checkpoint it.
    """
    chunks = split_into_chunks(text)
    out, ok = [], True
    for j, chunk in enumerate(chunks, 1):
        translated = None
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                translated = translator.translate(chunk)
                break
            except Exception as exc:
                if attempt < MAX_RETRIES:
                    wait = 2 ** attempt + random.uniform(0, 1)
                    print(f"      [WARN] chunk {j}/{len(chunks)} failed "
                          f"({exc!r}); retry {attempt}/{MAX_RETRIES - 1} "
                          f"in {wait:.0f}s...")
                    time.sleep(wait)
                else:
                    print(f"      [WARN] chunk {j}/{len(chunks)} failed "
                          f"after {MAX_RETRIES} attempts ({exc!r}); "
                          f"keeping original Russian for that chunk.")
        if translated:
            out.append(translated)
        else:
            out.append(chunk)
            ok = False
        if j < len(chunks):
            time.sleep(random.uniform(*DELAY))
    return "\n".join(out), ok


def looks_untranslated(text: str) -> bool:
    """True if a supposed English translation is still mostly Cyrillic."""
    if not text:
        return False
    cyr = len(re.findall(r"[А-Яа-яЁё]", text))
    return cyr / max(len(text), 1) > 0.30


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    import pandas as pd

    test_mode = TEST_MODE_DEFAULT
    if "--full" in sys.argv:
        test_mode = False
    if "--test" in sys.argv:
        test_mode = True
    assemble_only = "--assemble-only" in sys.argv

    print(f"Loading: {INPUT_FILE}")
    df = pd.read_excel(INPUT_FILE, dtype=str)

    if FILE_COLUMN not in df.columns and TEXT_COLUMN not in df.columns:
        print(f"[x] Neither '{FILE_COLUMN}' nor '{TEXT_COLUMN}' found in "
              f"{INPUT_FILE}. Columns: {list(df.columns)}")
        sys.exit(1)

    # ── Build the work list: unique opinions, keyed by their source .txt ──────
    # Prefer the full .txt file (never truncated); fall back to the Excel
    # cell text only if the file is missing.
    # jobs: {key: {"source_id", "text", "txt_path" (or None)}}
    jobs = {}
    row_keys = []  # per-row key (or None), same order as df rows

    for _, row in df.iterrows():
        txt_path = row.get(FILE_COLUMN)
        cell_text = row.get(TEXT_COLUMN)
        key = None

        if isinstance(txt_path, str) and txt_path.strip():
            p = Path(txt_path.strip())
            if p.exists():
                key = opinion_key(str(p))
                if key not in jobs:
                    jobs[key] = {"source_id": str(p),
                                 "text": p.read_text(encoding="utf-8"),
                                 "txt_path": p}
            else:
                print(f"    [!] Listed .txt not found on disk: {p} -- "
                      f"falling back to the Excel cell text for this row.")

        if key is None and isinstance(cell_text, str) and cell_text.strip():
            # Fallback: cell text (may carry a truncation marker; strip it).
            body = re.sub(r"^\[TRUNCATED[^\]]*\]\n", "", cell_text.strip())
            key = opinion_key(body[:500])  # key on content prefix
            if key not in jobs:
                jobs[key] = {"source_id": "(excel cell)", "text": body,
                             "txt_path": None}

        row_keys.append(key)

    print(f"[i] {len(df)} rows; {sum(k is not None for k in row_keys)} with "
          f"opinion text; {len(jobs)} unique opinions to translate.")

    # ── Purge poisoned checkpoints ─────────────────────────────────────────────
    # Earlier runs cached failed "translations" that are still Russian
    # (every chunk errored -> fallback kept the original -> got checkpointed).
    # Those must be invalidated or they'd be skipped forever.
    purged = 0
    for key in jobs:
        cached = load_cached_translation(key)
        if cached is not None and looks_untranslated(cached):
            checkpoint_path(key).unlink()
            purged += 1
    if purged:
        print(f"[i] Purged {purged} cached checkpoint(s) that were still "
              f"Russian (failed earlier runs); they will be re-translated.")

    # ── Translate (with cache/resume) ──────────────────────────────────────────
    if not assemble_only:
        pending = [k for k in jobs if load_cached_translation(k) is None]
        already = len(jobs) - len(pending)
        if test_mode:
            pending = pending[:TEST_OPINIONS]
            print(f"[i] TEST MODE: translating at most {TEST_OPINIONS} "
                  f"opinions this run. Use --full for everything.")
        print(f"[i] {already} already cached (skipped); "
              f"{len(pending)} to translate now.")

        translator = GoogleTranslator(source=SOURCE_LANG, target=TARGET_LANG)

        # Smoke test: fail fast with a clear message instead of grinding
        # through every opinion with per-chunk errors.
        if pending:
            try:
                probe = translator.translate("Привет, мир")
                assert probe and not looks_untranslated(probe)
            except Exception as exc:
                print(f"[x] Translator smoke test failed ({exc!r}). "
                      f"Nothing will translate until this is fixed -- check "
                      f"your internet connection / whether "
                      f"translate.google.com is reachable, or try "
                      f"'pip install -U deep-translator'. Aborting.")
                sys.exit(1)

        failed = 0
        for i, key in enumerate(pending, 1):
            job = jobs[key]
            n_chunks = len(split_into_chunks(job["text"]))
            print(f"[{i}/{len(pending)}] Translating {job['source_id']} "
                  f"({len(job['text'])} chars, {n_chunks} chunk(s))...")
            translated, ok = translate_opinion(translator, job["text"])
            if ok:
                save_cached_translation(key, job["source_id"], translated)
            else:
                failed += 1
                print(f"      [!] Not checkpointed (had failed chunks); "
                      f"will be retried on the next run.")
            time.sleep(random.uniform(*DELAY))
        if failed:
            print(f"[!] {failed} opinion(s) had failed chunks and were not "
                  f"saved -- re-run to retry them.")
    else:
        print("[i] --assemble-only: skipping translation, using cache only.")

    # ── Write translated .txt files + assemble the Excel columns ──────────────
    out_texts, out_files = [], []
    written_txt = {}

    for key in row_keys:
        if key is None:
            out_texts.append("")
            out_files.append("")
            continue
        translated = load_cached_translation(key)
        if translated is None:
            out_texts.append("")   # not translated yet (e.g. test mode)
            out_files.append("")
            continue

        # Write the full translated .txt (once per unique opinion).
        if key not in written_txt:
            src_path = jobs[key]["txt_path"]
            if src_path is not None:
                en_path = src_path.with_name(src_path.stem + "_EN.txt")
            else:
                Path("opinion_fulltexts").mkdir(exist_ok=True)
                en_path = Path("opinion_fulltexts") / f"opinion_{key}_EN.txt"
            en_path.write_text(translated, encoding="utf-8")
            written_txt[key] = str(en_path)
        en_path_str = written_txt[key]

        if len(translated) > EXCEL_CELL_LIMIT:
            cell = (f"[TRUNCATED -- full translation in {en_path_str}]\n"
                    + translated[:EXCEL_CELL_LIMIT - 100])
        else:
            cell = translated
        out_texts.append(cell)
        out_files.append(en_path_str)

    df[OUT_TEXT_COLUMN] = out_texts
    df[OUT_FILE_COLUMN] = out_files
    df.to_excel(OUTPUT_FILE, index=False)

    n_done = sum(1 for t in out_texts if t)
    print(f"\n✓ Done — '{OUT_TEXT_COLUMN}' and '{OUT_FILE_COLUMN}' columns "
          f"written to {OUTPUT_FILE} (original {INPUT_FILE} left untouched)")
    print(f"  {n_done} rows have translations; "
          f"{len(written_txt)} translated .txt files written.")
    if n_done < sum(1 for k in row_keys if k is not None):
        print(f"  [i] {sum(1 for k in row_keys if k is not None) - n_done} "
              f"rows still untranslated -- re-run (with --full) to continue; "
              f"already-translated opinions will be skipped automatically.")


if __name__ == "__main__":
    main()
