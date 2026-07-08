"""
translate_opinions.py
---------------------
Translates the scraped court decision texts (opinions) to English.

Reads court_cases_and_opinions.xlsx (produced by scrape_opinions.py),
translates each opinion, and writes back the same file with two new
columns:
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
INPUT_FILE = _cli_files[0] if _cli_files else "court_cases_and_opinions.xlsx"

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

# Max characters per translate() call (Google's limit is ~5000).
MAX_CHARS = 4500

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


def translate_opinion(translator: GoogleTranslator, text: str) -> str:
    """Translate one full opinion, chunk by chunk, preserving paragraphs."""
    chunks = split_into_chunks(text)
    out = []
    for j, chunk in enumerate(chunks, 1):
        try:
            translated = translator.translate(chunk)
            out.append(translated if translated else chunk)
        except Exception as exc:
            print(f"      [WARN] chunk {j}/{len(chunks)} failed ({exc!r}); "
                  f"keeping original Russian for that chunk.")
            out.append(chunk)
        if j < len(chunks):
            time.sleep(random.uniform(*DELAY))
    return "\n".join(out)


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
        for i, key in enumerate(pending, 1):
            job = jobs[key]
            n_chunks = len(split_into_chunks(job["text"]))
            print(f"[{i}/{len(pending)}] Translating {job['source_id']} "
                  f"({len(job['text'])} chars, {n_chunks} chunk(s))...")
            translated = translate_opinion(translator, job["text"])
            save_cached_translation(key, job["source_id"], translated)
            time.sleep(random.uniform(*DELAY))
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
    df.to_excel(INPUT_FILE, index=False)

    n_done = sum(1 for t in out_texts if t)
    print(f"\n✓ Done — '{OUT_TEXT_COLUMN}' and '{OUT_FILE_COLUMN}' columns "
          f"written to {INPUT_FILE}")
    print(f"  {n_done} rows have translations; "
          f"{len(written_txt)} translated .txt files written.")
    if n_done < sum(1 for k in row_keys if k is not None):
        print(f"  [i] {sum(1 for k in row_keys if k is not None) - n_done} "
              f"rows still untranslated -- re-run (with --full) to continue; "
              f"already-translated opinions will be skipped automatically.")


if __name__ == "__main__":
    main()
