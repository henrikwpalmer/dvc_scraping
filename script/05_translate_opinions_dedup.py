"""
translate_opinions.py  (paragraph-deduplicating version)
---------------------------------------------------------
Translates the scraped court decision texts (opinions) to English.

Reads court_cases_and_opinions.xlsx (produced by scrape_opinions.py) and
writes a SEPARATE file, court_cases_and_opinions_translated.xlsx (never
overwrites the input), with all the same columns plus two new ones:
  - opinion_text_ENGLISH:      translated text (truncated with a marker
                               if over Excel's 32,767-char cell limit)
  - opinion_text_file_ENGLISH: path to a .txt file with the FULL
                               translated text, saved next to the Russian
                               original as {same_name}_EN.txt

The FULL Russian .txt files (from the opinion_text_file column) are used
as the translation source -- NOT the Excel cells -- so decisions that were
truncated in Excel still get translated in full.

------------------------------------------------------------------------
WHAT CHANGED vs. the per-opinion version
------------------------------------------------------------------------
Court decisions are formulaic: measured on this corpus, ~73% of all
paragraph text is a repeat of a paragraph that appears in another
decision (court headers, statute citations, procedural boilerplate,
appeal formulas...). The old script translated every opinion end-to-end,
paying full price for that boilerplate hundreds of times.

This version translates at the PARAGRAPH level with a corpus-wide cache:

1. Every opinion is split into paragraphs (lines). Each paragraph is
   normalized (whitespace collapsed) and hashed.
2. Paragraphs with no Cyrillic at all (blank lines, bare numbers, dates,
   case numbers) are passed through untouched -- no API call.
3. All UNIQUE not-yet-translated paragraphs across the whole corpus are
   collected, PACKED into batches of ~MAX_CHARS (separated by a sentinel
   the translator leaves alone), and translated batch by batch -- so the
   dedup does not cost us the efficiency of batched calls.
4. Each batch's result is split back on the sentinel and VERIFIED
   (segment count must match). If the translator mangled the sentinel,
   that batch automatically falls back to one-call-per-paragraph, so
   correctness never depends on the sentinel surviving.
5. Every translated paragraph is appended to paragraph_cache.jsonl
   immediately (crash/Ctrl+C loses at most the batch in flight; re-runs
   resume). Opinions are then REASSEMBLED line-by-line from the cache,
   preserving the original paragraph structure exactly.

Consequences:
  * Repeated boilerplate is translated exactly once, ever -- including
    across future runs after you scrape more decisions.
  * Identical Russian paragraphs always get the identical English
    rendering (more consistent corpus for analysis).
  * The old translation_checkpoints/ directory is no longer used and can
    be deleted.

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

# Written to a SEPARATE file (never overwrites the input): the scraper
# rebuilds the input workbook from scratch on every run, so saving the
# *_ENGLISH columns into it would let a later scrape silently wipe them.
OUTPUT_FILE = INPUT_FILE.rsplit(".", 1)[0] + "_translated.xlsx"

SOURCE_LANG = "ru"
TARGET_LANG = "en"

TEXT_COLUMN = "opinion_text"                    # source cell column (fallback only)
FILE_COLUMN = "opinion_text_file"               # source .txt path column (preferred)
OUT_TEXT_COLUMN = "opinion_text_ENGLISH"
OUT_FILE_COLUMN = "opinion_text_file_ENGLISH"

# Append-only corpus-wide paragraph translation cache (JSON Lines).
CACHE_FILE = Path("paragraph_cache.jsonl")

TEST_MODE_DEFAULT = True      # True -> only the paragraphs needed to
TEST_OPINIONS = 10            # complete the first TEST_OPINIONS opinions

EXCEL_CELL_LIMIT = 32000      # Excel hard cap is 32,767 chars per cell

# Max characters per translate() call. deep-translator's advertised limit
# is 5000 chars, BUT it sends the text as a GET-request URL parameter and
# Cyrillic URL-encodes to ~5.5 bytes per character ("П" -> "%D0%9F").
# 1200 Cyrillic chars encode to ~6.5 KB, safely under Google's ~8 KB URL
# limit (4500 chars -> ~24 KB -> every request rejected).
MAX_CHARS = 1200

# Sentinel used to pack several paragraphs into one API call. Runs of '@'
# pass through Google Translate unchanged in practice; if a particular
# batch comes back mangled anyway, that batch is retried paragraph-by-
# paragraph, so nothing is ever lost or misaligned.
DELIM = "\n@@@@@\n"
DELIM_RE = re.compile(r"\s*@{3,}\s*")

MAX_RETRIES = 3               # retries per API call (exponential backoff)
DELAY = (0.15, 0.35)          # polite delay range between API calls (s)

CYRILLIC_RE = re.compile(r"[А-Яа-яЁё]")


# ── Paragraph normalization / cache ────────────────────────────────────────────

def normalize_para(p: str) -> str:
    """Canonical form used for cache matching: NBSP -> space, collapse
    whitespace runs, strip. Prevents invisible whitespace differences
    from defeating the dedup."""
    return re.sub(r"\s+", " ", p.replace("\xa0", " ")).strip()


def para_key(norm: str) -> str:
    return hashlib.sha1(norm.encode("utf-8")).hexdigest()[:16]


def load_cache() -> dict:
    """key -> English translation. Tolerates a torn last line (crash
    mid-append) by skipping unparseable lines."""
    cache = {}
    if CACHE_FILE.exists():
        with CACHE_FILE.open(encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    rec = json.loads(line)
                    cache[rec["k"]] = rec["en"]
                except Exception:
                    continue
    return cache


def append_cache(entries: dict, cache: dict):
    """entries: key -> {'ru': ..., 'en': ...}; appended to disk at once."""
    if not entries:
        return
    with CACHE_FILE.open("a", encoding="utf-8") as f:
        for k, rec in entries.items():
            f.write(json.dumps({"k": k, "ru": rec["ru"], "en": rec["en"]},
                               ensure_ascii=False) + "\n")
        f.flush()
    for k, rec in entries.items():
        cache[k] = rec["en"]


def looks_untranslated(text: str) -> bool:
    """True if a supposed English translation is still mostly Cyrillic."""
    if not text:
        return False
    cyr = len(CYRILLIC_RE.findall(text))
    return cyr / max(len(text), 1) > 0.30


# ── Low-level translation with retries ─────────────────────────────────────────

def translate_raw(translator, text: str):
    """One API call with retries/backoff. Returns str or None."""
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            out = translator.translate(text)
            if out:
                return out
        except Exception as exc:
            if attempt < MAX_RETRIES:
                wait = 2 ** attempt + random.uniform(0, 1)
                print(f"      [WARN] call failed ({exc!r}); retry "
                      f"{attempt}/{MAX_RETRIES - 1} in {wait:.0f}s...")
                time.sleep(wait)
            else:
                print(f"      [WARN] call failed after {MAX_RETRIES} "
                      f"attempts ({exc!r}).")
    return None


def translate_one_paragraph(translator, norm: str):
    """Translate a single (possibly oversized) paragraph. Oversized ones
    are split at sentence boundaries; hard-split as a last resort.
    Returns str or None."""
    if len(norm) <= MAX_CHARS:
        return translate_raw(translator, norm)
    sentences = re.split(r"(?<=[.!?])\s+", norm)
    pieces, piece = [], ""
    for s in sentences:
        while len(s) > MAX_CHARS:                 # pathological: hard split
            pieces.append(s[:MAX_CHARS])
            s = s[MAX_CHARS:]
        if len(piece) + len(s) + 1 > MAX_CHARS:
            pieces.append(piece)
            piece = s
        else:
            piece = f"{piece} {s}".strip()
    if piece:
        pieces.append(piece)
    out = []
    for i, p in enumerate(pieces):
        t = translate_raw(translator, p)
        if t is None:
            return None
        out.append(t)
        if i < len(pieces) - 1:
            time.sleep(random.uniform(*DELAY))
    return " ".join(out)


def translate_batch(translator, paras: list):
    """Translate several paragraphs in one call, packed with DELIM.
    Verifies the sentinel survived (segment count must match); if not,
    falls back to per-paragraph calls. Returns {key: {'ru','en'}} for the
    paragraphs that succeeded."""
    done = {}
    out = translate_raw(translator, DELIM.join(paras))
    if out is not None:
        parts = [s.strip() for s in DELIM_RE.split(out)]
        parts = [p for p in parts if p]
        if len(parts) == len(paras):
            for ru, en in zip(paras, parts):
                done[para_key(ru)] = {"ru": ru, "en": en}
            return done
        print(f"      [i] sentinel mangled in this batch "
              f"({len(parts)} segments back for {len(paras)} paragraphs); "
              f"falling back to per-paragraph calls.")
    for i, ru in enumerate(paras):
        t = translate_one_paragraph(translator, ru)
        if t is not None:
            done[para_key(ru)] = {"ru": ru, "en": t}
        if i < len(paras) - 1:
            time.sleep(random.uniform(*DELAY))
    return done


def pack_batches(norms: list):
    """Greedily pack normalized paragraphs into batches whose joined
    length stays <= MAX_CHARS. Oversized paragraphs and ones containing
    the sentinel character run become single-paragraph batches."""
    batches, current, cur_len = [], [], 0
    for norm in norms:
        if len(norm) > MAX_CHARS or DELIM_RE.search(norm):
            if current:
                batches.append(current)
                current, cur_len = [], 0
            batches.append([norm])
            continue
        add = len(norm) + (len(DELIM) if current else 0)
        if cur_len + add > MAX_CHARS:
            batches.append(current)
            current, cur_len = [norm], len(norm)
        else:
            current.append(norm)
            cur_len += add
    if current:
        batches.append(current)
    return batches


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

    # ── Build the work list: unique opinions, keyed by source ────────────────
    # jobs: {opinion_id: {"source_id", "lines", "txt_path" (or None)}}
    # where lines = [(original_line, norm_or_None)]; norm is None for
    # passthrough lines (empty / no Cyrillic).
    jobs = {}
    row_keys = []  # per-row opinion_id (or None), same order as df rows

    def opinion_id(source_id: str) -> str:
        return hashlib.sha1(source_id.encode("utf-8")).hexdigest()[:16]

    def split_lines(text: str):
        lines = []
        for line in text.split("\n"):
            norm = normalize_para(line)
            lines.append((line, norm if CYRILLIC_RE.search(norm) else None))
        return lines

    for _, row in df.iterrows():
        txt_path = row.get(FILE_COLUMN)
        cell_text = row.get(TEXT_COLUMN)
        oid = None

        if isinstance(txt_path, str) and txt_path.strip():
            p = Path(txt_path.strip())
            if p.exists():
                oid = opinion_id(str(p))
                if oid not in jobs:
                    jobs[oid] = {"source_id": str(p),
                                 "lines": split_lines(
                                     p.read_text(encoding="utf-8")),
                                 "txt_path": p}
            else:
                print(f"    [!] Listed .txt not found on disk: {p} -- "
                      f"falling back to the Excel cell text for this row.")

        if oid is None and isinstance(cell_text, str) and cell_text.strip():
            body = re.sub(r"^\[TRUNCATED[^\]]*\]\n", "", cell_text.strip())
            oid = opinion_id(body[:500])
            if oid not in jobs:
                jobs[oid] = {"source_id": "(excel cell)",
                             "lines": split_lines(body), "txt_path": None}

        row_keys.append(oid)

    # ── Corpus-wide unique paragraphs ─────────────────────────────────────────
    cache = load_cache()
    # norm -> key for every Cyrillic paragraph in the corpus, insertion order
    all_paras = {}
    per_opinion_keys = {}
    for oid, job in jobs.items():
        keys = set()
        for _, norm in job["lines"]:
            if norm is not None:
                k = para_key(norm)
                keys.add(k)
                if norm not in all_paras:
                    all_paras[norm] = k
        per_opinion_keys[oid] = keys

    total_occurrences = sum(
        1 for job in jobs.values() for _, n in job["lines"] if n is not None)
    print(f"[i] {len(df)} rows; {len(jobs)} unique opinions; "
          f"{total_occurrences} paragraph occurrences -> "
          f"{len(all_paras)} unique paragraphs "
          f"({len(cache)} already in cache).")

    # ── Translate missing paragraphs (batched, cached, resumable) ─────────────
    if not assemble_only:
        pending_norms = [n for n, k in all_paras.items() if k not in cache]

        if test_mode:
            # Only the paragraphs needed to complete the first
            # TEST_OPINIONS opinions (in row order).
            wanted = set()
            seen = []
            for oid in row_keys:
                if oid and oid not in seen:
                    seen.append(oid)
                if len(seen) >= TEST_OPINIONS:
                    break
            for oid in seen:
                wanted |= per_opinion_keys[oid]
            pending_norms = [n for n in pending_norms
                             if all_paras[n] in wanted]
            print(f"[i] TEST MODE: only the {len(pending_norms)} "
                  f"paragraphs needed for the first {TEST_OPINIONS} "
                  f"opinions. Use --full for everything.")

        batches = pack_batches(pending_norms)
        n_chars = sum(len(n) for n in pending_norms)
        print(f"[i] {len(pending_norms)} paragraphs to translate "
              f"({n_chars} chars) in ~{len(batches)} batched API calls.")

        translator = GoogleTranslator(source=SOURCE_LANG, target=TARGET_LANG)

        # Smoke test: fail fast with a clear message instead of grinding
        # through thousands of batches with per-call errors.
        if batches:
            probe = None
            try:
                probe = translator.translate("Привет, мир")
            except Exception as exc:
                probe = None
                print(f"      [WARN] smoke test raised {exc!r}")
            if not probe or looks_untranslated(probe):
                print("[x] Translator smoke test failed. Nothing will "
                      "translate until this is fixed -- check your "
                      "internet connection / whether translate.google.com "
                      "is reachable, or try 'pip install -U "
                      "deep-translator'. Aborting.")
                sys.exit(1)

        failed_paras = 0
        for i, batch in enumerate(batches, 1):
            print(f"[{i}/{len(batches)}] Translating batch of "
                  f"{len(batch)} paragraph(s) "
                  f"({sum(len(b) for b in batch)} chars)...")
            done = translate_batch(translator, batch)
            failed_paras += len(batch) - len(done)
            append_cache(done, cache)      # checkpointed immediately
            if i < len(batches):
                time.sleep(random.uniform(*DELAY))
        if failed_paras:
            print(f"[!] {failed_paras} paragraph(s) failed and were not "
                  f"cached -- re-run to retry them (everything else is "
                  f"served from cache).")
    else:
        print("[i] --assemble-only: skipping translation, using cache only.")

    # ── Reassemble opinions from the paragraph cache ──────────────────────────
    def assemble_opinion(job):
        """Returns (english_text, complete). Untranslated Cyrillic lines
        are kept in Russian with a marker, and flag the opinion
        incomplete."""
        out, complete = [], True
        for line, norm in job["lines"]:
            if norm is None:
                out.append(line)               # passthrough (no Cyrillic)
                continue
            en = cache.get(para_key(norm))
            if en is None:
                out.append(line)
                complete = False
            else:
                out.append(en)
        return "\n".join(out), complete

    out_texts, out_files = [], []
    written_txt = {}
    n_complete = 0

    for oid in row_keys:
        if oid is None:
            out_texts.append("")
            out_files.append("")
            continue

        if oid not in written_txt:
            translated, complete = assemble_opinion(jobs[oid])
            if not complete:
                # Not fully translated yet (test mode / failures): leave
                # this opinion's outputs empty rather than shipping a
                # half-Russian file.
                written_txt[oid] = None
            else:
                src_path = jobs[oid]["txt_path"]
                if src_path is not None:
                    en_path = src_path.with_name(src_path.stem + "_EN.txt")
                else:
                    Path("opinion_fulltexts").mkdir(exist_ok=True)
                    en_path = (Path("opinion_fulltexts")
                               / f"opinion_{oid}_EN.txt")
                en_path.write_text(translated, encoding="utf-8")
                written_txt[oid] = (str(en_path), translated)

        rec = written_txt[oid]
        if rec is None:
            out_texts.append("")
            out_files.append("")
            continue
        en_path_str, translated = rec
        n_complete += 1
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

    n_files = sum(1 for v in written_txt.values() if v is not None)
    n_rows_with_text = sum(1 for k in row_keys if k is not None)
    print(f"\n✓ Done — '{OUT_TEXT_COLUMN}' and '{OUT_FILE_COLUMN}' columns "
          f"written to {OUTPUT_FILE} (original {INPUT_FILE} left untouched)")
    print(f"  {n_complete} rows have complete translations; "
          f"{n_files} translated .txt files written; "
          f"{len(cache)} paragraphs in cache.")
    if n_complete < n_rows_with_text:
        print(f"  [i] {n_rows_with_text - n_complete} rows still "
              f"untranslated -- re-run (with --full) to continue; "
              f"already-translated paragraphs are never re-translated.")


if __name__ == "__main__":
    main()
