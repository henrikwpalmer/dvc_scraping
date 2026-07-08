# sudrf.ru Court Case Scraper & Translator

A pipeline of four scripts for collecting case data and judicial decisions
from sudrf.ru-style Russian court websites (the standardized "GAS
Pravosudie" template used across essentially all Russian courts), and
translating the results to English.

The pipeline runs in four stages:

```
1. 01_sudrf_scraper.py        court URLs (Excel) ──► per-court case lists + combined workbook
2. 02_translate_court_data.py combined workbook   ──► + English tab (judge/solution)
3. 03_dedupe_and_analyze_cases.py deduped workbook  ──► + Dedupes the data based on case number and outputs a little summary table
4. 04_scrape_opinions.py      case list w/ links  ──► + full decision text per case
5. 05_translate_opinions.py   decisions           ──► + English translation of each decision
```

Each stage's output feeds the next, but every script can also be re-run
independently and safely at any time — see [Re-running / resuming](#re-running--resuming-safely)
below.

---

## Requirements

```bash
pip install requests beautifulsoup4 lxml openpyxl pandas deep-translator
```

Python 3.8+.

---

## Stage 1 — `sudrf_scraper.py`

Scrapes the case-list search-results pages for one or more courts: case
number, link to the case page, date of receipt, judge, decision date,
solution/outcome, date of entry into force, and a link to the judicial
decision document (when available).

### Setup

1. Generate a starter input file:
   ```bash
   python sudrf_scraper.py --make-template
   ```
   This writes `court_urls_template.xlsx` with two columns:
   - `base_url` — a court's full search-results URL, **with your filter
     already applied** (i.e. exactly the URL you'd copy from the browser
     address bar after running a search on that court's site, including
     `page=1`)
   - `notes` — optional free-text description of the filter, for your own
     reference only (not used in the crawl logic)

2. Fill in one row per court, save as `court_urls.xlsx` (or point
   `INPUT_XLSX` at wherever you saved it).

### Running

```bash
python sudrf_scraper.py            # TEST MODE: first 2 result pages per court
python sudrf_scraper.py --full     # every result page, every court
python sudrf_scraper.py --debug    # force test mode even if DEBUG=False in the file
```

Test mode is the default (`DEBUG = True` in the config section) so you
can sanity-check the parsing against a new court before committing to a
full crawl of it.

### Output

- **One file per court**: `cases_output_{CourtName}.csv` and `.xlsx`,
  auto-named by scraping the court's own name from its page `<title>`
  (transliterated to Latin by default — see `TRANSLITERATE_COURT_NAME`).
- **One combined workbook**: `cases_output_ALL_COURTS.xlsx`, with every
  court's rows stacked together and tagged with `court_name`,
  `source_base_url`, and `notes` columns.
- `debug_page1.html` — the raw HTML of the first court's first page,
  saved every run for troubleshooting (see below).

### Columns produced

| Column | Notes |
|---|---|
| `case_number` | |
| `case_link` | link to the case's own page |
| `date_of_receipt` | |
| `judge` | |
| `decision_date` | blank if not yet decided |
| `solution` | outcome text, blank if not yet decided |
| `entry_into_force_date` | blank if not yet in force |
| `judicial_acts_link` | link to the decision document, blank if none published |
| `source_page` | which result page this row came from |
| `court_name` / `source_base_url` / `notes` | combined-workbook only |

### If something looks wrong

Every run saves `debug_page1.html` — open it and compare to what a real
browser shows at that URL. Two failure modes we've hit and fixed, in case
you see similar symptoms on a new court:

- **Bounced to a blank search form** (page title
  `Поиск информации по делам`, no real rows): the site needs a "warm-up"
  visit to establish a session before deep-linking into results — this is
  handled automatically, but if a new court still shows this, its
  template may need a small tweak.
- **"Необходимо задать критерии поиска" (please specify search
  criteria)**: usually means a filter parameter got corrupted somewhere.
  These sites percent-encode Cyrillic filter values in **Windows-1251**,
  not UTF-8 — if you ever modify the URL-handling code, avoid routing it
  through `urllib.parse.parse_qs`/`urlencode`, which decode/re-encode as
  UTF-8 and silently mangle those bytes.

### Being polite to the server

- Randomized 4–9s delay between requests, longer 20–40s break every ~15
  requests.
- Extra 15–30s pause between finishing one court and starting the next.
- One persistent session per run with realistic browser headers and
  `Referer` chaining between pages (mimics normal click-through
  navigation).
- Exponential backoff + retry on 403/429/timeouts; sequential requests
  only, no concurrency.
- Checks `robots.txt` and warns (doesn't silently ignore) if disallowed.

---

## Stage 2 — `translate_court_data.py`

Translates the `judge` and `solution` columns of the combined workbook to
English (these are the only two columns with free-form Russian prose —
everything else is a date, URL, case number, or already-Latin metadata).

```bash
python translate_court_data.py                                    # default input file
python translate_court_data.py path/to/cases_output_ALL_COURTS.xlsx
```

**Output**: a new file, `{input_name}_translated.xlsx` — the original
input file is never modified (see [Re-running / resuming](#re-running--resuming-safely)
for why that matters). The output has the same sheet as the input plus a
second `ENGLISH` tab with the same rows, translated.

Deduplicates before translating (each unique judge name / outcome string
is translated once, no matter how many rows it appears on).

---

## Stage 3 — `03_dedupe_and_analyze_cases.py`


---

## Stage 4 — `scrape_opinions.py`

Reads a case workbook, filters to rows with a non-blank
`judicial_acts_link`, and fetches the full text of each judicial decision.

```bash
python scrape_opinions.py                  # TEST MODE: first 10 links only
python scrape_opinions.py --full           # every link
python scrape_opinions.py --assemble-only  # rebuild the Excel from what's already fetched, no network calls (works offline)
python scrape_opinions.py --retry-failures # re-attempt links that errored on a previous run
```

By default reads `ENGLISH_deduped.xlsx` and writes
`court_cases_and_opinions.xlsx` — change `INPUT_XLSX` / `OUTPUT_XLSX` at
the top of the file if your filenames differ.

### Output columns added

- `opinion_text` — the decision text inline (truncated with a
  `[TRUNCATED -- full text in ...]` marker if it exceeds Excel's
  32,767-character cell limit)
- `opinion_text_file` — path to a `.txt` file with the **full,
  untruncated** text, written for *every* fetched decision (not just long
  ones), named after the case number in `opinion_fulltexts/`

### Interrupt-safe by design

Every fetched decision is checkpointed to `opinion_checkpoints/`
(one small JSON file per decision, keyed by a hash of its URL) the
moment it's fetched — not batched. This means:

- You can disconnect wi-fi, close your laptop, or Ctrl+C at any moment
  and lose at most the single request in flight.
- Just re-run the same command to resume — already-checkpointed links
  are skipped automatically.
- Duplicate links (the same decision referenced by multiple case rows)
  are only fetched once.
- Permanently-failing links are also checkpointed (with an error marker)
  so they don't stall every future run; use `--retry-failures` to retry
  them specifically.
- The final Excel is rebuilt from checkpoints at the end of *every* run —
  even a test run produces a valid, complete-so-far
  `court_cases_and_opinions.xlsx`.

---

## Stage 5 — `translate_opinions.py`

Translates the scraped decision texts to English.

```bash
python translate_opinions.py                  # TEST MODE: first 10 opinions
python translate_opinions.py --full           # every opinion
python translate_opinions.py --assemble-only  # rebuild output from cache only, no API calls
python translate_opinions.py my_file.xlsx --full
```

By default reads `court_cases_and_opinions.xlsx` and writes
`court_cases_and_opinions_translated.xlsx`.

### Output columns added

- `opinion_text_ENGLISH` — translation inline (same truncation
  convention as above)
- `opinion_text_file_ENGLISH` — path to a `.txt` with the full
  translation, saved as `{original_name}_EN.txt`

Translates from the full `.txt` file (`opinion_text_file`), not the
(possibly truncated) Excel cell, so long decisions are translated in
full. Long texts are split into ≤4,500-character chunks **at paragraph
boundaries** (never mid-sentence) before translating and rejoined
afterward, since the free Google Translate endpoint caps input length
per call.

### Also interrupt-safe / cached

Same checkpoint pattern as Stage 3
(`translation_checkpoints/`, one file per unique opinion): safe to
interrupt, cheap to resume, and opinions shared across multiple case rows
are translated only once.

---

## Re-running / resuming safely

**Every stage writes to a different file than it reads from.** This is
deliberate: Stages 1 and 3 (the scrapers) *regenerate their output file
from scratch on every run* — that's what makes them resumable — so if a
translation script wrote back into that same file, the next scraper run
would silently overwrite it and destroy your translations with no
warning.

The safe workflow this enables:

1. Scrape some courts / some decisions (in as many partial, interrupted
   sessions as you need).
2. Translate whenever you like.
3. Go back and scrape more (more courts, more decisions) — this rebuilds
   the *original* scraper output file as normal; your translated file is
   untouched.
4. Re-run the translation script — it re-reads the now-larger scraper
   output and, thanks to the on-disk translation caches, only spends API
   calls on genuinely new rows.

You can loop through steps 3–4 indefinitely as you add more courts over
time.

---

## A note on the source data

These scripts read from `sudrf.ru`, the official public case-search
system used by Russian courts (including courts in Russian-occupied
Ukrainian territory). Case numbers, hearing dates, judge names, and
decision text are public court records, comparable in nature to scraping
PACER in the US.

## Politeness / rate-limiting philosophy (all scripts)

- Randomized delays between requests (never a fixed interval), with
  longer periodic breaks.
- A single persistent session per run with realistic browser headers and
  `Referer` chaining, rather than cold, unrelated requests.
- Sequential requests only — no concurrency/parallel hammering of any
  server.
- Exponential backoff and retry on transient errors, 403s, and 429s,
  rather than immediate re-hammering.
- `robots.txt` is checked at startup (warns rather than silently
  ignoring).

If you're running this across many court sites, please still use your
own judgment about spacing out *which* sites you hit and how often — this
is meant to be conservative by default, not a guarantee against being
rate-limited.
