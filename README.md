# Remote AI Job Search Engine

A Python command-line tool that asks you for a **job title**, a **region** (USA, UK, Australia
or Europe) and confirms **Remote only**, then searches **Indeed and LinkedIn plus four
dedicated remote-only job boards — RemoteOK, Remotive, We Work Remotely and Jobspresso** (see
§7b), also expanding your title into closely-related keywords (see §7a), removes duplicate
postings, filters out anything that isn't genuinely remote, and exports the results to a
single **Excel (.xlsx)** file.

If your chosen region doesn't have enough remote jobs, the remaining regions are searched
automatically as a fallback until the target count is met — with your selected region's jobs
kept first in the file.

This is **Stage 1**. It is deliberately simple: one script, one config file, one Excel file out.
There is no website, no database, no login and no chatbot.

---

## Table of contents

1. [Objective](#1-objective)
2. [Scope of this stage](#2-scope-of-this-stage)
3. [How it works (workflow)](#3-how-it-works-workflow)
4. [Technology stack](#4-technology-stack)
5. [Folder structure](#5-folder-structure)
6. [Installation](#6-installation)
7. [Configuration (`config.yaml`)](#7-configuration-configyaml)
7a. [Similar-title keyword expansion](#7a-similar-title-keyword-expansion-keyword_expansion)
7b. [Job platforms](#7b-job-platforms-job_sourcespy)
8. [How to run](#8-how-to-run)
9. [Excel output columns](#9-excel-output-columns)
10. [Implementation steps](#10-implementation-steps)
11. [Testing plan](#11-testing-plan)
12. [Acceptance criteria (definition of done)](#12-acceptance-criteria-definition-of-done)
13. [Troubleshooting](#13-troubleshooting)
14. [Notes and assumptions](#14-notes-and-assumptions)
15. [Resources](#15-resources)

---

## 1. Objective

> Build a Python script that produces **one Excel file** listing **remote-only** AI/tech job
> postings, with a fixed set of useful fields per row.

The search itself is driven by three questions asked at the prompt — job title, region, and
remote-only confirmation. Settings that stay the same between searches (which platforms, how
many results, how far back to look) live in `config.yaml`, so you never have to edit a file to
run a new search.

---

## 2. Scope of this stage

**In scope**

- A Python script (`main.py`) plus a settings file (`config.yaml`), run manually from the command line.
- Three interactive inputs: job title, region (USA / UK / Australia / Europe), remote-only confirmation.
- Automatic fallback to the other regions when the selected one comes up short, selected region first in the output.
- Job discovery on Indeed and LinkedIn (via JobSpy) plus RemoteOK, Remotive, We Work Remotely
  and Jobspresso (via `job_sources.py`), using only public APIs/RSS feeds — see §7b.
- **Remote-only** filtering — zero on-site or hybrid rows in the final output.
- Duplicate removal across platforms.
- A clean, fixed-column Excel export plus a console run summary.

**Out of scope (not built in Stage 1)**

- Web frontend or UI
- Database
- User accounts / login
- Chatbot or LLM features
- Scheduling / automation (the script is run by hand)
- Evaluating alternative scraping libraries — JobSpy is the chosen tool for this stage

---

## 3. How it works (workflow)

```
                      config.yaml
                (platforms, limits, output folder)
                            |
                            v
[1] Ask the user three questions
      1. Job title      -> e.g. "AI Engineer"     (re-asks if left blank)
      2. Location       -> pick one region:
                             1. USA
                             2. UK
                             3. Australia
                             4. Europe  (Germany, Netherlands, Ireland,
                                         France, Spain)
      3. Remote only?   -> always remote in Stage 1
                            |
                            v
[2] Expand the chosen region into the country strings JobSpy accepts
      USA / UK / Australia -> one country each
      Europe               -> five countries, searched in turn
      ("Europe" is not a searchable value on any of these job boards)
                            |
                            v
[2b] Expand the job title into similar keywords          (keywords.py)
      "AI Engineer" -> also "Machine Learning Engineer", "ML Engineer",
                       "Artificial Intelligence Engineer", "AI/ML Engineer", ...
      the title you typed is ALWAYS keyword #1 and is never rewritten;
      every added keyword is validated for relatedness before it is searched
      (see §7a). Set keyword_expansion.enabled: false to switch this off.
                            |
                            v
[3] Search the SELECTED region first
      each country x each KEYWORD x each platform:
        indeed -> linkedin -> remoteok -> remotive -> weworkremotely -> jobspresso
      keywords are searched primary-first, and the keyword loop stops
      early once the target is met - so extra keywords cost nothing
      when the primary title already fills the file
      Indeed/LinkedIn: is_remote=True always sent; paginated via offset up
      to max_pages_per_platform, stopping early when a platform repeats rows
      RemoteOK/Remotive/WWR/Jobspresso: dedicated remote-only boards, each
      reached through its own public API or RSS feed (see §7b) - fetched
      once per run and filtered by keyword + required-location, not paginated
      NOTE: hours_old is NOT sent to Indeed - doing so silently cancels
      Indeed's remote filter (see §14), so the date cut-off is applied
      locally to Indeed rows instead (and to every §7b platform, which has
      no hours_old parameter of its own to send)
      EACH platform call wrapped in try/except: a failure is logged, the run continues
                            |
                            v
[4] Filter + dedupe what the selected region returned, then check the count
                            |
              +-------------+-------------+
              |                           |
      target reached?                target NOT reached?
              |                           |
              v                           v
   STOP - fallback regions      [5] FALL BACK through the remaining
   are never searched               regions in fixed order:
                                    USA -> UK -> Australia -> Europe
                                    (skipping the one already searched)
                                    Adding only valid remote jobs, and
                                    stopping the moment the target is met
                            |
                            v
[6] Combine, keeping the SELECTED region's jobs first
      frames are concatenated in search order and dedup keeps the first
      occurrence, so a posting found in two regions is credited to the
      selected one and the selected region stays at the top of the sheet
                            |
                            v
[7] Cap to target_total_jobs, reorder to the fixed 18-column schema,
    add date_fetched, truncate over-long descriptions
                            |
                            v
[8] Export -> output/remote_ai_jobs_YYYY-MM-DD.xlsx  (one file)
                            |
                            v
[9] Print the run summary: per-region and per-country counts, which
    regions were skipped and why, filter and dedup breakdowns, and how
    many rows of the final file came from each region
```

---

## 4. Technology stack

| Purpose | Choice | Why |
|---|---|---|
| Language | Python 3.10+ (tested on 3.11) | Required baseline |
| Job discovery (Indeed/LinkedIn) | **JobSpy** (`python-jobspy`) | MIT-licensed, actively maintained; returns a pandas DataFrame; has a built-in `is_remote` filter |
| Job discovery (remote boards) | **`job_sources.py`** (this project) + **`requests`** | RemoteOK/Remotive public JSON APIs and We Work Remotely/Jobspresso public RSS feeds — see §7b for why these four and not Google Jobs |
| Data handling | **pandas** | Both discovery paths return/are normalised to a DataFrame; dedup and column ordering are one-liners |
| Excel writing | **openpyxl** | The engine pandas uses for `.to_excel()` |
| Settings file | **PyYAML** | YAML is human-editable, so a non-engineer can change settings without touching code |

Everything above is free and open-source. No paid API keys are needed.

---

## 5. Folder structure

```
remote-job-search-engine/
├── README.md            <- this file
├── requirements.txt     <- the four dependencies
├── config.yaml          <- settings that stay the same between searches
├── main.py              <- the script
├── .gitignore           <- keeps .venv and generated files out of Git
└── output/
    ├── .gitkeep                          <- keeps the empty folder in Git
    ├── remote_ai_jobs_YYYY-MM-DD.xlsx    <- generated results
    └── dropped_for_review_YYYY-MM-DD.csv <- rows the remote filter removed
```

---

## 6. Installation

You need **Python 3.10 or newer**. Check with `python --version`.

**Windows (PowerShell)**

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

**macOS / Linux**

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

`requirements.txt` contains:

```
python-jobspy
pandas
openpyxl
pyyaml
```

> **What is a virtual environment?** A private folder of Python packages just for this project,
> so installing things here cannot break other Python projects on your machine. `activate`
> switches it on; type `deactivate` to switch it off.

---

## 7. Configuration (`config.yaml`)

You do **not** need to edit this file to run a search — the job title and location are asked at
the prompt. This file only holds the settings that stay the same between searches.

```yaml
platforms: ["indeed", "linkedin", "remoteok", "remotive", "weworkremotely", "jobspresso"]
hours_old: 168
results_wanted_per_platform: 50
max_pages_per_platform: 3
target_total_jobs: 50
output_dir: "./output"
remote_strictness: "balanced"
```

| Key | Meaning |
|---|---|
| `platforms` | Which job boards to search. Default: Indeed, LinkedIn (via JobSpy) plus RemoteOK, Remotive, We Work Remotely and Jobspresso (via `job_sources.py` — see §7b). Google Jobs is not included — see §7b. |
| `hours_old` | Only include postings newer than this many hours (168 = 7 days). |
| `results_wanted_per_platform` | How many results to request per platform before filtering. Fetch generously — duplicates and non-remote rows get removed afterwards. |
| `max_pages_per_platform` | How many pages to request per platform via JobSpy's `offset`. Collects more when a platform has more; stops early once a platform starts repeating rows. |
| `remote_strictness` | `balanced` (default) trusts each platform's server-side remote filter for rows whose wording says nothing either way. `strict` keeps only rows with positive remote evidence - fewer rows, higher confidence. |
| `target_total_jobs` | Maximum rows in the final Excel file. |
| `output_dir` | Where the `.xlsx` file is written. |

If `config.yaml` is missing or empty the script warns and uses these same values as built-in
defaults, so it always stays runnable.

## 7a. Similar-title keyword expansion (`keyword_expansion`)

A search for `AI Engineer` misses a company that advertises the same job as
`Machine Learning Engineer`. Expansion fixes that: the title you type is searched **plus** the
titles recruiters use for the same role, across every platform in `platforms`.

```
AI Engineer  ->  AI Engineer                        (primary - exactly what you typed)
                 Artificial Intelligence Engineer
                 Machine Learning Engineer
                 Applied AI Engineer
                 ML Engineer
                 AI/ML Engineer
```

Three guarantees:

1. **The original wins.** Your title is always keyword #1, searched first. Deduplication keeps
   the first occurrence, so a job found by both your title and a similar one is credited to
   yours. Nothing you typed is ever rewritten or dropped.
2. **Nothing unrelated gets in.** Every candidate is validated before any platform is called:
   it must name the same *kind* of role (an engineering search never pulls in `Data Analyst`
   or `Product Manager`), must not be a broad catch-all (`Engineer`, `Data`, `Machine
   Learning` on their own are blocked), must be 2–6 words, and must either belong to the same
   curated title family or share a subject word with your title. Rejects are counted and
   printed in the run summary with the reason.
3. **No invented jobs.** Expansion produces *search terms* only. The job boards remain the
   sole source of postings.

| Key | Meaning |
|---|---|
| `enabled` | `false` searches exactly the title you typed, i.e. the pre-expansion behaviour. |
| `max_keywords` | Total keywords per run **including** your own title. Each extra keyword is another pass over every platform, so this is the main cost/coverage dial. It is a ceiling, not a fixed cost — the keyword loop stops as soon as `target_total_jobs` is met. |
| `role_noun_synonyms` | Generate `Engineer` ↔ `Developer` variants. |
| `abbreviation_variants` | Generate the opposite spelling of an abbreviation (`AI Engineer` ↔ `Artificial Intelligence Engineer`). Boards index the two differently, so both are worth searching. |
| `preserve_seniority` | Carry `Senior`/`Lead` from your title onto generated keywords. Off by default — boards match seniority loosely and it narrows results sharply. |
| `extra_synonyms` | `{"your title": ["extra keyword", ...]}` — force in a term the built-in tables lack. Trusted: shape-checked only. |
| `blocked_terms` | Extra terms to reject, added to the built-in broad-term list. |
| `ai.enabled` | Optionally ask a model for more titles. **Off by default**; the built-in tables need no API key and no network call. |

### Adding your own titles

The curated families live in `ROLE_FAMILIES` in [`keywords.py`](keywords.py) — one list per
role, each entry a title a recruiter would use for the *same* job. Add a title to a family, or
a whole new family, and every member becomes reachable from every other. For a one-off, use
`extra_synonyms` in `config.yaml` instead and leave the code alone.

### Using AI for keywords

Set `keyword_expansion.ai.enabled: true`, `pip install anthropic`, and export
`ANTHROPIC_API_KEY`. The model is asked for job **titles** only — never for job listings — and
everything it returns goes through the same validation as every other keyword, so a
hallucinated or off-topic title is rejected before any platform is searched. If the call fails
or the key is missing, the run logs a warning and continues on the built-in tables.

### Reusing the expander

`keywords.py` has no dependency on the rest of the project:

```python
from keywords import KeywordExpander, expand_job_title

expand_job_title("Data Scientist")["keywords"]
KeywordExpander(my_settings).expand("Backend Engineer")   # per-search settings
```

`expand()` returns the keywords, the rejected candidates with reasons, the matched family, and
where each keyword came from.

---

## 7b. Job platforms (`job_sources.py`)

Indeed and LinkedIn are searched through **JobSpy**, unchanged. **Google Jobs has been
removed** — live searches kept returning results with no clear connection to the query,
Google's own remote flag is a text guess it derives itself rather than a real filter (see
§14), and JobSpy's Google scraper has no server-side keyword search of its own. In its place,
four dedicated **remote-only** job boards are searched, each through a source it actually
publishes for reuse — never a scrape of a page meant for browsers, and never a site whose
terms or `robots.txt` say no:

| Platform | Access method | Type |
|---|---|---|
| **RemoteOK** | [`remoteok.com/api`](https://remoteok.com/api) | Public JSON API, no key |
| **Remotive** | [`remotive.com/api/remote-jobs`](https://remotive.com/api/remote-jobs) | Public JSON API, no key |
| **We Work Remotely** | per-category RSS feeds | Public RSS |
| **Jobspresso** | [`jobspresso.co/jobs/feed/`](https://jobspresso.co/jobs/feed/) | Public RSS |

RemoteOK's and Remotive's own terms ask for a link back to the original listing and credit to
the source in return for API access — both are satisfied automatically, because every exported
row's `job_url` points at that platform's own listing page and `source_platform` names it (see
§9). We Work Remotely's `robots.txt` allows crawling everywhere except account/admin pages, and
its RSS feeds are the same syndication format the site links from its own categories page.
Jobspresso's `robots.txt` disallows only query-string URLs (`Disallow: /*?`), so its feed is
fetched with no query string and filtered locally instead of ever being asked to search.

**Not integrated, on purpose** — both were checked and neither has a legitimate free/public
path in:

- **Wellfound (formerly AngelList Talent)** — job search only works through an authenticated
  session calling the site's internal GraphQL API; there is no public REST/JSON API, and
  `robots.txt` explicitly disallows the query-string URLs (`?jobId=`, `?jobSlug=`, …) that
  identify a listing. Scraping the rendered page would mean automating a login and ignoring
  `robots.txt` — both against this project's rule of never bypassing access controls.
- **FlexJobs** — a paid subscription board. Listings are behind a paywall and FlexJobs' own
  terms prohibit automated collection or redistribution of its (paid) content. It offers no
  public API.

### How the dedicated boards fit the existing search

None of these four boards support the country-partitioned search JobSpy gets from
`country_indeed` — each returns one global feed. Two things follow, both handled in
`job_sources.py`:

- **Fetched once, reused across countries.** A Europe search tries five countries in turn;
  re-fetching RemoteOK's whole feed five times for one run would be wasteful and rude to a free
  API. Each board's raw feed is cached for the life of one pipeline run (`job_sources.
  reset_cache()`, called once in `collect_jobs()`) and re-filtered locally per country.
- **Required-location filtering.** A job with no stated restriction ("Worldwide", or nothing at
  all) is offered to every country. A job that says "USA Only" is offered only to a USA search.
  `job_sources.location_permits()` implements this against the short, structured location field
  each board actually provides (Remotive's `candidate_required_location`, We Work Remotely's
  `<region>`, RemoteOK's `location`) — never against the free-text description, where a passing
  mention of a country is not the same claim as a stated restriction. This is what satisfies
  the "exclude jobs restricted to a location that is not selected" requirement for these four
  platforms; Indeed/LinkedIn already get the equivalent from JobSpy's own `country_indeed`
  parameter.

Beyond that, nothing else changes: the same remote-safety filter, deduplication, ICP
classification, company scoring and Excel export in §3 run over rows from these platforms
exactly as they do over Indeed/LinkedIn rows, because every connector returns the identical
column shape JobSpy does (see `job_sources._ROW_COLUMNS`).

### Reliability

Every platform call is wrapped in `try/except` in both `job_sources.py` (timeout + up to
`max_retries` retries with backoff) and again in `main.fetch_platform` — a platform that times
out, rate-limits, or returns malformed data logs a warning and is skipped; it never stops the
other platforms or the run. Configure timeouts/retries per platform under `job_sources:` in
`config.yaml`.

### Adding a platform

1. Write `search_<name>(keyword, country_indeed, hours_old, results_wanted, settings) ->
   pd.DataFrame` in `job_sources.py`, building rows with `_make_row(...)` so the column shape
   matches every other source.
2. Add it to `PLATFORM_REGISTRY` at the bottom of that file.
3. Add its name to `platforms` in `config.yaml` (and, optionally, a settings block under
   `job_sources:`).

Nothing in `main.py` or `pipeline.py` needs to change — `fetch_platform` already dispatches any
platform not in `JOBSPY_PLATFORMS` through `job_sources.PLATFORM_REGISTRY`.

### Why not ZipRecruiter / Glassdoor?

ZipRecruiter only covers the US and Canada, and Glassdoor's country coverage is patchier than
Indeed's. Both are a poor fit for a UK/EU/AU/NZ-heavy search. Add them to `platforms` later
only if a region keeps coming back thin.

---

## 8. How to run

```bash
python main.py
```

Optionally pass a different settings file: `python main.py my-config.yaml`

The script then asks:

```
1. Job title (e.g., 'AI Engineer'): AI Engineer

2. Select a location:
     1. USA
     2. UK
     3. Australia
     4. Europe  (Germany, Netherlands, Ireland, France, Spain)
   Enter 1-4: 3
3. Work mode - remote only? [Y/n]: y
```

If the region you pick returns enough jobs, the run stops there. If it comes up short, the
remaining regions are searched automatically in the order USA → UK → Australia → Europe, and
only valid remote jobs are added until the target is met. Your selected region's jobs always
stay at the top of the spreadsheet.

Example output, illustrative of the shape (platform counts vary run to run - live counts for an
"AI Engineer" search are in §11 'Verified results — new platform lineup'):

```
[SELECTED] Region: Australia  (need 50 more)
    -- Australia --
      indeed         fetched   2 | stale  0 | carried forward   2
      linkedin       fetched  50 | stale  0 | carried forward  50
      remoteok       fetched   3 | stale  0 | carried forward   3
      remotive       fetched   2 | stale  0 | carried forward   2
      weworkremotely fetched   5 | stale  1 | carried forward   4
      jobspresso     fetched   1 | stale  0 | carried forward   1
    Australia: running total 39/50

[FALLBACK] Region: USA  (need 11 more)
    -- USA --
      indeed         fetched  49 | stale  0 | carried forward  49
      linkedin       fetched  50 | stale  0 | carried forward  50
      remoteok       fetched   4 | stale  0 | carried forward   4
      remotive       fetched   3 | stale  0 | carried forward   3
      weworkremotely fetched   6 | stale  0 | carried forward   6
      jobspresso     fetched   1 | stale  0 | carried forward   1
    USA: running total 103/50
    Target of 50 reached - no further regions will be searched.

======================================================================
RUN SUMMARY
======================================================================

Selected region : Australia
Target count    : 50
Final exported  : 50

Region-by-region (searched in this order, selected region first):

  [SELECTED] Australia
      raw rows fetched    : 52
      valid remote rows   : 42
      running total after : 39/50
        Australia      raw  52   (indeed:2 | linkedin:50 | remoteok:3 | remotive:2 | weworkremotely:4 | jobspresso:1)

  [FALLBACK] USA
      raw rows fetched    : 99
      valid remote rows   : 87
      running total after : 103/50  (capped to 50 at export)
        USA            raw  99   (indeed:49 | linkedin:50 | remoteok:4 | remotive:3 | weworkremotely:6 | jobspresso:1)

  Not searched (target already met): UK, Europe

Remote filter (remote_strictness=balanced) - why rows were kept or dropped:
  KEPT    platform explicitly confirmed remote      : 81
  KEPT    text/location contained a remote indicator: 0
  KEPT    no evidence either way, server-side remote
          filter trusted                            : 48
  DROPPED text explicitly said on-site/hybrid       : 22
  DROPPED no positive remote evidence (strict mode) : 0

Deduplication - rows removed by each pass:
  same job_url                        : 0
  same title+company+location (no URL): 0
  same title+company, different city  : 0

Rows per region in the exported file (selected region first):
  Australia     39  <-- selected
  USA           11
======================================================================

Output written to: output\remote_ai_jobs_2026-09-08.xlsx
Dropped rows (for your review) written to: output\dropped_for_review_2026-09-08.csv
```

Every skipped region, every zero, and every filter decision is reported rather than silently
swallowed.

---

## 9. Excel output

The export is one `.xlsx` file with **two sheets**.

### Sheet 1 — "Company Leads" (primary output)

One row per company, not per job. This is the lead list.

| Column | Meaning |
|---|---|
| `company_name` | Most frequently seen spelling of the name |
| `company_domain` | The company's own website domain, when a board returned one — never guessed |
| `company_url`, `company_url_direct` | Board profile page and company site as returned |
| `company_industry`, `company_num_employees` | As returned by the board; blank when absent |
| `country`, `region` | Where the postings were found |
| `matching_position_count` | How many matching openings this company has |
| `matching_job_titles` | All distinct titles, `\|`-separated |
| `matching_job_urls` | All job URLs, `\|`-separated |
| `locations`, `source_platforms`, `search_keywords` | Combined across the company's postings |
| `earliest_job_posted`, `latest_job_posted` | Date range of the openings |
| `date_fetched` | Run date |

### Sheet 2 — "Remote Jobs" (job-level audit trail)

### The job-level column schema

Every job row behind the leads, kept for audit and debugging — **exactly these 18 columns in
exactly this order**:

| # | Column | Source | Notes |
|---|---|---|---|
| 1 | `original_job_title` | the job title you typed | Identical on every row of a run |
| 2 | `matched_keyword` | the keyword that returned this row | Equals column 1 unless a similar keyword found it — see §7a |
| 3 | `source_platform` | `site` (JobSpy, or set by job_sources.py) | `indeed`, `linkedin`, `remoteok`, `remotive`, `weworkremotely` or `jobspresso` — see §7b |
| 4 | `country` | resolved from your location | The country the row was fetched under |
| 5 | `job_title` | `title` | — |
| 6 | `company_name` | `company` | — |
| 7 | `company_url` | `company_url` | Blank if not returned — none of the §7b boards expose a verified company-owned domain, so it stays blank there too, same as any other platform that doesn't provide one |
| 8 | `company_industry` | `company_industry` | Blank if not returned |
| 9 | `location_raw` | `location` | City/state/country as returned; for §7b platforms, their stated required-location ("Worldwide", "USA Only", …) |
| 10 | `is_remote` | the remote filter's decision | `True` for every row |
| 11 | `job_type` | `job_type` | fulltime / parttime / contract / internship, if available |
| 12 | `date_posted` | `date_posted` | — |
| 13 | `salary_min` | `min_amount` | Blank if unavailable — **never estimated** |
| 14 | `salary_max` | `max_amount` | Blank if unavailable — **never estimated** |
| 15 | `salary_currency` | `currency` | Blank if unavailable |
| 16 | `job_url` | `job_url` | **The original listing on that platform** — RemoteOK/Remotive/WWR/Jobspresso rows link to their own listing page, never a copy or a search result |
| 17 | `job_description` | `description` | Full text, truncated — see below |
| 18 | `date_fetched` | set by the script | Today's date at run time |

> ⚠️ **This schema is fixed.** Do not add, remove, rename or reorder columns without approval.

> ⚠️ **Excel cell limit:** a cell holds at most ~32,767 characters. `job_description` is
> truncated at 32,000 characters with a `... [truncated]` marker so the export never fails on
> one oversized row.

---

## 10. Implementation steps

1. **Project setup** — create the folder structure from §5 and install the four dependencies.
2. **Load settings** — read `config.yaml`, applying defaults for anything missing.
3. **Ask the three questions** — re-prompt on a blank job title or location; remote-only is fixed for this stage.
4. **Expand the chosen region** into the country strings JobSpy accepts. USA, UK and Australia are one country each; Europe expands to five, searched in turn with an early exit once the target is met.
5. **Fetch per platform** — call JobSpy with `is_remote=True`, `hours_old` and `results_wanted_per_platform`. Wrap every call in `try/except`, log the failure, continue. Request LinkedIn descriptions explicitly, or LinkedIn rows arrive with an empty description and the text safety check has nothing to read.
6. **Combine** — tag each row with `search_keyword` and `country`, concatenate into one DataFrame.
7. **Remote safety filter** — apply the three-way rule from §3, keeping the dropped rows aside for review.
8. **Deduplicate** — prefer `job_url`; fall back to normalised `title + company + location` for rows without a URL.
9. **Cap and reshape** — trim to `target_total_jobs`, map to the fixed 18-column schema, add `date_fetched`, truncate long descriptions.
10. **Export** — write the `.xlsx`, plus the dropped-rows CSV when anything was dropped.
11. **Print the run summary** — counts at each stage, per-platform breakdown, and a note when the final count is below target.

---

## 11. Testing plan

| # | Test | Expected result |
|---|---|---|
| 1 | **Smoke test**: one job title, one location, `target_total_jobs: 10` | Runs end to end, produces one `.xlsx` with ≤10 rows |
| 2 | **Full run** with the default config | Completes without crashing; run summary printed |
| 3 | **Per-platform sanity check** | No platform is unexpectedly empty; a `0` is flagged in the summary |
| 4 | **Remote-only spot check** | Open 10 random `job_url` values; each posting is genuinely remote |
| 5 | **Duplicate check** | No two rows share a `job_url` |
| 6 | **Column check** | Exactly the 18 columns of §9, in that order |
| 7 | **Long-description check** | Export succeeds; no cell exceeds Excel's limit |
| 8 | **Region menu** | Only 1-4 (or a region name) is accepted; anything else re-prompts |
| 9 | **Failure resilience** | A platform that errors is logged; the run still produces a file |
| 10 | **Blank input** | Empty job title or location re-prompts instead of crashing |
| 11 | **Locked output file** | If the `.xlsx` is open in Excel, the run saves under a `_2` suffix instead of crashing |
| 12 | **Indeed parameter check** | `hours_old` is never sent to Indeed alongside `is_remote`; the date cut-off is applied locally instead |
| 13 | **Recency filter** | Stale rows dropped; rows with an unknown `date_posted` are kept, not discarded |
| 14 | **Strictness modes** | `balanced` keeps no-evidence rows, `strict` drops them; both drop rows whose text says on-site/hybrid |
| 15 | **Cross-city dedup** | The same title+company listed in two cities collapses to one row, while genuinely different roles at the same company survive |
| 16 | **Pagination** | Repeated rows across pages are not double-counted, and pagination stops when a platform runs out |
| 17 | **No fallback when not needed** | A region that meets the target is the only one searched; the rest are reported as skipped |
| 18 | **Fallback order** | A short region falls back through USA → UK → Australia → Europe, skipping the one already searched |
| 19 | **Selected region first** | Every selected-region row appears above every fallback row in the exported file |
| 20 | **Europe early exit** | Europe stops after the first country that meets the target instead of searching all five |
| 21 | **Cross-region dedup** | A posting found in two regions appears once, credited to the selected region |
| 22 | **Nothing anywhere** | If no region returns jobs, all four are tried, the run does not crash, and the file still has the 18 columns |

### Verified results (2026-09-08)

**Scenario A — selected region has enough.** `AI Engineer` / USA / target 50.

| | |
|---|---|
| USA | 98 raw → 81 valid remote → **50 exported** |
| UK, Australia, Europe | **not searched** — target already met |

**Scenario B — selected region comes up short.** `Prompt Engineer` / Australia / target 50.

| | |
|---|---|
| Australia (selected) | 52 raw → 42 valid remote → 39 kept |
| USA (fallback) | 99 raw → 87 valid remote → 11 added |
| UK, Europe | **not searched** — target met after USA |
| Exported | **50** — rows 1–39 Australia, rows 40–50 USA |

Checks on the Scenario B file:

- Columns exactly match §9, in order ✔
- Every row's `is_remote` is `True` ✔
- Zero duplicate `job_url` values ✔
- Every Australia row precedes every USA row ✔
- The internal `_region` column never reaches the export ✔
- Longest description 10,894 characters — under the Excel limit ✔

42 automated logic checks also pass, covering region expansion, the no-fallback and
fallback paths, fallback ordering, selected-region-first ordering, Europe's early exit,
cross-region dedup, the fixed schema, per-platform parameter construction, both strictness
modes, all three dedup passes, pagination, and platform-error resilience.

### Verified results — new platform lineup (2026-09-09)

Live `AI Engineer` / USA search, default keyword expansion capped to 3 keywords for the run
(`AI Engineer`, `Artificial Intelligence Engineer`, `Machine Learning Engineer`), 20 results
requested per platform, target 120:

| Platform | Raw fetched (USA) | Notes |
|---|---|---|
| indeed | 57 | via JobSpy |
| linkedin | 60 | via JobSpy |
| remoteok | 1 match, 0 kept | the one match (`AI Engineer Data APIs` @ Benzinga) was 10 days old — outside the default 168h window |
| remotive | 1 match, 0 kept | the one match (`Senior Independent AI Engineer / Architect`, required location `Americas, Europe, Israel`) was 24 days old |
| weworkremotely | 7 matches, 0 kept | all 7 matches (`AI/ML Engineer for an AI-Driven E-Commerce Platform` @ Toptal, `Senior Software AI Engineer` @ Collaboration.Ai, etc.) were 14–24 days old |
| jobspresso | 0 | its ~20-item recent feed had no title containing both "AI" and "Engineer" at the time of the run |

Re-run with `hours_old: 720` (30 days) instead of the 168-hour default, everything else
unchanged, to confirm the boutique platforms' full path through the pipeline — fetch, remote
filter, dedup, ICP, cap, export — once their postings are inside the window:

```
indeed              72 rows
linkedin            40 rows
remoteok             1 rows
remotive             1 rows
weworkremotely       6 rows
jobspresso           0 rows
                   -------
Total exported     120 rows   (target 120, USA 98 + UK fallback 22)
Duplicate job_urls in export: 0
Rows flagged remote: 120/120
```

Confirms the same three matches that were correctly excluded as stale above (RemoteOK's,
Remotive's, and 6 of We Work Remotely's 7) are correctly *included* once they're inside the
recency window, and reach the export with their `source_platform`/`job_url` intact -
`source_platform` values `remoteok`/`remotive`/`weworkremotely` and real listing URLs
(`remoteok.com/remote-jobs/...`, `remotive.com/remote-jobs/...`,
`weworkremotely.com/remote-jobs/...`) all present in the output workbook. Jobspresso's 0 is
unrelated to recency - see the row above.

This is the accurate, current state of these boards for one specific title, not a defect: they
are small, live, general-audience remote-job feeds (RemoteOK's whole feed is ~100 postings
across every field, not just tech; We Work Remotely's tech categories return ~170 recent
postings across all of software engineering). A title as specific as "AI Engineer" naturally
has few matches in a 7-day window on any one of them — which is exactly why keyword expansion
(§7a) and searching four extra platforms both matter more here than they would for Indeed or
LinkedIn's much larger indexes. Widening `hours_old`, adding more of §7a's expanded keywords,
or raising `job_sources.<platform>.results_wanted` in config.yaml all increase the number of
matches these boards contribute.

---

## 12. Acceptance criteria (definition of done)

- [x] Running `python main.py` and answering the three prompts produces **one** `.xlsx` file in `output/`.
- [x] Every row's `is_remote` is `True`.
- [x] No two rows share the same `job_url`.
- [x] The file has **exactly** the columns listed in §9, **in that order**.
- [x] Final row count is at or below `target_total_jobs`, and the console summary explains the count when it is below target.
- [x] The script does not crash when a platform returns zero results or errors — it logs and continues.
- [x] `README.md` exists and someone who did not write the code could run the project from it alone.

---

## 13. Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| A platform returns 0 results | Country not resolved, or that board has nothing matching | Include the country in your location answer (e.g. "Berlin, Germany") |
| A §7b platform (RemoteOK/Remotive/WWR/Jobspresso) returns 0 | Its feed genuinely has nothing matching that keyword right now (these are small, live, frequently-changing feeds - see §11 'Verified results — new platform lineup'), or the call failed and was logged as a `[WARN]` | Check the console for a `[WARN]` line naming that platform; if there isn't one, the feed is just thin for that search right now |
| Google Jobs used to return 0 | Removed - Google Jobs scraping was broken in JobSpy 1.1.82, returning 0 for every query tested including a plain US control search. Replaced by the §7b platforms, each with a real, working access method | N/A - no longer part of `platforms` |
| Far fewer jobs than expected | Sending `hours_old` and `is_remote` to Indeed together silently cancels Indeed's remote filter | Fixed - Indeed now gets `is_remote`, and the date cut-off is applied locally on `date_posted` |
| `ModuleNotFoundError: jobspy` | Virtual environment not activated, or dependencies not installed | Activate `.venv`, re-run `pip install -r requirements.txt` |
| `PermissionError` on export | The output file is still open in Excel | The script saves under a `_2` suffix automatically; close Excel to get the plain filename |
| Very few results overall | `hours_old` too small | Increase it in `config.yaml` |
| LinkedIn returns little or throws | Rate limiting on the platform side | Logged and skipped; re-run later |
| Row count below target | Genuinely few remote postings after filtering | Check the remote-filter breakdown and `dropped_for_review_*.csv` |

---

## 14. Notes and assumptions

1. **Interactive input vs. config file.** The Stage 1 brief specified a `config.yaml` holding
   `keywords` and `regions`. The agreed build asks for the job title and location at the prompt
   instead, so a non-engineer can run a new search without editing YAML. `config.yaml` still
   holds everything that doesn't change between searches. The Excel schema in §9 is unchanged
   from the brief.
2. **Work mode is always remote.** Question 3 is asked explicitly as required, but answering
   "n" still proceeds as remote-only — on-site and hybrid searches are out of scope for Stage 1.
3. **`target_total_jobs`** is a cap applied after deduplication, not a quota to fill. If fewer
   genuine remote postings exist, the output is smaller rather than padded.
4. **Indeed silently drops the remote filter if you also send `hours_old`.** JobSpy's Indeed
   scraper builds its filters with an `if/elif` chain that tests `hours_old` first, so passing
   both means `is_remote` is never applied. Measured on a live "AI engineer / India" search:
   with both, 19 of 20 rows came back not-remote — identical to sending no filters at all;
   with `is_remote` alone, 20 of 20 were remote. This project therefore sends `is_remote` to
   Indeed and applies the date cut-off locally on `date_posted`.

5. **JobSpy's `is_remote` flag means different things per platform.** Indeed sets it from the
   real server-side filter. LinkedIn ignores it on output and recomputes it with a text
   heuristic, so genuinely-remote rows (already restricted by LinkedIn's own `f_WT=2` filter)
   often come back flagged `False` — measured at 17–18 of 20. Google (now removed - see §7b)
   derived it from description text alone too. Dropping every `False` row therefore destroys
   valid results, which is exactly what produced a 3-row spreadsheet. The tiered filter in §3
   replaces that rule - and it's why the §7b connectors leave `is_remote` unset rather than
   asserting `True` from "this board is remote-only": a board that is remote-only in general
   can still list one hybrid posting by mistake, and the tiered text check is what actually
   catches that (see §11 'Verified results — new platform lineup').

6. **`remote_strictness` is a genuine trade-off, not a tuning knob.** Both platforms' remote
   filters were verified to work: requesting remote returns a materially different result set
   (LinkedIn shares only 10 of 30 URLs with an unfiltered search; Indeed 1 of 30). But rows
   with no remote wording anywhere cannot be confirmed from the scraped data alone, and Indeed
   serves a Cloudflare bot check to automated requests, so per-posting verification is not
   possible from this script. `balanced` keeps those rows and flags how many there are;
   `strict` excludes them. Opening a few `job_url`s by hand remains the final check, as the
   acceptance criteria require.

7a. **Company identity uses exact matching only — never similarity.** Two postings are the
   same company if they share a website domain, a board profile URL, or a normalised name.
   There is no fuzzy matching, because "Delta Systems" and "Delta Systems Group Holdings" are
   plausibly different firms. A guard also blocks a name match when the two rows advertise
   different websites. Legal suffixes (Ltd, Inc, GmbH) are stripped; `Group`, `Holdings`,
   `Company` and `Co` are deliberately **not**, since they are usually part of the name.

7b. **Every dropped row is written to `dropped_for_review_*.csv`** with the reason it was
   dropped, so the filter's decisions can be audited rather than trusted.

8. **"Europe" is not a searchable value** on Indeed or LinkedIn — the Stage 1 brief flags this
   too. It expands to Germany, Netherlands, Ireland, France and Spain,
   searched in turn with an early exit once the target is met. Edit `REGION_DEFINITIONS` in
   `main.py` to change that list; the countries should be confirmed with the business.

9. **The region is recorded in the existing `country` column,** not a new one. The 18-column
   schema is fixed by the brief, so region tracking uses an internal `_region` column that is
   dropped before export.

10. **Fallback is capped, not exhaustive.** Regions are searched in order only until the target
    is met, so a run that fills up on the selected region costs one region's worth of requests
    rather than four.

---

## 15. Resources

- JobSpy repository and README (parameters, supported countries, output schema): <https://github.com/speedyapply/JobSpy>
- `pandas.DataFrame.to_excel` documentation: <https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.to_excel.html>
