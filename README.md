# Remote AI Job Search Engine

A Python command-line tool that asks you for a **job title**, a **region** (USA, UK, Australia
or Europe) and confirms **Remote only**, then searches **Indeed, LinkedIn and Google Jobs**,
removes duplicate postings, filters out anything that isn't genuinely remote, and exports the
results to a single **Excel (.xlsx)** file.

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
- Job discovery on Indeed, LinkedIn and Google Jobs, using free / open-source tools only.
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
[3] Search the SELECTED region first
      each country x each platform: indeed -> linkedin -> google
      is_remote=True always sent; paginated via offset up to
      max_pages_per_platform, stopping early when a platform repeats rows
      NOTE: hours_old is NOT sent to Indeed - doing so silently cancels
      Indeed's remote filter (see §14), so the date cut-off is applied
      locally to Indeed rows instead
      EACH call wrapped in try/except: a failure is logged, the run continues
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
[7] Cap to target_total_jobs, reorder to the fixed 17-column schema,
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
| Job discovery | **JobSpy** (`python-jobspy`) | MIT-licensed, actively maintained; one function call covers Indeed, LinkedIn and Google Jobs; returns a pandas DataFrame; has a built-in `is_remote` filter |
| Data handling | **pandas** | JobSpy already returns a DataFrame; dedup and column ordering are one-liners |
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
platforms: ["indeed", "linkedin", "google"]
hours_old: 168
results_wanted_per_platform: 50
max_pages_per_platform: 3
target_total_jobs: 50
output_dir: "./output"
remote_strictness: "balanced"
```

| Key | Meaning |
|---|---|
| `platforms` | Which job boards to search. Default: Indeed, LinkedIn, Google Jobs. |
| `hours_old` | Only include postings newer than this many hours (168 = 7 days). |
| `results_wanted_per_platform` | How many results to request per platform before filtering. Fetch generously — duplicates and non-remote rows get removed afterwards. |
| `max_pages_per_platform` | How many pages to request per platform via JobSpy's `offset`. Collects more when a platform has more; stops early once a platform starts repeating rows. |
| `remote_strictness` | `balanced` (default) trusts each platform's server-side remote filter for rows whose wording says nothing either way. `strict` keeps only rows with positive remote evidence - fewer rows, higher confidence. |
| `target_total_jobs` | Maximum rows in the final Excel file. |
| `output_dir` | Where the `.xlsx` file is written. |

If `config.yaml` is missing or empty the script warns and uses these same values as built-in
defaults, so it always stays runnable.

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

Example output from a real run (Australia selected, target 50):

```
[SELECTED] Region: Australia  (need 50 more)
    -- Australia --
      indeed    fetched   2 | stale  0 | carried forward   2
      linkedin  fetched  50 | stale  0 | carried forward  50
      google    fetched   0 | stale  0 | carried forward   0
    Australia: running total 39/50

[FALLBACK] Region: USA  (need 11 more)
    -- USA --
      indeed    fetched  49 | stale  0 | carried forward  49
      linkedin  fetched  50 | stale  0 | carried forward  50
      google    fetched   0 | stale  0 | carried forward   0
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
        Australia      raw  52   (indeed:2 | linkedin:50 | google:0)

  [FALLBACK] USA
      raw rows fetched    : 99
      valid remote rows   : 87
      running total after : 103/50  (capped to 50 at export)
        USA            raw  99   (indeed:49 | linkedin:50 | google:0)

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

Every job row behind the leads, kept for audit and debugging — **exactly these 17 columns in
exactly this order**, unchanged from the brief:

| # | Column | Source | Notes |
|---|---|---|---|
| 1 | `search_keyword` | the job title you typed | Needed for QA/debugging |
| 2 | `source_platform` | JobSpy `site` | `indeed`, `linkedin` or `google` |
| 3 | `country` | resolved from your location | The country the row was fetched under |
| 4 | `job_title` | JobSpy `title` | — |
| 5 | `company_name` | JobSpy `company` | — |
| 6 | `company_url` | JobSpy `company_url` | Blank if not returned — not fetched separately in Stage 1 |
| 7 | `company_industry` | JobSpy `company_industry` | Blank if not returned |
| 8 | `location_raw` | JobSpy `location` | City/state/country exactly as returned |
| 9 | `is_remote` | the remote filter's decision | `True` for every row |
| 10 | `job_type` | JobSpy `job_type` | fulltime / parttime / contract / internship, if available |
| 11 | `date_posted` | JobSpy `date_posted` | — |
| 12 | `salary_min` | JobSpy `min_amount` | Blank if unavailable — **never estimated** |
| 13 | `salary_max` | JobSpy `max_amount` | Blank if unavailable — **never estimated** |
| 14 | `salary_currency` | JobSpy `currency` | Blank if unavailable |
| 15 | `job_url` | JobSpy `job_url` | The apply link |
| 16 | `job_description` | JobSpy `description` | Full text, truncated — see below |
| 17 | `date_fetched` | set by the script | Today's date at run time |

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
9. **Cap and reshape** — trim to `target_total_jobs`, map to the fixed 17-column schema, add `date_fetched`, truncate long descriptions.
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
| 6 | **Column check** | Exactly the 17 columns of §9, in that order |
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
| 22 | **Nothing anywhere** | If no region returns jobs, all four are tried, the run does not crash, and the file still has the 17 columns |

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
| Google Jobs returns 0 | Google Jobs scraping is broken in JobSpy 1.1.82 (the latest release) - it returned 0 for every query tested, including a plain US control search | Not fixable from this project. Indeed and LinkedIn carry the run; the summary reports the 0 rather than hiding it |
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
   often come back flagged `False` — measured at 17–18 of 20. Google derives it from
   description text alone. Dropping every `False` row therefore destroys valid results, which
   is exactly what produced a 3-row spreadsheet. The tiered filter in §3 replaces that rule.

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

8. **"Europe" is not a searchable value** on Indeed, LinkedIn or Google Jobs — the Stage 1
   brief flags this too. It expands to Germany, Netherlands, Ireland, France and Spain,
   searched in turn with an early exit once the target is met. Edit `REGION_DEFINITIONS` in
   `main.py` to change that list; the countries should be confirmed with the business.

9. **The region is recorded in the existing `country` column,** not a new one. The 17-column
   schema is fixed by the brief, so region tracking uses an internal `_region` column that is
   dropped before export.

10. **Fallback is capped, not exhaustive.** Regions are searched in order only until the target
    is met, so a run that fills up on the selected region costs one region's worth of requests
    rather than four.

---

## 15. Resources

- JobSpy repository and README (parameters, supported countries, output schema): <https://github.com/speedyapply/JobSpy>
- `pandas.DataFrame.to_excel` documentation: <https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.to_excel.html>
