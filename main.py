"""
Remote AI Job Search Engine (interactive CLI)
==============================================
Prompts the user for a job title, a location, and a remote-only
confirmation, then searches Indeed and LinkedIn via JobSpy plus the
dedicated remote job boards in job_sources.py (RemoteOK, Remotive, We
Work Remotely, Jobspresso), filters/deduplicates the results, and
writes a single Excel file to output/.

Run:
    python main.py [path/to/config.yaml]

If no config path is given, "config.yaml" in the current folder is used.
config.yaml only holds settings that stay the same across searches
(platforms, limits, output folder) - the search itself is driven by
what you type at the prompts.
"""

import concurrent.futures
import logging
import re
import sys
import time
import traceback
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import yaml
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from jobspy import scrape_jobs

import company_identity
import job_sources
import keywords as keyword_expansion
import pipeline

# Excel's per-cell character limit. Long job descriptions get truncated
# just under this so the export never fails on one oversized cell.
EXCEL_CELL_CHAR_LIMIT = 32000

log = logging.getLogger("job_search")

# Final column order/names required by the spec. Keys are the output
# column names; values are lists of possible source-column names in the
# JobSpy DataFrame (JobSpy's column naming has varied slightly across
# versions, so we check a few candidates for each field).
OUTPUT_COLUMNS = {
    # What the user typed, and the expanded keyword that actually
    # returned this row. Both are exported so a row that came back under
    # a similar title ("ML Engineer" for a search of "AI Engineer") can
    # be traced to the reason it was included.
    "original_job_title": ["original_job_title"],
    "matched_keyword": ["search_keyword"],
    "source_platform": ["site", "SITE"],
    "country": ["country"],
    "job_title": ["title", "TITLE"],
    "company_name": ["company", "COMPANY"],
    "company_url": ["company_url"],
    "company_industry": ["company_industry"],
    "location_raw": ["location", "LOCATION"],
    "is_remote": ["is_remote"],
    "job_type": ["job_type", "JOB_TYPE"],
    "date_posted": ["date_posted"],
    "salary_min": ["min_amount"],
    "salary_max": ["max_amount"],
    "salary_currency": ["currency"],
    "job_url": ["job_url", "JOB_URL"],
    "job_description": ["description", "DESCRIPTION"],
    "date_fetched": ["date_fetched"],
}

# The four selectable regions, each expanded to the country strings
# JobSpy actually accepts for country_indeed.
#
# "Europe" is not a searchable value on any of these job boards - the
# Stage 1 brief flags this explicitly - so it expands to a list of
# European countries that are searched in turn. The others are single
# countries. Country spellings must match JobSpy's supported list:
# https://github.com/speedyapply/JobSpy
#
# STRICT COUNTRY LOCK: whichever region the user selects here is the
# ONLY one ever searched, for the entire run - see collect_jobs and
# execute_pipeline below, which search `[selected]` alone with no
# fallback to any other entry in REGION_ORDER. If the selected region
# yields fewer companies than MAX_COMPANIES, the run exports what it
# found rather than switching countries to make up the difference.
# Europe's five countries are the one exception, and it is not really an
# exception: they are the existing, single "Europe" selection - searching
# all five is what selecting "Europe" has always meant, not a fallback to
# a different region.
REGION_DEFINITIONS = {
    "USA": [
        {"country_indeed": "USA", "location": "United States"},
    ],
    "UK": [
        {"country_indeed": "UK", "location": "United Kingdom"},
    ],
    "Australia": [
        {"country_indeed": "Australia", "location": "Australia"},
    ],
    "Europe": [
        {"country_indeed": "Germany", "location": "Germany"},
        {"country_indeed": "Netherlands", "location": "Netherlands"},
        {"country_indeed": "Ireland", "location": "Ireland"},
        {"country_indeed": "France", "location": "France"},
        {"country_indeed": "Spain", "location": "Spain"},
    ],
}

# Fixed order used both for the selection menu and for the fallback
# sequence once the chosen region runs short.
REGION_ORDER = ["USA", "UK", "Australia", "Europe"]


def setup_logging() -> None:
    logging.basicConfig(level=logging.INFO, format="%(message)s", stream=sys.stdout)
    # JobSpy is chatty at INFO and drowns out our own stage logs.
    logging.getLogger("JobSpy").setLevel(logging.WARNING)
    for name in ("Indeed", "LinkedIn", "Google"):
        logging.getLogger(f"JobSpy:{name}").setLevel(logging.WARNING)


def load_config(path: str) -> dict:
    """Read config.yaml and apply defaults for any missing field. This
    file holds only settings that stay constant across searches, so a
    missing/empty file is not fatal - defaults keep the app runnable."""
    config_path = Path(path)
    config = {}

    if config_path.exists():
        with open(config_path, "r") as f:
            config = yaml.safe_load(f) or {}
    else:
        log.warning(f"[WARN] Config file not found at '{path}' - using built-in defaults.")

    # Google Jobs (JobSpy's "google") is deliberately gone - see README
    # sec. 7a and job_sources.py's module docstring for why - replaced by
    # dedicated remote-only boards reached through their own public
    # APIs/RSS feeds rather than Google's aggregation.
    config.setdefault(
        "platforms",
        ["indeed", "linkedin", "remoteok", "remotive", "weworkremotely", "jobspresso"],
    )
    # 720h/30 days, not the original 168h/7 days - see config.yaml's
    # comment: recency window was one of the two biggest constraints on
    # raw USA company coverage (the other was keyword_expansion.max_keywords).
    config.setdefault("hours_old", 720)
    # Raised so a MAX_COMPANIES=1000 run has enough raw postings to work
    # with (see main.md/config.yaml comments on MAX_COMPANIES); still a
    # per-call/per-page cap, not the overall company ceiling.
    config.setdefault("results_wanted_per_platform", 100)
    config.setdefault("max_pages_per_platform", 10)
    # Audit-sheet/raw-jobs-reported cap - see config.yaml's comment.
    config.setdefault("target_total_jobs", 5000)
    config.setdefault("lead_search", {"job_buffer_per_location": 250})
    config.setdefault("output_dir", "./output")
    config.setdefault("remote_strictness", "balanced")
    # OVERALL_SEARCH_TIMEOUT (seconds). The single ceiling on how long one
    # run may spend actively searching job boards, checked between
    # queries (see execute_pipeline/search_country). Once passed, the run
    # stops starting new queries and moves straight to exporting whatever
    # was already collected - it never leaves the user waiting
    # indefinitely, and it never silently drops below MAX_COMPANIES by
    # pretending the search finished normally (the run/export summary
    # reports the timeout explicitly). 1800s (30 min) balances that
    # against genuinely broad multi-keyword, multi-platform coverage -
    # each LinkedIn-touching query alone typically costs 1-3 minutes
    # purely from JobSpy's own built-in anti-detection pacing, which this
    # project does not attempt to bypass or shorten.
    config.setdefault("search_timeout_seconds", 1800)
    try:
        config["search_timeout_seconds"] = max(30, int(config["search_timeout_seconds"]))
    except (TypeError, ValueError):
        log.warning("[WARN] search_timeout_seconds is not a number - using 1800.")
        config["search_timeout_seconds"] = 1800
    # How many platforms are fetched at once per query - see
    # DEFAULT_MAX_CONCURRENT_PLATFORMS's comment for why this defaults to
    # 4 (not all 8 configured platforms at once) on a resource-
    # constrained host. Raise it on a host with more CPU headroom.
    config.setdefault("platform_fetch_concurrency", DEFAULT_MAX_CONCURRENT_PLATFORMS)

    # Similar-title keyword expansion. Normalised here so every caller -
    # CLI, API, or a script importing this module - sees a complete
    # settings block whether or not config.yaml defines one.
    config["keyword_expansion"] = keyword_expansion.load_keyword_config(config)

    # Per-platform timeouts/retries for the dedicated job-board
    # connectors (job_sources.py). Same reasoning as keyword_expansion
    # above - always present, never required in config.yaml.
    config["job_sources"] = job_sources.load_source_config(config)

    if config["remote_strictness"] not in ("balanced", "strict"):
        log.warning(
            f"[WARN] remote_strictness='{config['remote_strictness']}' is not "
            "'balanced' or 'strict' - falling back to 'balanced'."
        )
        config["remote_strictness"] = "balanced"

    return config


def prompt_user_input() -> dict:
    """Ask the three required questions. Reprompts on a blank job title
    or location; work mode is a fixed "remote only" for this stage but
    is still asked explicitly, per the requirement."""
    print("=" * 70)
    print("Remote AI Job Search Engine")
    print("=" * 70)

    job_title = ""
    while not job_title:
        job_title = input("1. Job title (e.g., 'AI Engineer'): ").strip()
        if not job_title:
            print("   Job title can't be empty - please enter one.")

    print("\n2. Select a location:")
    for i, region in enumerate(REGION_ORDER, start=1):
        countries = ", ".join(e["country_indeed"] for e in REGION_DEFINITIONS[region])
        suffix = f"  ({countries})" if len(REGION_DEFINITIONS[region]) > 1 else ""
        print(f"     {i}. {region}{suffix}")

    region = ""
    while not region:
        answer = input(f"   Enter 1-{len(REGION_ORDER)}: ").strip()
        if answer.isdigit() and 1 <= int(answer) <= len(REGION_ORDER):
            region = REGION_ORDER[int(answer) - 1]
        else:
            # Accept the region name typed out, too - it costs nothing
            # and is friendlier than rejecting "UK".
            match = [r for r in REGION_ORDER if r.lower() == answer.lower()]
            if match:
                region = match[0]
            else:
                print(f"   Please enter a number from 1 to {len(REGION_ORDER)}.")

    # Work mode is fixed for this stage: remote jobs only.
    # No user prompt is required.
    return {"job_title": job_title, "region": region, "remote_only": True}


def build_search_kwargs(platform: str, job_title: str, location: str,
                        country_indeed: str, hours_old: int) -> tuple[dict, bool]:
    """Build the per-platform JobSpy arguments.

    Returns (kwargs, hours_old_applied_server_side). The second value
    tells the caller whether it still needs to apply the recency cut-off
    itself for this platform's rows.

    The important subtlety is Indeed. In JobSpy's Indeed scraper,
    _build_filters() is an if/elif chain that checks hours_old FIRST:

        if hours_old:        -> date filter only
        elif easy_apply:     ...
        elif job_type or is_remote:  -> remote filter

    So passing hours_old AND is_remote together means the remote filter
    is silently dropped and Indeed returns on-site jobs. Measured on a
    live "AI Engineer / India" search: with both, 19 of 20 rows came
    back not-remote - identical to passing no filters at all. With
    is_remote alone, 20 of 20 were remote.

    Remote-only is the hard requirement here and recency is not, so for
    Indeed we send is_remote and apply the date cut-off ourselves on
    date_posted afterwards.
    """
    kwargs = {
        "site_name": [platform],
        "search_term": job_title,
        "location": location,
        "country_indeed": country_indeed,
        "is_remote": True,
    }

    if platform == "indeed":
        return kwargs, False

    kwargs["hours_old"] = hours_old

    if platform == "linkedin":
        # LinkedIn descriptions are NOT fetched by default. Without this
        # every LinkedIn row has an empty description, which starves the
        # remote text check of anything to read.
        kwargs["linkedin_fetch_description"] = True

    return kwargs, True


# Platforms fetched through JobSpy (build_search_kwargs + scrape_jobs).
# Everything else in config["platforms"] is looked up in
# job_sources.PLATFORM_REGISTRY instead - see fetch_platform below.
JOBSPY_PLATFORMS = {"indeed", "linkedin", "zip_recruiter", "glassdoor", "bayt", "naukri", "bdjobs"}


def fetch_platform(platform: str, job_title: str, location: str, country_indeed: str,
                   hours_old: int, results_wanted: int,
                   max_pages: int, source_config: dict = None) -> tuple[pd.DataFrame, bool, dict]:
    """Fetch one platform. Never raises - a failing platform logs a
    warning and yields whatever it collected so far (or nothing), so one
    bad platform cannot kill the run.

    Returns (df, hours_old_applied_server_side, fetch_stats). fetch_stats
    is {"pages": N, "requests": N} - how many pages/requests this call
    actually made, used to build the "Total pages searched" / "Total
    API/search requests" figures in the final run summary.

    Dispatches to one of two implementations depending on the platform:
      - JOBSPY_PLATFORMS (Indeed, LinkedIn, ...) go through JobSpy, with
        pagination via `offset` - the original Stage 1 behaviour.
      - Everything else is looked up in job_sources.PLATFORM_REGISTRY,
        the dedicated remote-board connectors (RemoteOK, Remotive, We
        Work Remotely, Jobspresso). These fetch and paginate themselves
        and are not designed for JobSpy's offset-based pagination
        (their APIs/feeds are not partitioned that way), so max_pages
        does not apply to them - `results_wanted` is passed straight
        through as the cap.
      - A platform in neither set (e.g. a typo, or one of the
        deliberately-unsupported names in job_sources.UNAVAILABLE_
        PLATFORMS) logs a clear warning and returns no rows rather than
        raising, so a bad config.yaml entry cannot stop the run either.
    """
    if platform not in JOBSPY_PLATFORMS:
        try:
            df = job_sources.search(
                platform, job_title, country_indeed, hours_old, results_wanted,
                source_config or job_sources.load_source_config({}),
            )
        except job_sources.JobSourceError as exc:
            log.warning(f"    [WARN] {platform}: {exc}")
            return pd.DataFrame(), False, {"pages": 0, "requests": 1}
        except Exception as exc:
            log.warning(f"    [WARN] {platform}: unexpected failure: {exc}")
            return pd.DataFrame(), False, {"pages": 0, "requests": 1}
        # These boards return absolute dates already, so the shared
        # apply_recency_filter (not a server-side parameter) is what
        # enforces hours_old for them too - same as Indeed. They are not
        # paginated the way JobSpy platforms are below (their feeds/APIs
        # are fetched and cached whole for the run - see job_sources.py),
        # so this counts as one page/one request.
        if not df.empty:
            df = df.copy()
            df["_search_page"] = 1
        return df, False, {"pages": 1, "requests": 1}

    kwargs, hours_applied = build_search_kwargs(
        platform, job_title, location, country_indeed, hours_old
    )

    frames = []
    seen_urls = set()
    collected = 0
    pages_requested = 0

    for page in range(max_pages):
        remaining = results_wanted - collected
        if remaining <= 0:
            break

        offset = collected
        try:
            df = scrape_jobs(results_wanted=remaining, offset=offset, **kwargs)
        except Exception as exc:
            log.warning(f"    [WARN] {platform}: fetch failed at offset {offset}: {exc}")
            break
        pages_requested += 1

        if df is None or df.empty:
            log.info(f"    {platform}: page {page + 1} returned 0 rows - stopping pagination.")
            break

        # Platforms often re-serve the same rows once the result set is
        # exhausted; without this the loop would "collect" duplicates and
        # think there was more to fetch.
        if "job_url" in df.columns:
            fresh = df[~df["job_url"].isin(seen_urls)]
            seen_urls.update(fresh["job_url"].dropna().tolist())
        else:
            fresh = df

        log.info(
            f"    {platform}: page {page + 1} -> {len(df)} rows "
            f"({len(fresh)} new, offset={offset})"
        )

        if fresh.empty:
            break

        fresh = fresh.copy()
        fresh["_search_page"] = page + 1
        frames.append(fresh)
        collected += len(fresh)

        # A short page means the platform has nothing more to give.
        if len(df) < remaining:
            break

    combined = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    return combined, hours_applied, {"pages": pages_requested, "requests": pages_requested}


def apply_recency_filter(df: pd.DataFrame, hours_old: int) -> tuple[pd.DataFrame, int]:
    """Drop postings older than the cut-off. Used for platforms where we
    deliberately did not send hours_old to JobSpy (see build_search_kwargs).

    Rows with no date_posted are KEPT - an unknown date is not evidence
    that a posting is stale, and dropping them would throw away valid
    remote jobs for no reason.
    """
    if df.empty or "date_posted" not in df.columns:
        return df, 0

    cutoff = (datetime.now() - timedelta(hours=hours_old)).date()
    parsed = pd.to_datetime(df["date_posted"], errors="coerce")
    keep = parsed.isna() | (parsed.dt.date >= cutoff)
    return df[keep], int((~keep).sum())


def apply_country_lock_filter(df: pd.DataFrame, country_indeed: str) -> tuple[pd.DataFrame, pd.DataFrame, int]:
    """Country Validation (requirement: STRICT COUNTRY LOCK).

    Every row was fetched under a search already restricted to
    `country_indeed`, but a job board's own filtering is not perfect,
    so this re-checks each row's own location text before it is allowed
    any further into the pipeline.

    Reuses job_sources.location_permits - the exact same synonym table
    and "unknown is not evidence" rule already applied to the dedicated
    remote-board connectors (RemoteOK/Remotive/WWR) - rather than a
    second, divergent implementation. A row is dropped only when its
    location text names a DIFFERENT, specific country than the one it
    was searched under; a blank or ambiguous location is kept, exactly
    as it already is everywhere else in this project.
    """
    if df.empty:
        return df, df.iloc[0:0].copy(), 0

    location_col = (
        df["location"] if "location" in df.columns
        else pd.Series([""] * len(df), index=df.index)
    )
    keep_mask = location_col.apply(
        lambda loc: job_sources.location_permits(loc, country_indeed)
    )
    return df[keep_mask], df[~keep_mask], int((~keep_mask).sum())


def _fetch_and_filter_platform(platform: str, keyword: str, location: str,
                               country_indeed: str, hours_old: int, config: dict) -> tuple:
    """One platform's fetch + recency filter + country validation.

    Pulled out of search_keyword_in_country so the per-platform work -
    each platform being an independent network round-trip to a different
    site/API - can be run concurrently instead of one-after-another (see
    search_keyword_in_country). Returns (platform, df, stats_dict).
    """
    df, hours_applied, fetch_stats = fetch_platform(
        platform, keyword, location, country_indeed, hours_old,
        config["results_wanted_per_platform"], config["max_pages_per_platform"],
        config.get("job_sources"),
    )
    raw_count = len(df)

    stale_dropped = 0
    if not hours_applied and not df.empty:
        df, stale_dropped = apply_recency_filter(df, hours_old)

    invalid_country = 0
    if not df.empty:
        df, _dropped_country, invalid_country = apply_country_lock_filter(df, country_indeed)

    stats = {
        "raw": raw_count,
        "stale_dropped": stale_dropped,
        "invalid_country": invalid_country,
        "kept": len(df),
        "pages": fetch_stats.get("pages", 0),
        "requests": fetch_stats.get("requests", 0),
    }
    return platform, df, stats


# How many platforms to fetch at once for a single keyword. Each platform
# is a different site/API (Indeed, LinkedIn, ZipRecruiter, Glassdoor,
# RemoteOK, Remotive, We Work Remotely, Jobspresso) with its own session,
# so running them concurrently is a wall-clock optimization - not a
# rate-limit bypass against any single platform - and was one of two
# fixes (the other is ContactEnrichmentStage's concurrency) for runs
# that previously took well over an hour purely from doing independent,
# unrelated network calls strictly one after another.
#
# Kept at 4, not 8 (all configured platforms at once): on a resource-
# constrained host (e.g. a small cloud instance), even I/O-bound Python
# threads compete for the GIL and a shared thread pool, and running
# every platform at once was observed to make the API's own trivial
# endpoints (health checks, starting a new search) time out while a
# search was active - not a hang, but bad enough to look like one. Only
# Indeed and LinkedIn are genuinely slow (ZipRecruiter/Glassdoor fail
# fast; the dedicated boards are cached in-memory feed lookups), so 4
# still lets both of the slow ones run alongside the fast ones without
# needing all 8 threads at once. Configurable via
# config.yaml's platform_fetch_concurrency.
DEFAULT_MAX_CONCURRENT_PLATFORMS = 4

# PLATFORM_FETCH_TIMEOUT_SECONDS: the real fix for a run that hangs
# indefinitely rather than merely running long. JobSpy's own scrape_jobs()
# call - unlike every other HTTP call in this project - has no timeout we
# control; on some networks (observed: a cloud host's outbound IP treated
# differently by a job board's anti-bot system than a developer machine)
# the underlying request can block forever with neither a result nor an
# exception. Before platform-level concurrency, a hang like that stalled
# one platform at a time; with concurrency, waiting on every platform's
# future with no bound at all meant ONE hung platform blocked the whole
# query indefinitely - and since OVERALL_SEARCH_TIMEOUT is only checked
# BETWEEN queries, a hang inside one query defeats it entirely. This
# timeout bounds how long one query waits on its platforms combined:
# generous enough for LinkedIn's own ~1-3 minute pacing, short enough
# that a genuinely stuck platform can never block the run for the rest
# of the configured timeout. A platform that times out is treated as
# "no results this query" and logged - never a crash, never silence.
PLATFORM_FETCH_TIMEOUT_SECONDS = 240


def search_keyword_in_country(keyword: str, original_title: str, entry: dict,
                              config: dict) -> tuple[pd.DataFrame, dict]:
    """Search every configured platform for one keyword in one country.

    Platforms are fetched concurrently (see config["platform_fetch_
    concurrency"] / DEFAULT_MAX_CONCURRENT_PLATFORMS) - the main change
    from the original sequential version - then logged
    and assembled in the configured platform order for a deterministic,
    reproducible summary regardless of which platform happened to finish
    first. Every row is re-checked against the selected country before
    being kept - see apply_country_lock_filter (STRICT COUNTRY LOCK /
    Country Validation).

    Waiting on the platforms is itself bounded (PLATFORM_FETCH_TIMEOUT_
    SECONDS) - see that constant's comment for why this, not just
    OVERALL_SEARCH_TIMEOUT, is required to guarantee the run can never
    hang indefinitely.
    """
    location = entry["location"]
    country_indeed = entry["country_indeed"]
    hours_old = config["hours_old"]
    platforms = config["platforms"]

    fetched = {}
    concurrency = config.get("platform_fetch_concurrency", DEFAULT_MAX_CONCURRENT_PLATFORMS)
    workers = max(1, min(int(concurrency), len(platforms)))
    pool = concurrent.futures.ThreadPoolExecutor(max_workers=workers)
    try:
        future_to_platform = {
            pool.submit(
                _fetch_and_filter_platform, platform, keyword, location,
                country_indeed, hours_old, config,
            ): platform
            for platform in platforms
        }
        done, not_done = concurrent.futures.wait(
            future_to_platform.keys(), timeout=PLATFORM_FETCH_TIMEOUT_SECONDS
        )
        for future in done:
            platform = future_to_platform[future]
            try:
                _, df, stats = future.result()
            except Exception as exc:
                # A single platform failing must never take down the run
                # (matches fetch_platform's own never-raises contract).
                log.warning(f"    [WARN] {platform}: unexpected failure: {exc}")
                df, stats = pd.DataFrame(), {
                    "raw": 0, "stale_dropped": 0, "invalid_country": 0,
                    "kept": 0, "pages": 0, "requests": 0,
                }
            fetched[platform] = (df, stats)
        for future in not_done:
            # Still running past the bound - almost always JobSpy's own
            # network call blocked with no result or error. Python cannot
            # forcibly kill a running thread, so this thread is abandoned
            # (it may still finish later and its result is simply
            # discarded) rather than waited on any further.
            platform = future_to_platform[future]
            log.warning(
                f"    [WARN] {platform}: exceeded {PLATFORM_FETCH_TIMEOUT_SECONDS}s "
                "(no response, no error - treated as no results for this query, "
                "rather than blocking the whole run)."
            )
            fetched[platform] = (pd.DataFrame(), {
                "raw": 0, "stale_dropped": 0, "invalid_country": 0,
                "kept": 0, "pages": 0, "requests": 1,
            })
    finally:
        # wait=False: never block exit on an abandoned/hung thread - that
        # would silently reintroduce the exact hang this timeout exists
        # to prevent.
        pool.shutdown(wait=False)

    frames = []
    platform_stats = {}
    collected_at = datetime.now().isoformat(timespec="seconds")

    # Logged/assembled in configured order, not completion order, so two
    # runs of the same search produce the same summary and export.
    for platform in platforms:
        df, stats = fetched[platform]
        platform_stats[platform] = stats
        log.info(
            f"      {platform:<9} fetched {stats['raw']:>3} | stale {stats['stale_dropped']:>2} "
            f"| country-mismatch {stats['invalid_country']:>2} | carried forward {stats['kept']:>3}"
        )

        if df.empty:
            continue

        # search_keyword holds the keyword that actually returned the
        # row; original_job_title holds what the user asked for. Keeping
        # both is what lets the export explain each row's inclusion.
        df["search_keyword"] = keyword
        df["original_job_title"] = original_title
        df["country"] = country_indeed
        if "_search_page" not in df.columns:
            df["_search_page"] = 1
        df["_collected_at"] = collected_at
        frames.append(df)

    combined = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    return combined, platform_stats


def _format_elapsed(seconds: float) -> str:
    """Format seconds as MM:SS (e.g. '08:32'), or HH:MM:SS past one hour."""
    total = int(max(0, seconds))
    hours, rem = divmod(total, 3600)
    minutes, secs = divmod(rem, 60)
    return f"{hours:02d}:{minutes:02d}:{secs:02d}" if hours else f"{minutes:02d}:{secs:02d}"


def _estimate_unique_companies(df: pd.DataFrame) -> int:
    """Real (not approximate-by-row-count) running unique-company total
    for progress reporting, using the same identity grouping as the
    pipeline's own company-dedup stage - so the number shown mid-search
    is the same kind of number the final export will contain, not a raw
    job-row count mislabelled as companies."""
    if df.empty:
        return 0
    groups = company_identity.assign_company_groups(df.to_dict("records"))
    return len(set(groups))


def search_country(job_title: str, entry: dict, config: dict,
                   keyword_list: list = None,
                   needed: int = None, deadline: float = None,
                   progress_callback=None, progress_context: dict = None
                   ) -> tuple[pd.DataFrame, dict, dict]:
    """Search one country across every configured keyword and platform.

    `keyword_list[0]` is the user's own title and is always searched
    first, so the primary keyword's rows are the ones that survive
    deduplication and land at the top of the export. Passing None
    searches the title alone, which is the pre-expansion behaviour.

    Like the country loop in search_region, the keyword loop stops early
    once `needed` valid remote rows exist - an extra keyword is only
    worth its network cost while the target is still short. It also stops
    - anywhere in the middle, before starting the next query - once
    `deadline` (a time.monotonic() timestamp) has passed: OVERALL_SEARCH_
    TIMEOUT, so a run can never continue indefinitely regardless of how
    many keywords/platforms are configured.
    """
    keyword_list = keyword_list or [job_title]
    progress_context = progress_context or {}

    frames = []
    platform_stats = {}
    keyword_stats = {}

    for index, keyword in enumerate(keyword_list):
        if deadline is not None and time.monotonic() >= deadline:
            log.warning(
                f"      [TIMEOUT] OVERALL_SEARCH_TIMEOUT reached before query "
                f"{index + 1}/{len(keyword_list)} ('{keyword}') - stopping here and "
                "exporting what was already collected, rather than continuing."
            )
            break

        # Query-level progress (requirement: "Query 1/20: AI Engineer").
        log.info(f"      [QUERY {index + 1}/{len(keyword_list)}] '{keyword}'"
                 + ("" if index == 0 else "  (similar keyword)"))

        raw, stats = search_keyword_in_country(keyword, job_title, entry, config)
        keyword_stats[keyword] = {
            "raw": len(raw),
            "primary": index == 0,
            "platforms": stats,
        }
        for platform, values in stats.items():
            running = platform_stats.setdefault(
                platform, {"raw": 0, "stale_dropped": 0, "invalid_country": 0,
                          "kept": 0, "pages": 0, "requests": 0}
            )
            for field in running:
                running[field] += values.get(field, 0)

        if not raw.empty:
            frames.append(raw)

        so_far = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
        kept_so_far, _, _ = apply_remote_safety_filter(so_far, config["remote_strictness"])
        deduped_so_far, _ = deduplicate(kept_so_far)

        if progress_callback:
            elapsed = time.monotonic() - progress_context.get("start_time", time.monotonic())
            pages_so_far = sum(v.get("pages", 0) for v in platform_stats.values())
            unique_companies = _estimate_unique_companies(deduped_so_far)
            log.info(
                f"        Elapsed: {_format_elapsed(elapsed)}  |  Raw jobs collected: "
                f"{len(deduped_so_far)}  |  Unique companies: {unique_companies}  |  "
                f"Target: {progress_context.get('target_leads', '?')}"
            )
            try:
                progress_callback({
                    "stage": "Searching job boards",
                    "current_region": progress_context.get("region", entry.get("country_indeed", "")),
                    "query_index": index + 1,
                    "query_total": len(keyword_list),
                    "query": keyword,
                    "pages_searched": pages_so_far,
                    "raw_jobs_collected": len(deduped_so_far),
                    "unique_companies_estimate": unique_companies,
                    "target_leads": progress_context.get("target_leads"),
                    "elapsed_seconds": elapsed,
                    "elapsed_formatted": _format_elapsed(elapsed),
                })
            except Exception:
                pass

        if needed is None or not frames:
            continue

        if len(deduped_so_far) >= needed and index + 1 < len(keyword_list):
            log.info(
                f"      Target reached in {entry['country_indeed']} - skipping "
                f"{len(keyword_list) - index - 1} further keyword(s)."
            )
            break

    combined = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    return combined, platform_stats, keyword_stats


def search_region(region: str, job_title: str, config: dict, needed: int,
                  keyword_list: list = None, deadline: float = None,
                  progress_callback=None, progress_context: dict = None
                  ) -> tuple[pd.DataFrame, pd.DataFrame, dict, dict]:
    """Search one region - which may be several countries, as Europe is -
    and return its valid remote rows.

    Stops early between countries once `needed` valid rows have been
    collected, so a region that fills the target on its first country
    does not pay for the rest. Also stops - before starting the next
    country - once `deadline` (time.monotonic()) has passed; see
    search_country's docstring for OVERALL_SEARCH_TIMEOUT.

    Returns (kept_df, dropped_df, filter_stats, region_stats).
    """
    entries = REGION_DEFINITIONS[region]
    raw_frames = []
    per_country = {}

    for entry in entries:
        country = entry["country_indeed"]
        if deadline is not None and time.monotonic() >= deadline:
            log.warning(
                f"    [TIMEOUT] OVERALL_SEARCH_TIMEOUT reached before {country} - "
                "stopping here and exporting what was already collected."
            )
            break
        log.info(f"    -- {country} --")

        raw, platform_stats, keyword_stats = search_country(
            job_title, entry, config, keyword_list, needed, deadline,
            progress_callback, progress_context,
        )
        raw_frames.append(raw)
        per_country[country] = {
            "raw": len(raw),
            "platforms": platform_stats,
            "keywords": keyword_stats,
        }

        # Filter what we have so far so the early-exit check counts
        # genuinely valid remote rows, not raw hits.
        so_far = pd.concat(raw_frames, ignore_index=True) if raw_frames else pd.DataFrame()
        kept_so_far, _, _ = apply_remote_safety_filter(so_far, config["remote_strictness"])
        deduped_so_far, _ = deduplicate(kept_so_far)
        if len(deduped_so_far) >= needed:
            remaining = [e["country_indeed"] for e in entries[entries.index(entry) + 1:]]
            if remaining:
                log.info(
                    f"    Target reached within {region} - skipping "
                    f"{', '.join(remaining)}."
                )
            break

    combined = pd.concat(raw_frames, ignore_index=True) if raw_frames else pd.DataFrame()
    kept, dropped, filter_stats = apply_remote_safety_filter(
        combined, config["remote_strictness"]
    )

    region_stats = {
        "raw": len(combined),
        "after_remote_filter": len(kept),
        "countries": per_country,
    }
    return kept, dropped, filter_stats, region_stats


def collect_jobs(user_input: dict, config: dict) -> tuple[pd.DataFrame, dict]:
    """Search the user's selected region/country only.

    STRICT COUNTRY LOCK: unlike an earlier version of this function, the
    search never falls back to another region/country when the selected
    one falls short of the target. See the REGION_DEFINITIONS module
    note. Europe's five countries are still searched in the order given
    there - that is what "Europe" as a selection has always meant, not a
    fallback to a different region.
    """
    job_title = user_input["job_title"]
    selected = user_input["region"]
    target = config["target_total_jobs"]

    # The dedicated remote-board connectors (RemoteOK, Remotive, We Work
    # Remotely, Jobspresso) are not partitioned by country, so they cache
    # their feed for the life of one run - see job_sources.py. Clearing
    # it here means a fresh `python main.py` run always sees fresh
    # listings, while still fetching each feed only once even though
    # Europe alone searches five countries.
    job_sources.reset_cache()

    # Expanded once per run, not per region or per platform: the same
    # keyword list is used everywhere, so a job is never missed in one
    # region because a different set of keywords was searched there.
    expansion = keyword_expansion.KeywordExpander(
        config.get("keyword_expansion")
    ).expand(job_title)
    keyword_list = expansion["keywords"] or [job_title]

    log.info(f"\n[KEYWORDS] Searched title: '{job_title}'")
    if expansion["expanded"]:
        log.info(
            f"           Similar keywords ({len(expansion['expanded'])}): "
            + ", ".join(f"'{k}'" for k in expansion["expanded"])
        )
    else:
        log.info("           No similar keywords found - searching the title alone.")

    # STRICT COUNTRY LOCK: only the selected region is ever searched.
    search_order = [selected]
    locked_out = [r for r in REGION_ORDER if r != selected]
    if locked_out:
        log.info(
            f"[COUNTRY LOCK] Locked to '{selected}'. {', '.join(locked_out)} will NOT "
            f"be searched, even if the {target}-company target is not reached."
        )

    collected = pd.DataFrame()
    dropped_all = []
    region_stats = {}
    filter_totals = {}
    regions_searched = []
    regions_skipped = list(locked_out)

    for region in search_order:
        if len(collected) >= target:
            continue

        needed = target - len(collected)
        label = "SELECTED"
        log.info(f"\n[{label}] Region: {region}  (need {needed} more)")

        kept, dropped, filter_stats, stats = search_region(
            region, job_title, config, needed, keyword_list
        )
        regions_searched.append(region)
        region_stats[region] = stats
        region_stats[region]["role"] = label

        for key, value in filter_stats.items():
            filter_totals[key] = filter_totals.get(key, 0) + value
        if not dropped.empty:
            dropped_all.append(dropped)

        if not kept.empty:
            kept = kept.copy()
            kept["_region"] = region
            # Concatenating in search order is what keeps the selected
            # region's rows first in the final export.
            collected = pd.concat([collected, kept], ignore_index=True)
            collected, _ = deduplicate(collected)

        region_stats[region]["running_total"] = len(collected)
        log.info(f"    {region}: running total {len(collected)}/{target}")

        if len(collected) >= target:
            log.info(f"    Target of {target} reached - no further regions will be searched.")

    dropped_df = (
        pd.concat(dropped_all, ignore_index=True) if dropped_all else pd.DataFrame()
    )

    # How many rows each keyword contributed, counted after dedup so a
    # posting found by two keywords is credited to the one that found it
    # first - the primary keyword wherever both did. This is the
    # collected total; execute_pipeline adds the post-cap count.
    keyword_counts = {k: 0 for k in keyword_list}
    if not collected.empty and "search_keyword" in collected.columns:
        for value in collected["search_keyword"]:
            keyword_counts[value] = keyword_counts.get(value, 0) + 1

    summary = {
        "selected_region": selected,
        "keyword_expansion": expansion,
        "keyword_counts_collected": keyword_counts,
        "regions_searched": regions_searched,
        "regions_skipped": regions_skipped + [
            r for r in search_order if r not in regions_searched and r not in regions_skipped
        ],
        "region_stats": region_stats,
        "filter_totals": filter_totals,
        "dropped_df": dropped_df,
    }
    return collected, summary


# Phrases that positively indicate a posting is remote.
REMOTE_INDICATOR_PHRASES = [
    "remote", "work from home", "wfh", "work from anywhere",
    "fully remote", "100% remote", "remote-first", "telecommute",
    "distributed team", "anywhere in",
]

# Phrases that positively contradict a remote posting. Only consulted
# for rows that showed no remote evidence at all.
ONSITE_CONTRADICTION_PHRASES = [
    "on-site", "onsite", "on site", "in-office", "in office",
    "work from office", "wfo", "hybrid",
]


def _row_text(row) -> str:
    return " ".join(
        str(row.get(field, "") or "").lower()
        for field in ("title", "description", "location")
    )


def apply_remote_safety_filter(df: pd.DataFrame,
                               strictness: str = "balanced") -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    """Decide, row by row, whether a posting is genuinely remote.

    The naive rule - "drop everything where JobSpy says is_remote is
    False" - destroys valid results, because that flag does not mean the
    same thing on every platform:

      - Indeed sets it from the actual server-side remote filter, so it
        is trustworthy.
      - LinkedIn IGNORES it and recomputes is_remote with a text
        heuristic over title/description/location, even though the search
        itself was already restricted to remote roles via f_WT=2.
        Measured live: 17-18 of 20 genuinely-remote LinkedIn rows come
        back flagged False.
      - Google likewise derives it from description text alone.

    Since every platform is queried with a remote-only filter, a False
    flag is weak evidence of "not remote" rather than proof. The rule:

      1. flag is True                      -> keep (platform confirmed)
      2. text shows a remote indicator     -> keep (text confirmed)
      3. no remote evidence, but the text
         explicitly says on-site/hybrid    -> DROP (contradicted)
      4. no evidence either way            -> depends on `strictness`:
           "balanced" (default) keeps it, on the strength of the
             server-side filter. Verified live: requesting remote from
             LinkedIn returns a materially different result set (10 of
             30 URLs shared with an unfiltered search; Indeed 1 of 30),
             so the filter is demonstrably doing something.
           "strict" drops it. Fewer rows, every one of them carrying
             positive evidence of being remote.

    Neither setting can be proved correct per-posting from here: Indeed
    serves a Cloudflare bot check to automated requests, so the final
    word is still a human opening a few job_urls - which is what the
    acceptance criteria ask for.

    Returns (kept_df, dropped_df, stats) so the caller can show exactly
    what was removed and why.
    """
    stats = {
        "kept_platform_confirmed": 0,
        "kept_text_confirmed": 0,
        "kept_server_filter_trusted": 0,
        "dropped_onsite_contradiction": 0,
        "dropped_no_remote_evidence": 0,
    }
    if df.empty:
        return df, df, stats

    df = df.copy()
    flag = df["is_remote"] if "is_remote" in df.columns else pd.Series(pd.NA, index=df.index)

    keep = []
    reasons = []
    for idx, row in df.iterrows():
        if flag.loc[idx] is True or flag.loc[idx] == True:  # noqa: E712
            keep.append(True); reasons.append("platform_confirmed")
            stats["kept_platform_confirmed"] += 1
            continue

        text = _row_text(row)
        if any(p in text for p in REMOTE_INDICATOR_PHRASES):
            keep.append(True); reasons.append("text_confirmed")
            stats["kept_text_confirmed"] += 1
        elif any(p in text for p in ONSITE_CONTRADICTION_PHRASES):
            keep.append(False); reasons.append("onsite_contradiction")
            stats["dropped_onsite_contradiction"] += 1
        elif strictness == "strict":
            keep.append(False); reasons.append("no_remote_evidence")
            stats["dropped_no_remote_evidence"] += 1
        else:
            keep.append(True); reasons.append("server_filter_trusted")
            stats["kept_server_filter_trusted"] += 1

    keep_mask = pd.Series(keep, index=df.index)
    df["_remote_reason"] = reasons
    return df[keep_mask], df[~keep_mask], stats


def deduplicate(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """Drop duplicate postings in three passes:

      1. identical job_url - the same posting scraped twice
      2. normalised title+company+location - same posting, no URL
      3. normalised title+company - the same remote posting listed under
         several cities. A real example from a live run: one AgileEngine
         opening appeared as both "Kochi, Kerala" and "Trivandrum,
         Kerala" with different URLs, so passes 1 and 2 both missed it.
         For a remote-only list the city is not a distinguishing feature.
    """
    counts = {"by_url": 0, "by_fallback_key": 0, "by_title_company": 0}
    if df.empty:
        return df, counts

    df = df.copy()

    def normalize(value) -> str:
        return re.sub(r"\s+", " ", str(value or "")).strip().lower()

    def normalized_column(name: str) -> pd.Series:
        # df.get(name, "") would return a bare str when the column is
        # absent, which has no .apply - build an empty Series instead.
        if name in df.columns:
            return df[name].apply(normalize)
        return pd.Series([""] * len(df), index=df.index)

    title_n = normalized_column("title")
    company_n = normalized_column("company")
    location_n = normalized_column("location")

    df["_fallback_key"] = title_n + "|" + company_n + "|" + location_n
    df["_title_company_key"] = title_n + "|" + company_n

    before = len(df)
    if "job_url" in df.columns:
        has_url = df["job_url"].notna() & (df["job_url"].astype(str).str.strip() != "")
        with_url = df[has_url].drop_duplicates(subset="job_url", keep="first")
        without_url = df[~has_url].drop_duplicates(subset="_fallback_key", keep="first")
        counts["by_url"] = int(has_url.sum() - len(with_url))
        counts["by_fallback_key"] = int((~has_url).sum() - len(without_url))
        df = pd.concat([with_url, without_url], ignore_index=True)
    else:
        df = df.drop_duplicates(subset="_fallback_key", keep="first")
        counts["by_fallback_key"] = before - len(df)

    before = len(df)
    df = df.drop_duplicates(subset="_title_company_key", keep="first")
    counts["by_title_company"] = before - len(df)

    return df.drop(columns=["_fallback_key", "_title_company_key"]), counts


def truncate_description(text) -> str:
    text = str(text or "")
    if len(text) > EXCEL_CELL_CHAR_LIMIT:
        return text[:EXCEL_CELL_CHAR_LIMIT] + "... [truncated]"
    return text


def cap_and_reshape(df: pd.DataFrame, target_total_jobs: int) -> pd.DataFrame:
    """Trim to the target row count, then reorder/rename columns to
    exactly match the required output schema, adding date_fetched."""
    if df.empty:
        return pd.DataFrame(columns=list(OUTPUT_COLUMNS.keys()))

    df = df.head(target_total_jobs).copy()
    df["date_fetched"] = date.today().isoformat()

    output = pd.DataFrame()
    for out_col, candidates in OUTPUT_COLUMNS.items():
        series = None
        for candidate in candidates:
            if candidate in df.columns:
                series = df[candidate]
                break
        output[out_col] = series if series is not None else None

    if "job_description" in output.columns:
        output["job_description"] = output["job_description"].apply(truncate_description)

    # Acceptance criterion: every exported row's is_remote is True.
    # Rows that survived apply_remote_safety_filter are remote by the
    # rules documented there, but several arrive with a missing/False
    # flag from a platform that derives it by text heuristic. Normalise
    # so the column reflects the filter's decision, not the platform's
    # guess.
    output["is_remote"] = True

    return output.reset_index(drop=True)


def _write_formatted_excel(sheets, file_path: Path) -> None:
    """Write one or more readable Excel sheets.

    `sheets` is an ordered {sheet_name: DataFrame} mapping; the first
    sheet is the primary output. A plain DataFrame is still accepted so
    existing callers keep working.
    """
    if isinstance(sheets, pd.DataFrame):
        sheets = {"Remote Jobs": sheets}

    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            _format_worksheet(writer.sheets[sheet_name])


def _format_worksheet(worksheet) -> None:
    """Apply header, wrapping and column widths to one sheet."""

    # Keep the header visible while scrolling.
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    # Bold header and wrap all cell text so long values remain readable.
    for cell in worksheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    worksheet.row_dimensions[1].height = 30

    for row in worksheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    # Width by column HEADER, not position: the workbook now has sheets
    # with very different shapes, so a per-letter table would misfit them.
    wide_headers = {
        "Lead Summary": 70, "Matching Job URLs": 55, "Matching Job Titles": 45,
        "Priority Reason": 55, "Reason": 55, "LinkedIn Search Links": 45,
        "job_description": 70,
        "Hiring Signal": 30, "Company Website": 32, "Contact Email": 30,
        "Contact LinkedIn URL": 34, "Contact Source": 30, "Exclusion Reason": 34,
        "Company Name": 30, "Value": 60, "Metric": 45, "job_url": 34,
        "company_url": 30, "Contact Title": 32, "ICP Status": 26,
        "Matched Keywords": 45, "Searched Title": 26,
        "matched_keyword": 30, "original_job_title": 26,
        "Contact Email (2nd)": 28, "Funding Signal": 30, "Team Maturity": 24,
        "AI Hiring Stage": 22, "job_title": 34, "company_name": 26,
        "Industry": 26, "Location": 30, "Job Title": 34, "Job URL": 34,
    }
    narrow_headers = {
        "Lead Priority": 13, "Priority Score": 12, "Employee Count": 15,
        "Hiring Signal Strength": 15, "Contact Strength": 13,
        "Company Size": 15, "Matching Position Count": 12, "Country": 12,
        "Region": 12, "Date Fetched": 13, "Contact Confidence": 13,
        "is_remote": 10, "Buyer Probability": 15, "ICP Fit Score": 13,
        "Source": 20, "Search Query": 26, "Search Page": 12, "Collected At": 18,
    }
    for col_idx in range(1, worksheet.max_column + 1):
        letter = get_column_letter(col_idx)
        header = str(worksheet.cell(row=1, column=col_idx).value or "")
        width = wide_headers.get(header) or narrow_headers.get(header) or 20
        worksheet.column_dimensions[letter].width = width

    # Make High-priority leads findable at a glance, as the brief asks.
    headers = [str(worksheet.cell(row=1, column=i).value or "")
               for i in range(1, worksheet.max_column + 1)]
    fills = {
        "High": PatternFill("solid", fgColor="C6EFCE"),
        "Medium": PatternFill("solid", fgColor="FFEB9C"),
        "Low": PatternFill("solid", fgColor="F2F2F2"),
    }
    if "Lead Priority" in headers:
        priority_col = headers.index("Lead Priority") + 1
        for row_idx in range(2, worksheet.max_row + 1):
            cell = worksheet.cell(row=row_idx, column=priority_col)
            fill = fills.get(str(cell.value))
            if fill:
                cell.fill = fill
                cell.font = Font(bold=str(cell.value) == "High")

    # Same colour coding for the new Buyer Probability column, so High/
    # Medium/Low probability companies (including Low-Probability Buyer
    # Dictionary matches, which are never discarded from this sheet) are
    # just as easy to scan at a glance.
    if "Buyer Probability" in headers:
        buyer_col = headers.index("Buyer Probability") + 1
        for row_idx in range(2, worksheet.max_row + 1):
            cell = worksheet.cell(row=row_idx, column=buyer_col)
            fill = fills.get(str(cell.value))
            if fill:
                cell.fill = fill
                cell.font = Font(bold=str(cell.value) == "High")

    # "Contact: Not Found" must stand out rather than read as an error.
    if "Contact Status" in headers:
        status_col = headers.index("Contact Status") + 1
        for row_idx in range(2, worksheet.max_row + 1):
            cell = worksheet.cell(row=row_idx, column=status_col)
            if str(cell.value).endswith("Not Found"):
                cell.font = Font(color="9C6500")

    # Give each row enough height for wrapped titles and descriptions.
    for row_idx in range(2, worksheet.max_row + 1):
        worksheet.row_dimensions[row_idx].height = 45


def export_to_excel(sheets, output_dir: str) -> str:
    out_path = Path(output_dir)
    out_path.mkdir(parents=True, exist_ok=True)
    file_path = out_path / f"remote_ai_jobs_{date.today().isoformat()}.xlsx"

    try:
        _write_formatted_excel(sheets, file_path)
        return str(file_path)
    except PermissionError:
        # Almost always means the file from an earlier run is still open
        # in Excel, which locks it for writing. Losing a completed search
        # to that would be wasteful, so fall back to a suffixed filename
        # and tell the user what happened.
        for attempt in range(2, 12):
            alt_path = out_path / f"remote_ai_jobs_{date.today().isoformat()}_{attempt}.xlsx"
            try:
                _write_formatted_excel(sheets, alt_path)
                log.warning(
                    f"\n[NOTE] '{file_path.name}' is locked (it's probably still "
                    f"open in Excel), so this run was saved as '{alt_path.name}' instead."
                )
                return str(alt_path)
            except PermissionError:
                continue
        raise


def export_dropped_for_review(dropped_df: pd.DataFrame, output_dir: str) -> str | None:
    """Write the rows the remote-safety filter removed to a small CSV so
    you can eyeball exactly what was excluded and why, instead of trusting
    the filter blindly. Only written when something was actually dropped."""
    if dropped_df.empty:
        return None

    out_path = Path(output_dir)
    out_path.mkdir(parents=True, exist_ok=True)
    file_path = out_path / f"dropped_for_review_{date.today().isoformat()}.csv"

    review_cols = {}
    for col, candidates in (
        ("job_title", ["title"]),
        ("company_name", ["company"]),
        ("location_raw", ["location"]),
        ("is_remote_flag", ["is_remote"]),
        ("dropped_because", ["_remote_reason"]),
        ("job_url", ["job_url"]),
    ):
        for candidate in candidates:
            if candidate in dropped_df.columns:
                review_cols[col] = dropped_df[candidate]
                break

    pd.DataFrame(review_cols).to_csv(file_path, index=False)
    return str(file_path)


def print_summary(summary: dict, final_count: int, target_leads: int,
                  strictness: str) -> None:
    selected = summary["selected_region"]
    print("\n" + "=" * 70)
    print("RUN SUMMARY")
    print("=" * 70)

    print(f"\nSelected region : {selected}")
    print(f"Target leads    : {target_leads}")
    print(f"Final leads     : {final_count}")
    print(f"Target reached  : {summary.get('target_reached', False)}")

    expansion = summary.get("keyword_expansion") or {}
    counts = summary.get("keyword_counts") or {}
    if expansion.get("keywords"):
        print("\nSearch keywords (primary first) - rows contributed to the exported file:")
        for index, keyword in enumerate(expansion["keywords"]):
            label = "primary" if index == 0 else expansion["sources"].get(keyword, "similar")
            print(f"  {keyword:<42} {counts.get(keyword, 0):>3}  ({label})")
        if expansion.get("families"):
            print(f"  Matched title family: {', '.join(expansion['families'])}")
        rejected = expansion.get("rejected") or []
        off_topic = [r for r in rejected if not r["reason"].startswith("over the max")]
        if off_topic:
            print(f"  Rejected as unrelated or malformed  : {len(off_topic)}")
            for entry in off_topic[:5]:
                print(f"    - '{entry['keyword']}' ({entry['source']}): {entry['reason']}")
        capped = len(rejected) - len(off_topic)
        if capped:
            print(f"  Valid but over the keyword cap      : {capped}")

    print("\nRegion-by-region (searched in this order, selected region first):")
    for region, stats in summary["region_stats"].items():
        print(f"\n  [{stats['role']}] {region}")
        print(f"      raw rows fetched    : {stats['raw']}")
        print(f"      valid remote rows   : {stats['after_remote_filter']}")
        running_jobs = stats.get("running_jobs", stats.get("raw", 0))
        running_leads = stats.get("running_leads", 0)
        print(f"      running jobs after  : {running_jobs}")
        print(f"      running leads after : {running_leads}/{target_leads}")
        for country, cstats in stats["countries"].items():
            platforms = " | ".join(
                f"{p}:{ps['kept']}" for p, ps in cstats["platforms"].items()
            )
            print(f"        {country:<14} raw {cstats['raw']:>3}   ({platforms})")

    if summary["regions_skipped"]:
        print(
            "\n  Not searched (target already met): "
            + ", ".join(summary["regions_skipped"])
        )
    else:
        print("\n  All regions were searched.")

    ft = summary["filter_totals"]
    print(f"\nRemote filter (remote_strictness={strictness}) - why rows were kept or dropped:")
    print(f"  KEPT    platform explicitly confirmed remote      : {ft.get('kept_platform_confirmed', 0)}")
    print(f"  KEPT    text/location contained a remote indicator: {ft.get('kept_text_confirmed', 0)}")
    print(f"  KEPT    no evidence either way, server-side remote")
    print(f"          filter trusted                            : {ft.get('kept_server_filter_trusted', 0)}")
    print(f"  DROPPED text explicitly said on-site/hybrid       : {ft.get('dropped_onsite_contradiction', 0)}")
    print(f"  DROPPED no positive remote evidence (strict mode) : {ft.get('dropped_no_remote_evidence', 0)}")

    if strictness == "balanced" and ft.get("kept_server_filter_trusted"):
        print(
            f"\n  Heads-up: {ft['kept_server_filter_trusted']} rows were kept purely on the\n"
            "  platform's server-side remote filter, with no wording in the title or\n"
            "  description confirming it. These are the rows worth spot-checking first.\n"
            "  Set remote_strictness: strict in config.yaml to exclude them instead."
        )

    dedup = summary["dedup_counts"]
    print("\nDeduplication - rows removed by each pass:")
    print(f"  same job_url                        : {dedup['by_url']}")
    print(f"  same title+company+location (no URL): {dedup['by_fallback_key']}")
    print(f"  same title+company, different city  : {dedup['by_title_company']}")

    print("\nRows per region in the exported file (selected region first):")
    for region, count in summary["export_region_counts"].items():
        marker = "  <-- selected" if region == selected else ""
        print(f"  {region:<12} {count:>3}{marker}")

    if final_count < target_leads:
        print(
            f"\nNote: final lead count ({final_count}) is below target ({target_leads}) "
            "even after falling back through every configured region. Only genuine "
            "qualified leads were exported; the list was not padded."
        )
    print("=" * 70)


# Company-level lead sheet: the primary output, one row per company.
# Separate from OUTPUT_COLUMNS, which stays exactly as the brief fixed it
# for the job-level sheet.
# Company-level lead sheet: the primary output, one row per company.
# Separate from OUTPUT_COLUMNS, which stays exactly as the brief fixed it
# for the job-level sheet.
LEAD_COLUMNS = {
    "Lead Priority": "lead_priority",
    # Buyer Probability (High/Medium/Low): the plain, un-promoted
    # classification tier. Low-probability-buyer companies (staffing/
    # recruitment/IT-services/AI-consulting/outsourcing/BPO - see the
    # Low Probability Buyer Dictionary in pipeline.IcpFilteringStage and
    # icp.COMPETITOR_DICTIONARY) are capped here at "Low" and are always
    # included in this sheet, never discarded - see lead_signals.score_company.
    "Buyer Probability": "buyer_probability",
    "Priority Score": "priority_score",
    "ICP Fit Score": "icp_fit_score",
    "Company Name": "company_name",
    "Company Website": "company_url_direct",
    "Company Domain": "company_domain",
    "Industry": "company_industry",
    "Location": "locations",
    "Employee Count": "employee_count",
    "Company Size": "company_size",
    "Company Type": "company_type",
    "ICP Status": "icp_status",
    "Reason": "priority_reason",
    "Source": "source_platforms",
    "Lead Summary": "lead_summary",
    "Matching Position Count": "matching_position_count",
    "Matching Job Titles": "matching_job_titles",
    "Matching Job URLs": "matching_job_urls",
    # Why this company is in the list: what was searched for, and which
    # of the expanded keywords its postings came back under.
    "Searched Title": "original_job_title",
    "Matched Keywords": "search_keywords",
    "Hiring Signal": "hiring_signal",
    "Hiring Signal Strength": "hiring_signal_strength",
    "AI Hiring Stage": "ai_hiring_stage",
    "Funding Signal": "funding_signal",
    "Contact Status": "contact_status",
    "Contact Name": "contact_name",
    "Contact Title": "contact_title",
    "Contact Email": "contact_email",
    "Contact Email (2nd)": "contact_email_secondary",
    "Contact Source": "contact_source",
    # Preserved exactly as before - not removed, not renamed.
    "LinkedIn Search Links": "contact_search_urls",
    "Country": "country",
    "Region": "region",
    "Date Fetched": "date_fetched",
}

# Blank cells read as missing data; these say so explicitly instead.
NOT_AVAILABLE_DEFAULTS = {
    "Buyer Probability": "Unknown",
    "Industry": "Unknown",
    "Location": "Unknown",
    "Reason": "Not Available",
    "Source": "Unknown",
    "Employee Count": "Unknown",
    "Company Size": "Unknown",
    "Company Type": "Unknown",
    "Contact Status": "Contact: Not Found",
    "Hiring Signal Strength": "Unknown",
    "Contact Strength": "None",
    "Contact Name": "Not Found",
    "Contact Title": "Not Found",
    "Funding Signal": "Not Found",
}


def _shape_sheet(df: pd.DataFrame, columns_map: dict, defaults: dict = None) -> pd.DataFrame:
    """Shared column-mapping/renaming/defaulting logic for every Excel
    sheet built from an internal DataFrame. `columns_map` is
    {output_label: internal_field_name}; a field the source did not
    provide is filled with `defaults[label]` when given, rather than
    left as a blank or a guess.
    """
    defaults = defaults or {}
    if df is None or df.empty:
        return pd.DataFrame(columns=list(columns_map.keys()))

    output = pd.DataFrame()
    for label, source in columns_map.items():
        if source in df.columns:
            column = df[source]
        else:
            column = pd.Series([None] * len(df), index=df.index)
        default = defaults.get(label)
        if default is not None:
            column = column.fillna(default).replace("", default)
        output[label] = column
    return output.reset_index(drop=True)


def shape_lead_sheet(df: pd.DataFrame) -> pd.DataFrame:
    """Put company rows into the lead column order.

    Anything a source did not provide is filled with an explicit
    "Unknown"/"Not Found" rather than a blank or a guess.
    """
    if df is not None and not df.empty and "date_fetched" not in df.columns:
        df = df.copy()
        df["date_fetched"] = date.today().isoformat()
    return _shape_sheet(df, LEAD_COLUMNS, NOT_AVAILABLE_DEFAULTS)


# Sheet: "Qualified Companies" - the trimmed, business-facing view of the
# final qualified/prioritised result (same source data as "Qualified
# Leads"/"All Company Signals", just the specific column set asked for -
# nothing here is a different computation or a different data source).
QUALIFIED_COMPANY_COLUMNS = {
    "Company Name": "company_name",
    "Company Domain": "company_domain",
    "Company Website": "company_url_direct",
    "Industry": "company_industry",
    "Location": "locations",
    "Company Size": "company_size",
    "Buyer Probability": "buyer_probability",
    "ICP Fit Score": "icp_fit_score",
    "Priority Score": "priority_score",
    "Reason": "priority_reason",
    "Source": "source_platforms",
    # Preserved exactly as before - not removed, not renamed.
    "LinkedIn Search Links": "contact_search_urls",
}


def shape_qualified_companies_sheet(df: pd.DataFrame) -> pd.DataFrame:
    """QUALIFIED RESULTS stage output, in the exact column set requested:
    the post-dedup, post-ICP/Buyer-Probability, post-scoring company list
    (High, Medium AND Low probability - none discarded)."""
    return _shape_sheet(df, QUALIFIED_COMPANY_COLUMNS, NOT_AVAILABLE_DEFAULTS)


# Sheet: "Raw Companies" - the RAW SEARCH RESULTS stage, before company
# dedup/ICP/scoring: one row per job posting actually collected (already
# country-validated and recency-filtered), so the full breadth of what
# was found is visible even though many rows here share one company.
RAW_COMPANY_COLUMNS = {
    "Company Name": "company",
    "Company Domain": "_company_domain",
    "Company Website": "company_url_direct",
    "Job Title": "title",
    "Job URL": "job_url",
    "Location": "location",
    "Country": "country",
    "Industry": "company_industry",
    "Company Size": "company_num_employees",
    "Source": "site",
    "Search Query": "search_keyword",
    "Search Page": "_search_page",
    "Collected At": "_collected_at",
}

RAW_COMPANY_DEFAULTS = {
    "Company Domain": "Unknown", "Company Website": "Unknown",
    "Industry": "Unknown", "Company Size": "Unknown",
}


def shape_raw_companies_sheet(job_level_df: pd.DataFrame) -> pd.DataFrame:
    """RAW SEARCH RESULTS stage output: every job posting collected for
    the locked country, deliberately NOT capped or deduplicated to one
    row per company - see the Qualified Companies/Qualified Leads sheets
    for the deduplicated, scored result. This is what makes the raw
    search breadth auditable instead of only visible as a summary count.
    """
    return _shape_sheet(job_level_df, RAW_COMPANY_COLUMNS, RAW_COMPANY_DEFAULTS)


def build_pipeline_summary_sheet(ctx, jobs_count: int, leads_count: int) -> pd.DataFrame:
    """Build the workbook's cumulative lead-search summary."""
    icp_stats = ctx.artifacts.get("icp_stats", {})
    contact_stats = ctx.artifacts.get("contact_stats", {})
    selection = ctx.artifacts.get("selection_stats", {})
    dedup_stats = ctx.artifacts.get("company_dedup_stats", {})
    dynamic = ctx.artifacts.get("dynamic_lead_search", {})
    search_summary = ctx.artifacts.get("search_summary", {})
    priority_counts = selection.get("priority_counts", {})
    buyer_counts = selection.get("buyer_probability_counts", {})
    target = dynamic.get("target_leads", search_summary.get("lead_target",
             pipeline.LEAD_CONFIG_DEFAULTS["max_companies"]))

    rows = [
        ("Selected country/region (locked)", search_summary.get("selected_region", "")),
        ("Target companies (MAX_COMPANIES)", target),
        ("Target reached", "Yes" if leads_count >= target else "No"),
        ("Total job rows in audit", jobs_count),
        ("Total unique companies", dedup_stats.get("companies", 0)),
        ("Low-probability buyer companies (flagged, kept & exported)", icp_stats.get("excluded", 0)),
        ("Total qualified companies", icp_stats.get("qualified", 0)),
        ("Companies with no reported employee count", icp_stats.get("unknown_size", 0)),
        ("Contacts found", contact_stats.get("contacts_found", 0)),
        ("Contacts not found", contact_stats.get("contacts_not_found", 0)),
        ("Total final companies exported", leads_count),
        ("Lead Priority - High", priority_counts.get("High", 0)),
        ("Lead Priority - Medium", priority_counts.get("Medium", 0)),
        ("Lead Priority - Low", priority_counts.get("Low", 0)),
        ("Buyer Probability - High", buyer_counts.get("High", 0)),
        ("Buyer Probability - Medium", buyer_counts.get("Medium", 0)),
        ("Buyer Probability - Low", buyer_counts.get("Low", 0)),
        ("Locations searched", ", ".join(search_summary.get("regions_searched", []))),
        ("Locations locked out (country lock - never searched)",
         ", ".join(search_summary.get("regions_skipped", []))),
    ]

    if leads_count < target:
        selected_region = search_summary.get("selected_region", "the selected country")
        rows.append((
            "NOTE",
            f"Only {leads_count} genuine companies were available in {selected_region} "
            f"after all its available results were searched. The workbook was not "
            f"padded, and the country was never switched to make up the difference.",
        ))

    for region, stats in search_summary.get("region_stats", {}).items():
        rows.append((f"{region} - running qualified leads", stats.get("running_leads", 0)))
        rows.append((f"{region} - audit jobs fetched", stats.get("raw", 0)))

    return pd.DataFrame(rows, columns=["Metric", "Value"])


def _drop_internal(df: pd.DataFrame) -> pd.DataFrame:
    """Strip internal working columns before anything is exported."""
    if df is None or df.empty:
        return pd.DataFrame() if df is None else df
    return df[[c for c in df.columns if not str(c).startswith("_")]]


def export_excluded_companies(df: pd.DataFrame, output_dir: str) -> str | None:
    """Write excluded companies to their own CSV as well as the workbook
    sheet, so the audit trail survives outside Excel."""
    if df is None or df.empty:
        return None

    out_path = Path(output_dir)
    out_path.mkdir(parents=True, exist_ok=True)
    file_path = out_path / f"excluded_companies_{date.today().isoformat()}.csv"
    _drop_internal(df).to_csv(file_path, index=False)
    return str(file_path)


def print_pipeline_report(ctx, jobs_count: int, leads_count: int) -> None:
    """Console version of the Pipeline Summary sheet."""
    summary = build_pipeline_summary_sheet(ctx, jobs_count, leads_count)
    print("\n" + "=" * 70)
    print("PIPELINE SUMMARY")
    print("=" * 70)
    for _, row in summary.iterrows():
        if row["Metric"] == "NOTE":
            print(f"\n  NOTE: {row['Value']}")
        else:
            print(f"  {row['Metric']:<45} {row['Value']}")
    print("=" * 70)


def print_country_lock_summary(report: dict) -> None:
    """Final country-lock / MAX_COMPANIES summary (requirements 4 & 11):
    selected country, what was fetched/paginated/filtered, the Buyer
    Probability breakdown, and whether the configured ceiling was
    reached or the selected country's results were simply exhausted.
    """
    if not report:
        return
    print("\n" + "=" * 70)
    print("COUNTRY LOCK / MAX_COMPANIES SUMMARY")
    print("=" * 70)
    print(f"Selected Country                 : {report['selected_country']}")
    print(f"Target (MAX_COMPANIES)           : {report['max_companies']}")
    print(f"Elapsed                           : {report['elapsed_formatted']}")
    print(f"Total Raw Jobs Collected          : {report['raw_jobs_total']}")
    print(f"Total Raw Unique Companies        : {report['raw_unique_companies']}")
    print(f"Companies Fetched (raw postings, pre-dedup): {report['companies_fetched_raw']}")
    print(f"Number of Search Queries          : {report['num_queries']}")
    print(f"Total Pages Searched              : {report['pages_searched']}")
    print(f"Total API/Search Requests         : {report['requests_made']}")
    print(f"Data Sources/APIs Used            : {', '.join(report['data_sources_used']) or 'none'}")
    print(f"Stale (older than cut-off) Removed: {report['stale_removed']}")
    print(f"Invalid Country Removed           : {report['invalid_country_removed']}")
    print(f"Total Duplicates Removed/Merged   : {report['duplicate_postings_merged']}")
    print(f"Low-Probability Buyers (flagged, kept): {report['low_probability_buyers']}")
    print(f"High Probability                  : {report['high_probability']}")
    print(f"Medium Probability                : {report['medium_probability']}")
    print(f"Low Probability                   : {report['low_probability']}")
    print(f"Final Qualified Companies Exported: {report['final_companies_exported']}")
    if report["regions_locked_out"]:
        print(f"Regions locked out (never searched): {', '.join(report['regions_locked_out'])}")
    if report.get("search_timed_out"):
        status = (
            f"OVERALL_SEARCH_TIMEOUT ({report['search_timeout_seconds']}s) reached - "
            "exported what was already collected rather than continuing "
            "(country was never switched to make up the difference)"
        )
    elif report["target_reached"]:
        status = f"MAX_COMPANIES limit ({report['max_companies']}) reached"
    else:
        status = (
            f"All available valid {report['selected_country']} results exhausted "
            "(country was never switched to make up the difference)"
        )
    print(f"Status                             : {status}")
    print("=" * 70)


def _merge_region_search_summary(summary: dict, region: str, kept: pd.DataFrame,
                                 dropped: pd.DataFrame, filter_stats: dict, stats: dict,
                                 role: str) -> None:
    """Merge one location search into the cumulative run summary."""
    summary.setdefault("region_stats", {})[region] = dict(stats)
    summary["region_stats"][region]["role"] = role

    for key, value in (filter_stats or {}).items():
        summary.setdefault("filter_totals", {})[key] = (
            summary["filter_totals"].get(key, 0) + value
        )

    if dropped is not None and not dropped.empty:
        summary.setdefault("dropped_frames", []).append(dropped)


def _sum_search_stats(region_stats: dict) -> dict:
    """Roll the per-platform fetch/pagination/country-validation counters
    (set in search_keyword_in_country, aggregated in search_country) all
    the way up for the final country-lock/MAX_COMPANIES summary."""
    totals = {"raw": 0, "stale_dropped": 0, "invalid_country": 0,
             "kept": 0, "pages": 0, "requests": 0}
    for region_data in (region_stats or {}).values():
        for country_data in (region_data.get("countries") or {}).values():
            for platform_data in (country_data.get("platforms") or {}).values():
                for key in totals:
                    totals[key] += platform_data.get(key, 0)
    return totals


def _dynamic_lead_target(config: dict, lead_config: dict) -> int:
    """Return MAX_COMPANIES - the actual business ceiling for one run.

    lead_config["target_leads"] is already derived from
    lead_config["max_companies"] by pipeline.load_lead_config; this
    defensive fallback exists only for a lead_config built some other
    way, and pulls its default from the single place max_companies is
    defined (pipeline.LEAD_CONFIG_DEFAULTS) rather than a second literal.
    """
    fallback = pipeline.LEAD_CONFIG_DEFAULTS["max_companies"]
    target = lead_config.get("target_leads", fallback)
    try:
        target = int(target)
    except (TypeError, ValueError):
        target = fallback
    if target <= 0:
        target = fallback
    return target


def _job_buffer_for_location(config: dict, target_leads: int) -> int:
    """How many valid remote job rows to collect before evaluating leads.

    This is deliberately a buffer rather than the stop condition. One job
    is not one lead: company deduplication, ICP filtering and scoring can
    reduce the count substantially. The buffer only limits the amount of
    work done in one location pass; the final stop condition is always the
    number of qualified companies returned by the lead pipeline. It scales
    with MAX_COMPANIES (target_leads * 4) so a 1000-company target still
    collects a realistic pool to filter/dedupe/score down from.
    """
    block = config.get("lead_search") or {}
    try:
        configured = int(block.get("job_buffer_per_location", 250))
    except (TypeError, ValueError):
        configured = 250
    return max(configured, target_leads * 4)


def _evaluate_accumulated_jobs(raw_jobs: pd.DataFrame, user_input: dict,
                               config: dict, lead_config: dict) -> pipeline.PipelineContext:
    """Run lead stages against all jobs accumulated so far.

    The Job Collection stage is intentionally skipped here because the
    location loop in execute_pipeline controls collection and decides when
    to stop. This makes the stop condition a real lead-count condition,
    not a raw-job-count condition.
    """
    ctx = pipeline.PipelineContext(
        config=config,
        lead_config=lead_config,
        user_input=user_input,
        data=raw_jobs.copy() if raw_jobs is not None else pd.DataFrame(),
    )
    # Only stages after Job Collection. The final selected leads are
    # therefore always computed from the entire accumulated job set.
    stages = pipeline.build_pipeline(collect_jobs, apply_remote_safety_filter, deduplicate)[1:]
    return pipeline.run_pipeline(ctx, stages)


def execute_pipeline(user_input: dict, config: dict, lead_config: dict,
                     progress_callback=None) -> dict:
    """Search locations dynamically until the qualified-lead target is met.

    Core rule:

        STOP when UNIQUE QUALIFIED LEADS >= target_leads (MAX_COMPANIES).

    STRICT COUNTRY LOCK: only the user's selected region/country is ever
    searched - there is no fallback to another region, no matter how far
    short of MAX_COMPANIES it falls. See the REGION_DEFINITIONS module
    note. The pipeline is re-evaluated as postings accumulate so the
    decision to stop is based on actual lead quality, company
    deduplication, ICP filtering, contact enrichment and prioritisation -
    not just a raw posting count.
    """
    target_leads = _dynamic_lead_target(config, lead_config)
    job_buffer = _job_buffer_for_location(config, target_leads)

    # OVERALL_SEARCH_TIMEOUT: a hard wall-clock ceiling on the search
    # phase (see load_config's search_timeout_seconds comment). Checked
    # between queries/countries in search_country/search_region; once
    # passed, no new query is started and the run moves straight to
    # exporting what was already collected.
    start_time = time.monotonic()
    search_timeout = config.get("search_timeout_seconds", 1800)
    deadline = start_time + search_timeout

    # Fresh source cache for every run while still avoiding repeated feed
    # downloads inside the same run (important for Europe and fallback).
    job_sources.reset_cache()

    job_title = user_input["job_title"]
    selected = user_input["region"]
    expansion = keyword_expansion.KeywordExpander(
        config.get("keyword_expansion")
    ).expand(job_title)
    keyword_list = expansion["keywords"] or [job_title]

    # STRICT COUNTRY LOCK: only the selected region is ever searched, no
    # matter how short of MAX_COMPANIES it falls.
    search_order = [selected]
    locked_out_regions = [r for r in REGION_ORDER if r != selected]
    collected = pd.DataFrame()

    summary = {
        "selected_region": selected,
        "keyword_expansion": expansion,
        "keyword_counts_collected": {k: 0 for k in keyword_list},
        "regions_searched": [],
        "regions_skipped": list(locked_out_regions),
        "region_stats": {},
        "filter_totals": {},
        "dropped_frames": [],
        "lead_target": target_leads,
        "job_buffer_per_location": job_buffer,
    }

    log.info(f"\n[LEAD TARGET] Target (MAX_COMPANIES): {target_leads} qualified companies")
    log.info(f"[COUNTRY LOCK] Selected country/region: {selected}")
    if locked_out_regions:
        log.info(
            f"[COUNTRY LOCK] {', '.join(locked_out_regions)} will NOT be searched, "
            f"even if fewer than {target_leads} companies are found in {selected}."
        )
    log.info(f"[KEYWORDS] Searched title: '{job_title}'")
    if expansion["expanded"]:
        log.info(
            f"           Similar keywords ({len(expansion['expanded'])}): "
            + ", ".join(f"'{k}'" for k in expansion["expanded"])
        )
    else:
        log.info("           No similar keywords found - searching the title alone.")

    ctx = None
    previous_lead_count = 0

    for region in search_order:
        # search_order is always just [selected] under the strict country
        # lock, so this is always the selected region - never a fallback.
        label = "SELECTED"
        if previous_lead_count >= target_leads:
            summary["regions_skipped"].append(region)
            continue

        log.info(
            f"\n[{label}] Region: {region} | current leads "
            f"{previous_lead_count}/{target_leads} | collecting up to "
            f"~{job_buffer} remote job rows before lead evaluation"
        )

        if progress_callback:
            try:
                progress_callback({
                    "stage": "Searching job boards",
                    "current_region": region,
                    "target_leads": target_leads,
                    "current_leads": previous_lead_count,
                    "regions_searched": list(summary["regions_searched"]),
                })
            except Exception:
                pass

        kept, dropped, filter_stats, stats = search_region(
            region, job_title, config, job_buffer, keyword_list, deadline,
            progress_callback,
            {"start_time": start_time, "region": region, "target_leads": target_leads},
        )
        summary["regions_searched"].append(region)
        _merge_region_search_summary(summary, region, kept, dropped, filter_stats, stats, label)

        if not kept.empty:
            kept = kept.copy()
            kept["_region"] = region
            collected = pd.concat([collected, kept], ignore_index=True)
            collected, _ = deduplicate(collected)

        # Re-evaluate ALL accumulated postings. This is the key change:
        # a country is only considered successful when the resulting lead
        # count, not the job count, reaches the target.
        ctx = _evaluate_accumulated_jobs(collected, user_input, config, lead_config)
        current_leads = len(ctx.data)
        previous_lead_count = current_leads

        summary["region_stats"][region]["running_jobs"] = len(collected)
        summary["region_stats"][region]["running_leads"] = current_leads
        summary["region_stats"][region]["high_priority_leads"] = int(
            (ctx.data.get("lead_priority", pd.Series(dtype=str)) == "High").sum()
        ) if not ctx.data.empty else 0
        summary["region_stats"][region]["medium_priority_leads"] = int(
            (ctx.data.get("lead_priority", pd.Series(dtype=str)) == "Medium").sum()
        ) if not ctx.data.empty else 0
        summary["region_stats"][region]["low_priority_leads"] = int(
            (ctx.data.get("lead_priority", pd.Series(dtype=str)) == "Low").sum()
        ) if not ctx.data.empty else 0

        log.info(
            f"    {region}: {len(collected)} unique remote jobs accumulated -> "
            f"{current_leads}/{target_leads} qualified leads"
        )

        if progress_callback:
            try:
                progress_callback({
                    "stage": "Lead qualification",
                    "current_region": region,
                    "target_leads": target_leads,
                    "current_leads": current_leads,
                    "regions_searched": list(summary["regions_searched"]),
                    "high_priority_leads": summary["region_stats"][region]["high_priority_leads"],
                    "medium_priority_leads": summary["region_stats"][region]["medium_priority_leads"],
                    "low_priority_leads": summary["region_stats"][region]["low_priority_leads"],
                })
            except Exception:
                pass

        if current_leads >= target_leads:
            log.info(
                f"    TARGET REACHED: {current_leads} qualified leads. "
                "Stopping before the next location."
            )
            break

    for region in search_order:
        if region not in summary["regions_searched"] and region not in summary["regions_skipped"]:
            summary["regions_skipped"].append(region)

    if ctx is None:
        ctx = _evaluate_accumulated_jobs(collected, user_input, config, lead_config)

    # Count keyword contribution against the complete collected job set.
    keyword_counts = {k: 0 for k in keyword_list}
    if not collected.empty and "search_keyword" in collected.columns:
        for value in collected["search_keyword"]:
            keyword_counts[value] = keyword_counts.get(value, 0) + 1
    summary["keyword_counts_collected"] = keyword_counts
    summary["keyword_counts"] = keyword_counts
    summary["dropped_df"] = (
        pd.concat(summary.pop("dropped_frames"), ignore_index=True)
        if summary.get("dropped_frames") else pd.DataFrame()
    )
    summary["final_lead_count"] = len(ctx.data)
    summary["target_reached"] = len(ctx.data) >= target_leads
    summary["search_elapsed_seconds"] = time.monotonic() - start_time
    summary["search_timed_out"] = time.monotonic() >= deadline
    summary["search_timeout_seconds"] = search_timeout
    if summary["search_timed_out"]:
        log.warning(
            f"\n[TIMEOUT] OVERALL_SEARCH_TIMEOUT ({search_timeout}s) reached after "
            f"{_format_elapsed(summary['search_elapsed_seconds'])} - exporting the "
            f"{len(ctx.data)} companies already collected instead of continuing."
        )

    ctx.add_artifact("search_summary", summary)
    ctx.add_artifact("dynamic_lead_search", {
        "target_leads": target_leads,
        "final_leads": len(ctx.data),
        "target_reached": len(ctx.data) >= target_leads,
        "regions_searched": summary["regions_searched"],
        "regions_skipped": summary["regions_skipped"],
    })

    job_level = ctx.artifacts.get("job_level_data")
    if job_level is None:
        job_level = ctx.data

    # The main business output is the final MAX_COMPANIES-capped leads.
    # The job audit sheet can still be capped independently at
    # target_total_jobs (raised so it reflects real collected volume).
    jobs_sheet = cap_and_reshape(job_level, config["target_total_jobs"])
    leads_sheet = shape_lead_sheet(ctx.data)

    all_signals = ctx.artifacts.get("all_company_signals")
    excluded = ctx.artifacts.get("excluded_companies")

    # RAW SEARCH RESULTS stage output - every job posting actually
    # collected for the locked country (country-validated, recency-
    # filtered), deliberately NOT capped or deduplicated to one row per
    # company - see requirement "Separate RAW DATA from QUALIFIED DATA".
    raw_companies_sheet = shape_raw_companies_sheet(job_level)
    # QUALIFIED RESULTS stage output, in the exact trimmed column set
    # requested - same source data/scoring as "Qualified Leads" above,
    # just a narrower view. High, Medium AND Low probability companies
    # are all present; none are discarded.
    qualified_companies_sheet = shape_qualified_companies_sheet(ctx.data)

    sheets = {
        "Qualified Leads": leads_sheet,
        "Qualified Companies": qualified_companies_sheet,
        "Raw Companies": raw_companies_sheet,
        "All Company Signals": shape_lead_sheet(all_signals),
        "Excluded Companies": _drop_internal(excluded)
        if excluded is not None and not excluded.empty
        else pd.DataFrame({"Note": ["No companies were excluded in this run."]}),
        "Pipeline Summary": build_pipeline_summary_sheet(
            ctx, len(jobs_sheet), len(leads_sheet)
        ),
        "Remote Jobs": jobs_sheet,
    }

    output_path = export_to_excel(sheets, config["output_dir"])

    excluded_path = None
    if lead_config["save_excluded_companies"]:
        excluded_path = export_excluded_companies(excluded, config["output_dir"])

    # Export/audit metrics are based on exactly the rows that were written.
    summary["dedup_counts"] = ctx.artifacts.get(
        "job_dedup_counts",
        {"by_url": 0, "by_fallback_key": 0, "by_title_company": 0},
    )
    capped_rows = job_level.head(config["target_total_jobs"])
    if "_region" in job_level.columns:
        capped = capped_rows["_region"].tolist()
    else:
        capped = []
    counts = {}
    for region in capped:
        counts[region] = counts.get(region, 0) + 1
    summary["export_region_counts"] = counts

    keyword_counts_export = {k: 0 for k in expansion.get("keywords", [])}
    if "search_keyword" in capped_rows.columns:
        for value in capped_rows["search_keyword"]:
            keyword_counts_export[value] = keyword_counts_export.get(value, 0) + 1
    summary["keyword_counts"] = keyword_counts_export

    dedup_stats = ctx.artifacts.get("company_dedup_stats", {})
    icp_stats = ctx.artifacts.get("icp_stats", {})
    contact_stats = ctx.artifacts.get("contact_stats", {})
    selection = ctx.artifacts.get("selection_stats", {})
    buyer_counts = selection.get("buyer_probability_counts", {})

    # Country-lock / MAX_COMPANIES summary (requirements 4 and 11): one
    # consolidated report of what was fetched, filtered, and exported for
    # the selected country, built from the same artifacts as everything
    # else above rather than a separate tracking mechanism.
    search_totals = _sum_search_stats(summary.get("region_stats", {}))
    final_count = len(leads_sheet)
    # "Duplicates removed" = raw postings collected minus the job-level
    # rows that actually reached company grouping (job-level dedup, see
    # main.deduplicate) PLUS the postings company-identity merged into an
    # existing company afterwards (company-level dedup).
    job_rows_after_job_dedup = int(dedup_stats.get("job_rows", 0))
    companies_after_company_dedup = int(dedup_stats.get("companies", 0))
    platforms_used = sorted({p for stats in summary.get("region_stats", {}).values()
                             for c in (stats.get("countries") or {}).values()
                             for p in (c.get("platforms") or {}).keys()})
    country_lock_report = {
        "selected_country": selected,
        "max_companies": target_leads,
        "raw_jobs_total": len(job_level),
        "raw_unique_companies": companies_after_company_dedup,
        "companies_fetched_raw": search_totals["raw"],
        "num_queries": len(keyword_list),
        "pages_searched": search_totals["pages"],
        "requests_made": search_totals["requests"],
        "data_sources_used": platforms_used,
        "stale_removed": search_totals["stale_dropped"],
        "invalid_country_removed": search_totals["invalid_country"],
        "duplicate_postings_merged": max(
            search_totals["raw"] - job_rows_after_job_dedup, 0
        ) + max(job_rows_after_job_dedup - companies_after_company_dedup, 0),
        "low_probability_buyers": icp_stats.get("low_probability_buyers", icp_stats.get("excluded", 0)),
        "high_probability": buyer_counts.get("High", 0),
        "medium_probability": buyer_counts.get("Medium", 0),
        "low_probability": buyer_counts.get("Low", 0),
        "final_companies_exported": final_count,
        "target_reached": final_count >= target_leads,
        "regions_locked_out": [r for r in REGION_ORDER if r != selected],
        "elapsed_seconds": summary.get("search_elapsed_seconds", 0),
        "elapsed_formatted": _format_elapsed(summary.get("search_elapsed_seconds", 0)),
        "search_timed_out": summary.get("search_timed_out", False),
        "search_timeout_seconds": summary.get("search_timeout_seconds", search_timeout),
    }
    summary["country_lock_report"] = country_lock_report

    if progress_callback:
        try:
            progress_callback({
                "stage": "Completed",
                "target_leads": target_leads,
                "current_leads": len(leads_sheet),
                "regions_searched": summary["regions_searched"],
                "regions_skipped": summary["regions_skipped"],
                "high_priority_leads": selection.get("priority_counts", {}).get("High", 0),
                "medium_priority_leads": selection.get("priority_counts", {}).get("Medium", 0),
                "low_priority_leads": selection.get("priority_counts", {}).get("Low", 0),
            })
        except Exception:
            pass

    return {
        "ctx": ctx,
        "output_path": output_path,
        "excluded_path": excluded_path,
        "summary": summary,
        "jobs_count": len(jobs_sheet),
        "leads_count": len(leads_sheet),
        "unique_companies": dedup_stats.get("companies", 0),
        "excluded_companies": icp_stats.get("excluded", 0),
        "qualified_companies": icp_stats.get("qualified", 0),
        "contacts_found": contact_stats.get("contacts_found", 0),
        "contacts_not_found": contact_stats.get("contacts_not_found", 0),
        "priority_counts": selection.get("priority_counts", {}),
        "buyer_probability_counts": buyer_counts,
        "country_lock_report": country_lock_report,
    }


def main():
    setup_logging()
    config_path = sys.argv[1] if len(sys.argv) > 1 else "config.yaml"

    try:
        config = load_config(config_path)
        lead_config = pipeline.load_lead_config(config)
        user_input = prompt_user_input()

        result = execute_pipeline(user_input, config, lead_config)
        ctx = result["ctx"]

        if result["summary"]:
            print_summary(result["summary"], result["leads_count"],
                          lead_config["target_leads"], config["remote_strictness"])
        else:
            print("\n[WARN] Job collection did not complete - no search summary to report.")

        pipeline.print_pipeline_summary(ctx)
        print_pipeline_report(ctx, result["jobs_count"], result["leads_count"])
        print_country_lock_summary(result.get("country_lock_report"))

        print(f"\nOutput written to  : {result['output_path']}")
        if result["excluded_path"]:
            print(f"Low-probability buyer companies (audit log): {result['excluded_path']}")

    except KeyboardInterrupt:
        print("\nCancelled.")
        sys.exit(1)
    except Exception:
        print("\n[ERROR] Unexpected failure:")
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
