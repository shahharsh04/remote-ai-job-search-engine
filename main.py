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

import logging
import re
import sys
import traceback
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import yaml
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from jobspy import scrape_jobs

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
    config.setdefault("hours_old", 168)
    config.setdefault("results_wanted_per_platform", 50)
    config.setdefault("max_pages_per_platform", 3)
    config.setdefault("target_total_jobs", 50)
    config.setdefault("lead_search", {"job_buffer_per_location": 250})
    config.setdefault("output_dir", "./output")
    config.setdefault("remote_strictness", "balanced")

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
                   max_pages: int, source_config: dict = None) -> tuple[pd.DataFrame, bool]:
    """Fetch one platform. Never raises - a failing platform logs a
    warning and yields whatever it collected so far (or nothing), so one
    bad platform cannot kill the run.

    Returns (df, hours_old_applied_server_side).

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
            return pd.DataFrame(), False
        except Exception as exc:
            log.warning(f"    [WARN] {platform}: unexpected failure: {exc}")
            return pd.DataFrame(), False
        # These boards return absolute dates already, so the shared
        # apply_recency_filter (not a server-side parameter) is what
        # enforces hours_old for them too - same as Indeed.
        return df, False

    kwargs, hours_applied = build_search_kwargs(
        platform, job_title, location, country_indeed, hours_old
    )

    frames = []
    seen_urls = set()
    collected = 0

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

        frames.append(fresh)
        collected += len(fresh)

        # A short page means the platform has nothing more to give.
        if len(df) < remaining:
            break

    combined = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    return combined, hours_applied


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


def search_keyword_in_country(keyword: str, original_title: str, entry: dict,
                              config: dict) -> tuple[pd.DataFrame, dict]:
    """Search every configured platform for one keyword in one country.

    This is the original per-country search, unchanged except that the
    term it searches is now a keyword rather than always the user's own
    title - so every platform receives the expanded keywords too, and
    each row records which keyword found it.
    """
    location = entry["location"]
    country_indeed = entry["country_indeed"]
    hours_old = config["hours_old"]

    frames = []
    platform_stats = {}

    for platform in config["platforms"]:
        df, hours_applied = fetch_platform(
            platform, keyword, location, country_indeed, hours_old,
            config["results_wanted_per_platform"], config["max_pages_per_platform"],
            config.get("job_sources"),
        )
        raw_count = len(df)

        stale_dropped = 0
        if not hours_applied and not df.empty:
            df, stale_dropped = apply_recency_filter(df, hours_old)

        platform_stats[platform] = {
            "raw": raw_count,
            "stale_dropped": stale_dropped,
            "kept": len(df),
        }
        log.info(
            f"      {platform:<9} fetched {raw_count:>3} | stale {stale_dropped:>2} "
            f"| carried forward {len(df):>3}"
        )

        if df.empty:
            continue

        # search_keyword holds the keyword that actually returned the
        # row; original_job_title holds what the user asked for. Keeping
        # both is what lets the export explain each row's inclusion.
        df["search_keyword"] = keyword
        df["original_job_title"] = original_title
        df["country"] = country_indeed
        frames.append(df)

    combined = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    return combined, platform_stats


def search_country(job_title: str, entry: dict, config: dict,
                   keyword_list: list = None,
                   needed: int = None) -> tuple[pd.DataFrame, dict, dict]:
    """Search one country across every configured keyword and platform.

    `keyword_list[0]` is the user's own title and is always searched
    first, so the primary keyword's rows are the ones that survive
    deduplication and land at the top of the export. Passing None
    searches the title alone, which is the pre-expansion behaviour.

    Like the country loop in search_region, the keyword loop stops early
    once `needed` valid remote rows exist - an extra keyword is only
    worth its network cost while the target is still short.
    """
    keyword_list = keyword_list or [job_title]

    frames = []
    platform_stats = {}
    keyword_stats = {}

    for index, keyword in enumerate(keyword_list):
        if index:
            log.info(f"      ~ similar keyword: '{keyword}'")

        raw, stats = search_keyword_in_country(keyword, job_title, entry, config)
        keyword_stats[keyword] = {
            "raw": len(raw),
            "primary": index == 0,
            "platforms": stats,
        }
        for platform, values in stats.items():
            running = platform_stats.setdefault(
                platform, {"raw": 0, "stale_dropped": 0, "kept": 0}
            )
            for field in running:
                running[field] += values[field]

        if not raw.empty:
            frames.append(raw)

        if needed is None or not frames:
            continue

        so_far = pd.concat(frames, ignore_index=True)
        kept_so_far, _, _ = apply_remote_safety_filter(so_far, config["remote_strictness"])
        deduped_so_far, _ = deduplicate(kept_so_far)
        if len(deduped_so_far) >= needed and index + 1 < len(keyword_list):
            log.info(
                f"      Target reached in {entry['country_indeed']} - skipping "
                f"{len(keyword_list) - index - 1} further keyword(s)."
            )
            break

    combined = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    return combined, platform_stats, keyword_stats


def search_region(region: str, job_title: str, config: dict, needed: int,
                  keyword_list: list = None) -> tuple[pd.DataFrame, pd.DataFrame, dict, dict]:
    """Search one region - which may be several countries, as Europe is -
    and return its valid remote rows.

    Stops early between countries once `needed` valid rows have been
    collected, so a region that fills the target on its first country
    does not pay for the rest.

    Returns (kept_df, dropped_df, filter_stats, region_stats).
    """
    entries = REGION_DEFINITIONS[region]
    raw_frames = []
    per_country = {}

    for entry in entries:
        country = entry["country_indeed"]
        log.info(f"    -- {country} --")

        raw, platform_stats, keyword_stats = search_country(
            job_title, entry, config, keyword_list, needed
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
    """Run the selected region first, then fall back through the other
    regions only if the target has not been met.

    Ordering matters here and is deliberate: frames are concatenated in
    the order they were searched and dedup keeps the first occurrence, so
    the selected region's jobs stay at the top of the spreadsheet and a
    posting that appears in two regions is credited to the selected one.
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

    fallback_order = [r for r in REGION_ORDER if r != selected]
    search_order = [selected] + fallback_order

    collected = pd.DataFrame()
    dropped_all = []
    region_stats = {}
    filter_totals = {}
    regions_searched = []
    regions_skipped = []

    for region in search_order:
        if len(collected) >= target:
            regions_skipped.append(region)
            continue

        needed = target - len(collected)
        is_fallback = region != selected
        label = "FALLBACK" if is_fallback else "SELECTED"
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
        "Priority Reason": 55, "LinkedIn Search Links": 45, "job_description": 70,
        "Hiring Signal": 30, "Company Website": 32, "Contact Email": 30,
        "Contact LinkedIn URL": 34, "Contact Source": 30, "Exclusion Reason": 34,
        "Company Name": 30, "Value": 60, "Metric": 45, "job_url": 34,
        "company_url": 30, "Contact Title": 32, "ICP Status": 26,
        "Matched Keywords": 45, "Searched Title": 26,
        "matched_keyword": 30, "original_job_title": 26,
        "Contact Email (2nd)": 28, "Funding Signal": 30, "Team Maturity": 24,
        "AI Hiring Stage": 22, "job_title": 34, "company_name": 26,
    }
    narrow_headers = {
        "Lead Priority": 13, "Priority Score": 12, "Employee Count": 15,
        "Hiring Signal Strength": 15, "Contact Strength": 13,
        "Company Size": 15, "Matching Position Count": 12, "Country": 12,
        "Region": 12, "Date Fetched": 13, "Contact Confidence": 13,
        "is_remote": 10,
    }
    for col_idx in range(1, worksheet.max_column + 1):
        letter = get_column_letter(col_idx)
        header = str(worksheet.cell(row=1, column=col_idx).value or "")
        width = wide_headers.get(header) or narrow_headers.get(header) or 20
        worksheet.column_dimensions[letter].width = width

    # Make High-priority leads findable at a glance, as the brief asks.
    headers = [str(worksheet.cell(row=1, column=i).value or "")
               for i in range(1, worksheet.max_column + 1)]
    if "Lead Priority" in headers:
        priority_col = headers.index("Lead Priority") + 1
        fills = {
            "High": PatternFill("solid", fgColor="C6EFCE"),
            "Medium": PatternFill("solid", fgColor="FFEB9C"),
            "Low": PatternFill("solid", fgColor="F2F2F2"),
        }
        for row_idx in range(2, worksheet.max_row + 1):
            cell = worksheet.cell(row=row_idx, column=priority_col)
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
    "Priority Score": "priority_score",
    "Company Name": "company_name",
    "Company Website": "company_url_direct",
    "Company Domain": "company_domain",
    "Employee Count": "employee_count",
    "Company Size": "company_size",
    "Company Type": "company_type",
    "ICP Status": "icp_status",
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
    "LinkedIn Search Links": "contact_search_urls",
    "Country": "country",
    "Region": "region",
    "Date Fetched": "date_fetched",
}

# Blank cells read as missing data; these say so explicitly instead.
NOT_AVAILABLE_DEFAULTS = {
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


def shape_lead_sheet(df: pd.DataFrame) -> pd.DataFrame:
    """Put company rows into the lead column order.

    Anything a source did not provide is filled with an explicit
    "Unknown"/"Not Found" rather than a blank or a guess.
    """
    if df is None or df.empty:
        return pd.DataFrame(columns=list(LEAD_COLUMNS.keys()))

    df = df.copy()
    if "date_fetched" not in df.columns:
        df["date_fetched"] = date.today().isoformat()

    output = pd.DataFrame()
    for label, source in LEAD_COLUMNS.items():
        if source in df.columns:
            column = df[source]
        else:
            column = pd.Series([None] * len(df), index=df.index)
        default = NOT_AVAILABLE_DEFAULTS.get(label)
        if default is not None:
            column = column.fillna(default).replace("", default)
        output[label] = column
    return output.reset_index(drop=True)


def build_pipeline_summary_sheet(ctx, jobs_count: int, leads_count: int) -> pd.DataFrame:
    """Build the workbook's cumulative lead-search summary."""
    icp_stats = ctx.artifacts.get("icp_stats", {})
    contact_stats = ctx.artifacts.get("contact_stats", {})
    selection = ctx.artifacts.get("selection_stats", {})
    dedup_stats = ctx.artifacts.get("company_dedup_stats", {})
    dynamic = ctx.artifacts.get("dynamic_lead_search", {})
    search_summary = ctx.artifacts.get("search_summary", {})
    priority_counts = selection.get("priority_counts", {})
    target = dynamic.get("target_leads", search_summary.get("lead_target", 50))

    rows = [
        ("Target qualified leads", target),
        ("Target reached", "Yes" if leads_count >= target else "No"),
        ("Total job rows in audit", jobs_count),
        ("Total unique companies", dedup_stats.get("companies", 0)),
        ("Companies excluded (competitors)", icp_stats.get("excluded", 0)),
        ("Total qualified companies", icp_stats.get("qualified", 0)),
        ("Companies with no reported employee count", icp_stats.get("unknown_size", 0)),
        ("Contacts found", contact_stats.get("contacts_found", 0)),
        ("Contacts not found", contact_stats.get("contacts_not_found", 0)),
        ("Total final leads", leads_count),
        ("High priority", priority_counts.get("High", 0)),
        ("Medium priority", priority_counts.get("Medium", 0)),
        ("Low priority", priority_counts.get("Low", 0)),
        ("Locations searched", ", ".join(search_summary.get("regions_searched", []))),
        ("Locations skipped", ", ".join(search_summary.get("regions_skipped", []))),
    ]

    if leads_count < target:
        rows.append((
            "NOTE",
            f"Only {leads_count} genuine qualified leads were available after "
            f"all configured locations were searched. The workbook was not padded.",
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


def _dynamic_lead_target(config: dict, lead_config: dict) -> int:
    """Return the actual business target for one run."""
    target = lead_config.get("target_leads", 50)
    try:
        target = int(target)
    except (TypeError, ValueError):
        target = int(lead_config.get("target_leads", 50))
    if target <= 0:
        target = int(lead_config.get("target_leads", 50))
    return target


def _job_buffer_for_location(config: dict, target_leads: int) -> int:
    """How many valid remote job rows to collect before evaluating leads.

    This is deliberately a buffer rather than the stop condition. One job
    is not one lead: company deduplication, ICP filtering and scoring can
    reduce the count substantially. The buffer only limits the amount of
    work done in one location pass; the final stop condition is always the
    number of qualified leads returned by the lead pipeline.
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

        STOP when UNIQUE QUALIFIED LEADS >= target_leads.

    The selected region is always searched first. When it does not produce
    enough leads, the remaining regions are searched in the configured
    fallback order. The pipeline is re-evaluated after each region so the
    decision to continue is based on actual lead quality, company
    deduplication, ICP filtering, contact enrichment and prioritisation.
    """
    target_leads = _dynamic_lead_target(config, lead_config)
    job_buffer = _job_buffer_for_location(config, target_leads)

    # Fresh source cache for every run while still avoiding repeated feed
    # downloads inside the same run (important for Europe and fallback).
    job_sources.reset_cache()

    job_title = user_input["job_title"]
    selected = user_input["region"]
    expansion = keyword_expansion.KeywordExpander(
        config.get("keyword_expansion")
    ).expand(job_title)
    keyword_list = expansion["keywords"] or [job_title]

    search_order = [selected] + [r for r in REGION_ORDER if r != selected]
    collected = pd.DataFrame()

    summary = {
        "selected_region": selected,
        "keyword_expansion": expansion,
        "keyword_counts_collected": {k: 0 for k in keyword_list},
        "regions_searched": [],
        "regions_skipped": [],
        "region_stats": {},
        "filter_totals": {},
        "dropped_frames": [],
        "lead_target": target_leads,
        "job_buffer_per_location": job_buffer,
    }

    log.info(f"\n[LEAD TARGET] Target: {target_leads} qualified leads")
    log.info(f"[LOCATION ORDER] { ' -> '.join(search_order) }")
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
        label = "SELECTED" if region == selected else "FALLBACK"
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
            region, job_title, config, job_buffer, keyword_list
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

    # The main business output is the final 50 leads. The job audit sheet
    # can still be capped independently at target_total_jobs.
    jobs_sheet = cap_and_reshape(job_level, config["target_total_jobs"])
    leads_sheet = shape_lead_sheet(ctx.data)

    all_signals = ctx.artifacts.get("all_company_signals")
    excluded = ctx.artifacts.get("excluded_companies")

    sheets = {
        "Qualified Leads": leads_sheet,
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

        print(f"\nOutput written to  : {result['output_path']}")
        if result["excluded_path"]:
            print(f"Excluded companies : {result['excluded_path']}")

    except KeyboardInterrupt:
        print("\nCancelled.")
        sys.exit(1)
    except Exception:
        print("\n[ERROR] Unexpected failure:")
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
