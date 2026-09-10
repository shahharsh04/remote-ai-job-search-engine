"""
Dedicated remote-job-board connectors
======================================
JobSpy (main.py's `fetch_platform`) covers Indeed and LinkedIn. This
module adds the dedicated remote-only job boards that JobSpy does not
scrape, each through a source that is actually available:

  Platform            Access method                          Type
  -------------------------------------------------------------------
  remoteok            https://remoteok.com/api                public JSON API
  remotive            https://remotive.com/api/remote-jobs    public JSON API
  weworkremotely      per-category RSS feeds                  public RSS
  jobspresso          https://jobspresso.co/jobs/feed/         public RSS

Every one of these is a source the operator explicitly published for
programmatic reuse - a documented API endpoint or an RSS feed - not a
scrape of pages meant for browsers. Each API's own terms are followed:
RemoteOK and Remotive both require attribution (link back to the
original listing, name the source), which is satisfied automatically
because `job_url` always points at the source's own listing page and
`site` always names the source - see main.OUTPUT_COLUMNS. No page
protected by a login, CAPTCHA or a ToS that forbids automation is
touched by this module.

Platforms in the brief that are deliberately NOT here:

  Wellfound / AngelList Talent  - no public API. Job search is only
    reachable through the site's authenticated internal GraphQL calls,
    and robots.txt disallows the query-string URLs that carry a job
    id/slug. Scraping the rendered page would mean automating a login
    and ignoring robots.txt - both against the terms and against this
    project's rule of never bypassing access controls.
  FlexJobs                      - subscription service. Listings are
    behind a paywall and FlexJobs' terms prohibit automated collection
    or redistribution of paid content. No public API is offered.

Both are left as commented-out entries in config.yaml's `platforms`
list with the reason inline, so a future integration (e.g. a licensed
partner feed) is a one-line config change plus one new function here -
see "Adding a new platform" at the bottom of this file.

Design used by every connector below:

  - One function per platform: `search_<name>(keyword, country_indeed,
    hours_old, results_wanted, settings) -> pd.DataFrame`, registered in
    PLATFORM_REGISTRY under the platform's config key.
  - A DataFrame in the SAME shape JobSpy returns (title, company,
    job_url, location, description, is_remote, date_posted, site,
    job_type, company_url, company_url_direct, min_amount, max_amount,
    currency) so every later pipeline stage - the remote-safety filter,
    deduplication, company identity, ICP, Excel export - runs completely
    unchanged (requirement: keep that logic unchanged).
  - Every network call goes through `_get_json`/`_get_text`, which retry
    with backoff, enforce a timeout, and raise JobSourceError on final
    failure. The connector itself never raises past its own boundary -
    main.fetch_platform wraps each call in try/except too, but the
    intent is that a single bad platform never takes down the run
    (requirement: continue on a platform failure).
  - These boards are NOT partitioned by country the way JobSpy's
    country_indeed parameter is - a search returns the same global feed
    regardless of which of the run's countries is being searched. Two
    things follow: (1) the raw feed is fetched once per run and cached
    (_CACHE) rather than re-fetched for every one of Europe's five
    countries, and (2) a job that names a specific required location
    ("USA Only", "UK/Europe timezones") is checked against the country
    currently being searched with `_location_permits`, so a US-only
    remote job is not offered to a UK search. A job with no stated
    restriction, or a global one ("Worldwide", "Anywhere"), passes for
    every country - matching how JobSpy jobs sourced by country still
    end up compared against the full remote-safety text check.
"""

import logging
import re
import time
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from urllib.parse import quote

import pandas as pd
import requests

log = logging.getLogger("job_search")


class JobSourceError(Exception):
    """A platform could not be reached or returned something unusable.
    Always caught by the caller - see fetch_platform in main.py."""


# ---------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------
DEFAULT_SOURCE_SETTINGS = {
    "timeout_seconds": 12,
    "max_retries": 2,
    "retry_backoff_seconds": 1.5,
    "user_agent": "RemoteJobSearchEngine/1.0 (+job search tool; contact via repository)",
}

# Per-platform overrides/extras layered onto the defaults above.
PLATFORM_DEFAULTS = {
    "remoteok": {"results_wanted": 40},
    "remotive": {"results_wanted": 40},
    "weworkremotely": {
        "results_wanted": 40,
        # Category feeds actually worth checking for tech roles. Add or
        # remove slugs here (or via config.yaml) without touching code -
        # see https://weworkremotely.com/categories for the full list.
        "categories": [
            "remote-programming-jobs",
            "remote-full-stack-programming-jobs",
            "remote-back-end-programming-jobs",
            "remote-front-end-programming-jobs",
            "remote-devops-sysadmin-jobs",
        ],
    },
    "jobspresso": {"results_wanted": 40},
}

# Platforms named in the brief that have no legitimate free/public
# access path (see the module docstring for why). Kept here, rather
# than silently ignored, so adding one to config.yaml's `platforms` by
# mistake gives a clear, specific reason instead of a generic
# "unknown platform" warning.
UNAVAILABLE_PLATFORMS = {
    "wellfound": (
        "no public API; job search requires an authenticated session and "
        "robots.txt disallows the job-detail query strings"
    ),
    "angellist": "same product as Wellfound - see 'wellfound'",
    "flexjobs": (
        "subscription service; listings are paywalled and its terms "
        "prohibit automated collection - no public API is offered"
    ),
}


def load_source_config(config: dict) -> dict:
    """Read the `job_sources` block, filling in every default.

    Structure: {"defaults": {...}, "<platform>": {...per-platform...}}.
    A platform inherits `defaults`, then PLATFORM_DEFAULTS, then its own
    config.yaml entry, in that order - so widening a timeout for every
    platform is one line, and one platform can still be tuned alone.
    """
    supplied = (config or {}).get("job_sources") or {}
    if not isinstance(supplied, dict):
        log.warning("[WARN] job_sources in config is not a mapping - using defaults.")
        supplied = {}

    base = dict(DEFAULT_SOURCE_SETTINGS)
    base.update({k: v for k, v in (supplied.get("defaults") or {}).items()})

    resolved = {}
    for platform in PLATFORM_REGISTRY:
        settings = dict(base)
        settings.update(PLATFORM_DEFAULTS.get(platform, {}))
        settings.update(supplied.get(platform) or {})
        resolved[platform] = settings
    return resolved


# ---------------------------------------------------------------------
# HTTP helpers - shared retry/timeout/logging for every connector
# ---------------------------------------------------------------------
def _request(url: str, settings: dict, params: dict = None) -> requests.Response:
    """GET with a timeout and bounded retries. Raises JobSourceError only
    after every retry is exhausted, so a single hiccup does not drop the
    platform for the whole run."""
    timeout = float(settings.get("timeout_seconds", 12))
    max_retries = int(settings.get("max_retries", 2))
    backoff = float(settings.get("retry_backoff_seconds", 1.5))
    headers = {"User-Agent": settings.get("user_agent", DEFAULT_SOURCE_SETTINGS["user_agent"])}

    last_error = None
    for attempt in range(max_retries + 1):
        try:
            response = requests.get(url, params=params, headers=headers, timeout=timeout)
            if response.status_code == 429:
                # Rate-limited: back off and try again rather than giving
                # up on a platform after one busy response.
                last_error = JobSourceError(f"rate limited (429) on attempt {attempt + 1}")
            elif response.status_code >= 400:
                last_error = JobSourceError(f"HTTP {response.status_code}")
            else:
                return response
        except requests.RequestException as exc:
            last_error = exc

        if attempt < max_retries:
            time.sleep(backoff * (attempt + 1))

    raise JobSourceError(str(last_error))


def _get_json(url: str, settings: dict, params: dict = None):
    response = _request(url, settings, params)
    try:
        return response.json()
    except ValueError as exc:
        raise JobSourceError(f"invalid JSON response: {exc}")


def _get_text(url: str, settings: dict) -> str:
    response = _request(url, settings)
    return response.text


# ---------------------------------------------------------------------
# Location-restriction matching
# ---------------------------------------------------------------------
# Short, structured "required location" strings, as these boards
# actually write them (Remotive's candidate_required_location, WWR's
# <region>, RemoteOK's location). Deliberately NOT applied to full job
# descriptions - a paragraph mentioning "US business hours" is not the
# same claim as a required-location field saying "USA Only", and
# treating prose that loosely would create false rejections.
WORLDWIDE_PATTERNS = [
    "anywhere", "worldwide", "world wide", "global", "international",
    "any location", "any timezone",
]

REGION_LOCATION_SYNONYMS = {
    # "us" is included despite being an English word, because this table
    # is only ever matched against short structured location fields
    # (RemoteOK's `location`, Remotive's `candidate_required_location`,
    # WWR's `<region>`) - never a full description - where "us" reliably
    # means the country, not the pronoun.
    "USA": ["usa", "us", "us only", "u s", "united states", "america", "americas"],
    "UK": ["uk", "united kingdom", "britain", "england", "scotland", "wales"],
    "Australia": ["australia", "anz", "aus only"],
    "Germany": ["germany", "deutschland", "european union", "eu", "europe", "emea"],
    "Netherlands": ["netherlands", "dutch", "holland", "european union", "eu", "europe", "emea"],
    "Ireland": ["ireland", "irish", "european union", "eu", "europe", "emea"],
    "France": ["france", "french", "european union", "eu", "europe", "emea"],
    "Spain": ["spain", "spanish", "european union", "eu", "europe", "emea"],
}


def _tokens_regex(phrases) -> re.Pattern:
    parts = sorted({re.escape(p) for p in phrases}, key=len, reverse=True)
    return re.compile(r"\b(?:" + "|".join(parts) + r")\b", re.IGNORECASE)


_WORLDWIDE_RE = _tokens_regex(WORLDWIDE_PATTERNS)
_REGION_RE = {
    country: _tokens_regex(patterns) for country, patterns in REGION_LOCATION_SYNONYMS.items()
}
_ALL_COUNTRY_PATTERNS_RE = _tokens_regex(
    {p for patterns in REGION_LOCATION_SYNONYMS.values() for p in patterns}
)


def location_permits(required_location: str, country_indeed: str) -> bool:
    """True unless `required_location` names a specific place that is NOT
    `country_indeed` and does not include it.

      - blank/unstructured text -> True. An unknown restriction is not
        evidence of one, matching this project's existing "unknown is
        not evidence" rule (see apply_remote_safety_filter/
        apply_recency_filter in main.py).
      - "Worldwide"/"Anywhere"/etc. -> True.
      - text that names `country_indeed` (or its region, e.g. "Europe"
        for a European country_indeed) -> True.
      - text that names ONLY a different, specific place -> False.
    """
    text = str(required_location or "").strip()
    if not text or text.lower() in ("nan", "none"):
        return True
    if _WORLDWIDE_RE.search(text):
        return True

    own_pattern = _REGION_RE.get(country_indeed)
    if own_pattern and own_pattern.search(text):
        return True

    # Names at least one specific place, and none of them is this
    # country (or its region) -> restricted elsewhere.
    if _ALL_COUNTRY_PATTERNS_RE.search(text):
        return False

    # Some other, unrecognised place name - too ambiguous to reject.
    return True


# ---------------------------------------------------------------------
# Row normalisation
# ---------------------------------------------------------------------
# Every connector builds rows with exactly this column set so the
# DataFrame matches what JobSpy returns and every later stage - remote
# filtering, dedup, company identity, Excel export - needs no changes.
_ROW_COLUMNS = [
    "title", "company", "job_url", "location", "description",
    "is_remote", "date_posted", "site", "job_type",
    "company_url", "company_url_direct",
    "min_amount", "max_amount", "currency",
]


def _make_row(**fields) -> dict:
    row = dict.fromkeys(_ROW_COLUMNS, None)
    row.update(fields)
    return row


def _rows_to_frame(rows: list) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame(columns=_ROW_COLUMNS)
    return pd.DataFrame(rows, columns=_ROW_COLUMNS)


def _title_matches(text: str, keyword: str) -> bool:
    """Loose but not blind: every significant word of the keyword must
    appear in the text AS A WHOLE WORD. Used for the boards below that
    have no server-side search, so this is the only relevance gate
    before a posting is offered to the shared remote/keyword pipeline.

    Word-boundary matching matters here specifically because keywords
    can be as short as "AI" or "ML" - a plain substring check would
    match "ai" inside "airport" or "maintenance", and "ml" inside
    "html". Both were observed live against RemoteOK before this was a
    regex.
    """
    words = [w for w in re.findall(r"[a-z0-9]+", keyword.lower()) if len(w) > 1]
    if not words:
        return False
    haystack = text.lower()
    return all(re.search(rf"\b{re.escape(w)}\b", haystack) for w in words)


# ---------------------------------------------------------------------
# Connectors
# ---------------------------------------------------------------------
_CACHE = {}


def reset_cache() -> None:
    """Clear the per-run feed cache. Call once at the start of each
    pipeline run so a new run sees fresh listings; within a run, the
    cache is what stops a global feed (all four of these boards are not
    country-partitioned) from being re-fetched for every one of a
    multi-country region's countries."""
    _CACHE.clear()


def _cached(key, fetch):
    if key not in _CACHE:
        _CACHE[key] = fetch()
    return _CACHE[key]


def search_remoteok(keyword: str, country_indeed: str, hours_old: int,
                    results_wanted: int, settings: dict) -> pd.DataFrame:
    """RemoteOK's public JSON feed: https://remoteok.com/api

    Documented, keyless, and explicitly offered for reuse (its own first
    response element states the attribution terms, which this project
    satisfies via the exported source_platform/job_url columns). It is a
    fixed feed of the most recent postings - there is no search
    parameter - so it is fetched once per run and filtered here by
    keyword and by location.
    """
    def fetch():
        try:
            data = _get_json("https://remoteok.com/api", settings)
        except JobSourceError as exc:
            log.warning(f"    [WARN] remoteok: fetch failed: {exc}")
            return []
        # First element is the feed's legal/metadata notice, not a job.
        return [item for item in data if isinstance(item, dict) and item.get("id")]

    listings = _cached("remoteok:feed", fetch)

    rows = []
    for item in listings:
        title = str(item.get("position") or "").strip()
        if not title:
            continue
        haystack = " ".join([title, " ".join(item.get("tags") or [])])
        if not _title_matches(haystack, keyword):
            continue
        if not location_permits(item.get("location"), country_indeed):
            continue

        job_url = item.get("url") or (
            f"https://remoteok.com/remote-jobs/{item['id']}" if item.get("id") else None
        )
        date_posted = None
        if item.get("date"):
            date_posted = str(item["date"])[:10]

        rows.append(_make_row(
            title=title,
            company=item.get("company"),
            job_url=job_url,
            location=item.get("location") or "Remote",
            description=item.get("description"),
            date_posted=date_posted,
            site="remoteok",
            min_amount=item.get("salary_min"),
            max_amount=item.get("salary_max"),
        ))
        if len(rows) >= results_wanted:
            break

    log.info(f"    remoteok: {len(listings)} in feed -> {len(rows)} match '{keyword}'")
    return _rows_to_frame(rows)


def search_remotive(keyword: str, country_indeed: str, hours_old: int,
                    results_wanted: int, settings: dict) -> pd.DataFrame:
    """Remotive's public JSON API: https://remotive.com/api/remote-jobs

    Documented and keyless, with a real `search` parameter, so the
    keyword is sent server-side rather than filtered after the fact.
    Attribution (link back, name Remotive) is satisfied the same way as
    RemoteOK - see the module docstring.
    """
    def fetch():
        try:
            data = _get_json(
                "https://remotive.com/api/remote-jobs",
                settings,
                params={"search": keyword, "limit": max(results_wanted, 40)},
            )
        except JobSourceError as exc:
            log.warning(f"    [WARN] remotive: fetch failed for '{keyword}': {exc}")
            return []
        return data.get("jobs") or []

    listings = _cached(f"remotive:{keyword.lower()}", fetch)

    rows = []
    for item in listings:
        title = str(item.get("title") or "").strip()
        if not title:
            continue
        # Remotive's own `search` matches the full description, not just
        # the title - a search for "AI Engineer" comes back with plain
        # service-desk and copywriting roles that merely mention "AI"
        # somewhere in the body. Re-checking the title locally is what
        # keeps this platform from reintroducing the low-precision
        # results this change was meant to get rid of.
        if not _title_matches(title, keyword):
            continue
        if not location_permits(item.get("candidate_required_location"), country_indeed):
            continue

        rows.append(_make_row(
            title=title,
            company=item.get("company_name"),
            job_url=item.get("url"),
            location=item.get("candidate_required_location") or "Remote",
            description=item.get("description"),
            date_posted=(item.get("publication_date") or "")[:10] or None,
            site="remotive",
            job_type=item.get("job_type"),
        ))
        if len(rows) >= results_wanted:
            break

    log.info(f"    remotive: {len(listings)} returned -> {len(rows)} kept for '{keyword}'")
    return _rows_to_frame(rows)


# WWR's RSS is plain RSS 2.0 with two extra unnamespaced tags this
# project cares about: <region> (the required-location statement) and
# the standard <link>/<title>/<pubDate>/<description>.
def _parse_rss_items(xml_text: str) -> list:
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError as exc:
        raise JobSourceError(f"malformed RSS: {exc}")

    items = []
    for item in root.iter("item"):
        def text_of(tag):
            el = item.find(tag)
            return el.text.strip() if el is not None and el.text else ""

        items.append({
            "title": text_of("title"),
            "link": text_of("link") or text_of("guid"),
            "region": text_of("region"),
            "pubDate": text_of("pubDate"),
            "description": text_of("description"),
        })
    return items


def _rfc822_to_iso_date(value: str):
    if not value:
        return None
    try:
        return datetime.strptime(value[:25].strip(), "%a, %d %b %Y %H:%M:%S").date().isoformat()
    except ValueError:
        try:
            return datetime.strptime(value.split(" +")[0].strip(), "%a, %d %b %Y %H:%M:%S").date().isoformat()
        except ValueError:
            return None


def search_weworkremotely(keyword: str, country_indeed: str, hours_old: int,
                          results_wanted: int, settings: dict) -> pd.DataFrame:
    """We Work Remotely's public per-category RSS feeds.

    robots.txt allows crawling site-wide except account/admin paths, and
    these feeds are the site's own published syndication format (listed
    on weworkremotely.com/categories) - not a scrape of the browsing UI.
    There is no keyword search built into RSS, so the configured
    categories are fetched once per run and filtered here.
    """
    categories = settings.get("categories") or PLATFORM_DEFAULTS["weworkremotely"]["categories"]

    def fetch():
        all_items = []
        for category in categories:
            url = f"https://weworkremotely.com/categories/{category}.rss"
            try:
                xml_text = _get_text(url, settings)
                all_items.extend(_parse_rss_items(xml_text))
            except JobSourceError as exc:
                log.warning(f"    [WARN] weworkremotely: '{category}' feed failed: {exc}")
        return all_items

    listings = _cached("weworkremotely:feed", fetch)

    rows = []
    seen_links = set()
    for item in listings:
        raw_title = item["title"]
        if not raw_title or item["link"] in seen_links:
            continue

        # WWR titles are conventionally "Company: Job Title". Splitting
        # this out matters for more than a clean company_name column -
        # matching the keyword against the RAW title would let a company
        # whose name happens to contain a keyword word (e.g.
        # "Collaboration.Ai: Senior Software Engineer") false-positive a
        # search for "AI Engineer" even though the role itself is not
        # AI-specific. Only the job-title half is checked for relevance.
        if ":" in raw_title:
            company, _, job_title = raw_title.partition(":")
            company, job_title = company.strip(), job_title.strip()
        else:
            company, job_title = None, raw_title

        if not _title_matches(job_title, keyword):
            continue
        if not location_permits(item["region"], country_indeed):
            continue

        seen_links.add(item["link"])
        rows.append(_make_row(
            title=job_title,
            company=company,
            job_url=item["link"] or None,
            location=item["region"] or "Remote",
            description=item["description"],
            date_posted=_rfc822_to_iso_date(item["pubDate"]),
            site="weworkremotely",
        ))
        if len(rows) >= results_wanted:
            break

    log.info(f"    weworkremotely: {len(listings)} across {len(categories)} feed(s) -> {len(rows)} match '{keyword}'")
    return _rows_to_frame(rows)


def search_jobspresso(keyword: str, country_indeed: str, hours_old: int,
                      results_wanted: int, settings: dict) -> pd.DataFrame:
    """Jobspresso's public jobs RSS: https://jobspresso.co/jobs/feed/

    robots.txt disallows only query-string URLs (`Disallow: /*?`), so
    the plain feed URL is fetched with no query parameters and filtered
    here; Jobspresso is a curated 100%-remote board with no per-listing
    structured location field, so `_location_permits` is not applied -
    an unstated restriction is not evidence of one, same as elsewhere in
    this project.
    """
    def fetch():
        try:
            xml_text = _get_text("https://jobspresso.co/jobs/feed/", settings)
            return _parse_rss_items(xml_text)
        except JobSourceError as exc:
            log.warning(f"    [WARN] jobspresso: feed failed: {exc}")
            return []

    listings = _cached("jobspresso:feed", fetch)

    rows = []
    for item in listings:
        title = item["title"]
        if not title:
            continue
        if not _title_matches(title, keyword):
            continue

        rows.append(_make_row(
            title=title,
            company=None,  # not separable from the title's free text
            job_url=item["link"] or None,
            location="Remote",
            description=item["description"],
            date_posted=_rfc822_to_iso_date(item["pubDate"]),
            site="jobspresso",
        ))
        if len(rows) >= results_wanted:
            break

    log.info(f"    jobspresso: {len(listings)} in feed -> {len(rows)} match '{keyword}'")
    return _rows_to_frame(rows)


# ---------------------------------------------------------------------
# Registry - main.fetch_platform dispatches through this
# ---------------------------------------------------------------------
# Adding a new platform: write a search_<name>(keyword, country_indeed,
# hours_old, results_wanted, settings) -> pd.DataFrame function above
# returning rows built with _make_row, add it here, add its key to
# config.yaml's `platforms` list, and optionally give it defaults in
# PLATFORM_DEFAULTS. Nothing in main.py or pipeline.py needs to change.
PLATFORM_REGISTRY = {
    "remoteok": search_remoteok,
    "remotive": search_remotive,
    "weworkremotely": search_weworkremotely,
    "jobspresso": search_jobspresso,
}


def search(platform: str, keyword: str, country_indeed: str, hours_old: int,
          results_wanted: int, source_config: dict) -> pd.DataFrame:
    """Dispatch to the named platform's connector. Raises JobSourceError
    for a platform this module does not implement - callers (main.py)
    catch it and log a clear reason, then continue with the rest of the
    run."""
    handler = PLATFORM_REGISTRY.get(platform)
    if handler is None:
        if platform in UNAVAILABLE_PLATFORMS:
            raise JobSourceError(
                f"'{platform}' is not integrated - {UNAVAILABLE_PLATFORMS[platform]}"
            )
        raise JobSourceError(f"'{platform}' is not a recognised job source")

    settings = source_config.get(platform, dict(DEFAULT_SOURCE_SETTINGS))
    return handler(keyword, country_indeed, hours_old, results_wanted, settings)
