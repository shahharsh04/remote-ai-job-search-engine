"""
Lead-generation pipeline framework
==================================
The job search produces remote job postings. Turning those into a
CEO-ready lead list is a separate concern, and this module holds the
structure for it.

This file defines:

  * the ordered list of pipeline stages,
  * a Stage base class every stage implements,
  * a context object carrying the working data between stages,
  * a runner that executes stages in order and records what each did,
  * the lead-pipeline configuration block, with validation.

The stages themselves delegate their domain logic to focused modules -
company_identity, icp, contacts and lead_signals - so this file stays a
description of the pipeline rather than a container for its rules.

Every stage reports what it did in rows-in/rows-out terms, and no stage
fabricates data: anything the job boards did not supply is carried
through as "Unknown" or "Not Found".
"""

import logging
from abc import ABC, abstractmethod
from dataclasses import dataclass, field

import pandas as pd

import company_identity
import contacts
import icp
import lead_signals
import excluded_company_store

log = logging.getLogger("job_search")


# ---------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------

# Every tunable the lead pipeline uses, with its default. Business logic
# must read these from the config rather than embedding a number, so a
# non-engineer can retune the pipeline by editing config.yaml alone.
LEAD_CONFIG_DEFAULTS = {
    # Company-size bands, in employees.
    "preferred_min_employees": 20,
    "preferred_max_employees": 500,
    "medium_priority_employee_threshold": 500,
    "low_priority_employee_threshold": 5000,

    # Target size of the final qualified-lead list. The search can
    # expand to additional locations until this number is reached.
    "target_leads": 50,
    "max_final_leads": 50,
    "min_final_leads": 50,

    # Keep companies removed by ICP filtering in their own file.
    "save_excluded_companies": True,
    "excluded_company_dictionary_path": "data/excluded_companies.json",

    # Company types whose primary business competes with an AI
    # engineering/services offering. Remove an entry to stop excluding
    # that type. Large organisations are NOT listed here - they stay in
    # the dataset and are scored lower instead.
    "excluded_company_types": [
        "staffing_agency",
        "recruitment_agency",
        "it_services",
        "ai_ml_consulting",
        "software_outsourcing",
        "bpo_services",
    ],

    "contact_enrichment": {
        # Fetch the company's own public careers page over plain HTTP.
        "enable_careers_page_lookup": True,
        "respect_robots_txt": True,
        "request_timeout_seconds": 8,
        # Enrichment makes network calls, so it is budgeted.
        "max_companies_to_enrich": 20,
        # LinkedIn *search* links, not profiles - see contacts.py.
        "generate_linkedin_search_urls": True,
        "user_agent": "RemoteJobSearchEngine/1.0 (lead research)",
    },

    "scoring": {
        # Points available per factor. Raise a weight to make that factor
        # matter more; the total is normalised to 0-100 either way.
        "weights": {
            "icp_fit": 30,
            "hiring_signal": 25,
            "contact_found": 15,
            "multiple_openings": 10,
            "funding_signal": 10,
            "new_ai_initiative": 10,
        },
        # Score needed for each priority band.
        "priority_bands": {
            "high_min_score": 70,
            "medium_min_score": 45,
        },
    },
}

# Nested blocks are merged key-by-key so a config that overrides one
# weight keeps the defaults for the rest.
NESTED_SECTIONS = ("contact_enrichment", "scoring")


def _merge_section(defaults: dict, supplied, section_name: str, problems: list) -> dict:
    if not isinstance(supplied, dict):
        problems.append(f"{section_name} is not a mapping")
        return dict(defaults)

    merged = dict(defaults)
    for key, value in supplied.items():
        if key not in defaults:
            problems.append(f"unknown {section_name} setting '{key}' - ignored")
            continue
        if isinstance(defaults[key], dict):
            merged[key] = _merge_section(defaults[key], value, f"{section_name}.{key}", problems)
        else:
            merged[key] = value
    return merged


def load_lead_config(config: dict) -> dict:
    """Pull the `lead_pipeline` block out of the main config, fill in
    defaults, and sanity-check the values.

    Bad numbers are reported and corrected rather than allowed to reach
    the stages, because a silently inverted min/max would quietly filter
    out every lead and look like "no results found".
    """
    raw = config.get("lead_pipeline") or {}
    if not isinstance(raw, dict):
        log.warning("[WARN] 'lead_pipeline' in config.yaml is not a mapping - using defaults.")
        raw = {}

    problems = []
    lead_config = dict(LEAD_CONFIG_DEFAULTS)
    for section in NESTED_SECTIONS:
        lead_config[section] = dict(LEAD_CONFIG_DEFAULTS[section])

    for key, value in raw.items():
        if key not in LEAD_CONFIG_DEFAULTS:
            log.warning(f"[WARN] Unknown lead_pipeline setting '{key}' - ignoring it.")
            continue
        if key in NESTED_SECTIONS:
            lead_config[key] = _merge_section(
                LEAD_CONFIG_DEFAULTS[key], value, key, problems
            )
        else:
            lead_config[key] = value

    for key in (
        "preferred_min_employees", "preferred_max_employees",
        "medium_priority_employee_threshold", "low_priority_employee_threshold",
        "target_leads", "max_final_leads", "min_final_leads",
    ):
        value = lead_config[key]
        if not isinstance(value, int) or isinstance(value, bool) or value < 0:
            problems.append(f"{key}={value!r} is not a non-negative whole number")
            lead_config[key] = LEAD_CONFIG_DEFAULTS[key]

    if not isinstance(lead_config["save_excluded_companies"], bool):
        problems.append(
            f"save_excluded_companies={lead_config['save_excluded_companies']!r} is not true/false"
        )
        lead_config["save_excluded_companies"] = LEAD_CONFIG_DEFAULTS["save_excluded_companies"]

    if not isinstance(lead_config["excluded_company_dictionary_path"], str) or not lead_config["excluded_company_dictionary_path"].strip():
        problems.append("excluded_company_dictionary_path must be a non-empty string")
        lead_config["excluded_company_dictionary_path"] = LEAD_CONFIG_DEFAULTS["excluded_company_dictionary_path"]

    if not isinstance(lead_config["excluded_company_types"], list):
        problems.append("excluded_company_types is not a list")
        lead_config["excluded_company_types"] = list(
            LEAD_CONFIG_DEFAULTS["excluded_company_types"]
        )

    if lead_config["preferred_min_employees"] > lead_config["preferred_max_employees"]:
        problems.append(
            f"preferred_min_employees ({lead_config['preferred_min_employees']}) is above "
            f"preferred_max_employees ({lead_config['preferred_max_employees']})"
        )
        lead_config["preferred_min_employees"] = LEAD_CONFIG_DEFAULTS["preferred_min_employees"]
        lead_config["preferred_max_employees"] = LEAD_CONFIG_DEFAULTS["preferred_max_employees"]

    if lead_config["medium_priority_employee_threshold"] > lead_config["low_priority_employee_threshold"]:
        problems.append(
            "medium_priority_employee_threshold is above low_priority_employee_threshold "
            "(medium should be the smaller company size)"
        )
        lead_config["medium_priority_employee_threshold"] = \
            LEAD_CONFIG_DEFAULTS["medium_priority_employee_threshold"]
        lead_config["low_priority_employee_threshold"] = \
            LEAD_CONFIG_DEFAULTS["low_priority_employee_threshold"]

    # `target_leads` is the business target. Keep the legacy max/min
    # settings synchronized so older callers remain compatible.
    if lead_config["target_leads"] <= 0:
        problems.append("target_leads must be greater than zero")
        lead_config["target_leads"] = LEAD_CONFIG_DEFAULTS["target_leads"]

    lead_config["max_final_leads"] = lead_config["target_leads"]
    lead_config["min_final_leads"] = lead_config["target_leads"]

    bands = lead_config["scoring"]["priority_bands"]
    if bands.get("medium_min_score", 0) > bands.get("high_min_score", 0):
        problems.append("medium_min_score is above high_min_score")
        lead_config["scoring"]["priority_bands"] = dict(
            LEAD_CONFIG_DEFAULTS["scoring"]["priority_bands"]
        )

    for problem in problems:
        log.warning(f"[WARN] lead_pipeline: {problem} - reverted to the default.")

    return lead_config


# ---------------------------------------------------------------------
# Context
# ---------------------------------------------------------------------

@dataclass
class PipelineContext:
    """Carries everything a stage needs, and everything stages produce.

    `data` is the working set: each stage receives it and returns the
    version it wants the next stage to see. Side outputs (rows dropped
    for review, companies excluded by ICP rules) go into `artifacts` so
    they can be exported without being mixed into the main result.
    """
    config: dict
    lead_config: dict
    user_input: dict = field(default_factory=dict)
    data: pd.DataFrame = field(default_factory=pd.DataFrame)
    artifacts: dict = field(default_factory=dict)
    stage_records: list = field(default_factory=list)

    def add_artifact(self, name: str, value) -> None:
        self.artifacts[name] = value

    def append_excluded(self, df: pd.DataFrame, stage: str, reason: str) -> None:
        """Record rows a stage removed, tagged with which stage removed
        them and why. Written out later only if save_excluded_companies
        is on."""
        if df is None or df.empty:
            return
        tagged = df.copy()
        tagged["excluded_by_stage"] = stage
        tagged["excluded_reason"] = reason
        existing = self.artifacts.get("excluded_companies")
        self.artifacts["excluded_companies"] = (
            pd.concat([existing, tagged], ignore_index=True)
            if existing is not None and not existing.empty else tagged
        )


# ---------------------------------------------------------------------
# Stage base class
# ---------------------------------------------------------------------

class Stage(ABC):
    """One step of the pipeline.

    Subclasses set `name`/`description` and implement `_run`. Setting
    `implemented = False` marks a stage as declared-but-pending: the
    runner passes the data straight through and reports it as such,
    which keeps the pipeline shape visible before the work is done.
    """

    name: str = "unnamed"
    description: str = ""
    implemented: bool = True

    @abstractmethod
    def _run(self, ctx: PipelineContext) -> pd.DataFrame:
        """Return the working set for the next stage."""

    def run(self, ctx: PipelineContext) -> pd.DataFrame:
        if not self.implemented:
            return ctx.data
        return self._run(ctx)


class PendingStage(Stage):
    """A stage that is part of the design but not built yet.

    It deliberately does nothing. Producing placeholder company sizes or
    contact details here would make the export look complete while being
    fabricated, which is worse than an obviously empty stage.
    """

    implemented = False

    def _run(self, ctx: PipelineContext) -> pd.DataFrame:  # pragma: no cover
        return ctx.data


# ---------------------------------------------------------------------
# The stages
# ---------------------------------------------------------------------

class JobCollectionStage(Stage):
    """Search the job boards. Delegates to the existing region/fallback
    search, so the current job-search behaviour is unchanged."""

    name = "Job Collection"
    description = "Search configured job platforms across the selected region, with dynamic fallback"

    def __init__(self, collect_fn, dedupe_fn):
        self.collect_fn = collect_fn
        self.dedupe_fn = dedupe_fn

    def _run(self, ctx: PipelineContext) -> pd.DataFrame:
        collected, search_summary = self.collect_fn(ctx.user_input, ctx.config)
        ctx.add_artifact("search_summary", search_summary)

        # Final job-level dedup and cap happen here so every later stage
        # sees exactly the job set that will be exported - otherwise the
        # company rows could reference postings the Jobs sheet omits.
        deduped, dedup_counts = self.dedupe_fn(collected)
        ctx.add_artifact("job_dedup_counts", dedup_counts)

        target = ctx.config["target_total_jobs"]
        if len(deduped) > target:
            log.info(f"      trimming {len(deduped)} job rows to the target of {target}")
            deduped = deduped.head(target)
        return deduped


class RemoteFilteringStage(Stage):
    """Re-assert remote-only on the collected set.

    Collection already filters as it goes - it has to, because the
    region fallback decides when to stop based on how many VALID remote
    jobs it has. This stage re-checks the combined result so remote-only
    is enforced at a single, visible point in the pipeline. It should
    normally remove nothing; if it ever does, something upstream changed
    and the count will say so.
    """

    name = "Remote Filtering"
    description = "Verify every collected row is genuinely remote"

    def __init__(self, filter_fn):
        self.filter_fn = filter_fn

    def _run(self, ctx: PipelineContext) -> pd.DataFrame:
        if ctx.data.empty:
            return ctx.data
        kept, dropped, _ = self.filter_fn(ctx.data, ctx.config["remote_strictness"])
        if not dropped.empty:
            log.warning(
                f"    [NOTE] Remote re-check removed {len(dropped)} row(s) that collection "
                "had kept - worth investigating."
            )
            ctx.append_excluded(dropped, self.name, "failed remote re-check")
        return kept


class CompanyNormalizationStage(Stage):
    """Attach a stable company identity to every job row.

    Adds three internal columns and changes no data: `_company_domain`
    (the company's own website, when the board returned one),
    `_company_profile` (the board's company page) and `_company_name_key`
    (the normalised name). The next stage groups on these.

    Rows the boards gave no usable company information for are kept and
    counted, not dropped - a posting with a missing company name is a
    data-quality problem to surface, not a job to discard.
    """

    name = "Company Normalization"
    description = "Derive a stable company identity from name, website and profile URL"

    def _run(self, ctx: PipelineContext) -> pd.DataFrame:
        if ctx.data.empty:
            return ctx.data

        df = ctx.data.copy()
        df["_company_domain"] = df.get(
            "company_url_direct", pd.Series([None] * len(df), index=df.index)
        ).apply(company_identity.extract_company_domain)
        df["_company_profile"] = df.get(
            "company_url", pd.Series([None] * len(df), index=df.index)
        ).apply(company_identity.extract_profile_key)
        df["_company_name_key"] = df.get(
            "company", pd.Series([None] * len(df), index=df.index)
        ).apply(company_identity.normalize_company_name)

        with_domain = int((df["_company_domain"] != "").sum())
        with_profile = int((df["_company_profile"] != "").sum())
        unidentified = int(
            ((df["_company_domain"] == "") &
             (df["_company_profile"] == "") &
             (df["_company_name_key"] == "")).sum()
        )

        log.info(
            f"      identity: {with_domain}/{len(df)} rows have a company website, "
            f"{with_profile} have a board profile"
        )
        if unidentified:
            log.warning(
                f"      [NOTE] {unidentified} row(s) carry no usable company identifier "
                "and will each stay a separate lead."
            )

        ctx.add_artifact("normalization_stats", {
            "rows": len(df),
            "with_domain": with_domain,
            "with_profile": with_profile,
            "unidentified": unidentified,
        })
        return df


# Job-level columns worth carrying onto the company row, and how to
# combine them across that company's postings.
COMPANY_FIRST_VALUE_COLUMNS = [
    "company_url", "company_url_direct", "company_industry",
    "company_num_employees", "company_revenue", "company_addresses",
    "company_description", "country", "_region",
    # The title the user searched for. Identical on every row of a run;
    # carried through so the lead sheet can state what was searched.
    "original_job_title",
]

COMPANY_JOINED_COLUMNS = {
    "matching_job_titles": "title",
    "matching_job_urls": "job_url",
    "locations": "location",
    "source_platforms": "site",
    "search_keywords": "search_keyword",
    "emails_found": "emails",
}

# Advert text is concatenated per company so the later stages can read
# real hiring/funding language out of it. Capped because descriptions are
# long and only the wording matters, not the volume. Internal only - it
# never reaches the spreadsheet.
DESCRIPTION_BLOB_LIMIT = 60000

JOINER = " | "


def _first_real_value(series):
    """First non-empty value, or None. Company fields are sparse, so the
    first row's blank should not hide a later row's real value."""
    for value in series:
        if value is None:
            continue
        text = str(value).strip()
        if text and text.lower() not in ("nan", "none"):
            return value
    return None


def _unique_joined(series) -> str:
    """De-duplicated, order-preserving join. Order matters: the first
    title listed is the first one found, which follows the selected
    region's ordering."""
    seen = []
    for value in series:
        if value is None:
            continue
        text = str(value).strip()
        if not text or text.lower() in ("nan", "none"):
            continue
        if text not in seen:
            seen.append(text)
    return JOINER.join(seen)


class CompanyDeduplicationStage(Stage):
    """Collapse many job postings into one row per company.

    Grouping is by connected components over the exact identifiers from
    the previous stage: rows sharing a website domain, a board profile
    URL, or a normalised name become one company. Similar-but-different
    names never merge, because nothing here compares names for
    similarity.

    The job-level rows are preserved in full as the `job_level_data`
    artifact so the company rows can be audited back to their source.
    """

    name = "Company-Level Deduplication"
    description = "One row per company, with all matching jobs combined"

    def _run(self, ctx: PipelineContext) -> pd.DataFrame:
        # Keep the job-level view for audit and for the Jobs sheet before
        # anything is collapsed.
        ctx.add_artifact("job_level_data", ctx.data.copy())

        if ctx.data.empty:
            return ctx.data

        df = ctx.data.reset_index(drop=True)
        df["_company_group"] = company_identity.assign_company_groups(
            df.to_dict("records")
        )

        companies = []
        # groupby(sort=False) keeps first-appearance order, which is what
        # preserves the selected-region-first ordering into the lead list.
        for _, rows in df.groupby("_company_group", sort=False):
            companies.append(self._build_company_row(rows))

        result = pd.DataFrame(companies)
        multi = int((result["matching_position_count"] > 1).sum()) if not result.empty else 0
        log.info(
            f"      {len(df)} job rows -> {len(result)} companies "
            f"({multi} with more than one opening)"
        )

        ctx.add_artifact("company_dedup_stats", {
            "job_rows": len(df),
            "companies": len(result),
            "companies_with_multiple_jobs": multi,
        })
        return result

    def _build_company_row(self, rows: pd.DataFrame) -> dict:
        company = {}

        # Display name: the most frequently seen spelling, with the
        # longest as tie-break, so "IBM" does not beat "IBM Corporation"
        # by accident of row order.
        names = [
            str(n).strip() for n in rows.get("company", pd.Series(dtype=str))
            if n is not None and str(n).strip().lower() not in ("", "nan", "none")
        ]
        if names:
            counts = {}
            for name in names:
                counts[name] = counts.get(name, 0) + 1
            company["company_name"] = sorted(
                counts, key=lambda n: (counts[n], len(n)), reverse=True
            )[0]
        else:
            company["company_name"] = None

        company["company_domain"] = _first_real_value(
            rows.get("_company_domain", pd.Series(dtype=str))
        )

        for column in COMPANY_FIRST_VALUE_COLUMNS:
            if column in rows.columns:
                company[column.lstrip("_")] = _first_real_value(rows[column])

        company["matching_position_count"] = len(rows)
        for out_column, source in COMPANY_JOINED_COLUMNS.items():
            if source in rows.columns:
                company[out_column] = _unique_joined(rows[source])
            else:
                company[out_column] = ""

        if "description" in rows.columns:
            blob = "\n".join(
                str(d) for d in rows["description"]
                if d is not None and str(d).strip().lower() not in ("", "nan", "none")
            )
            company["_description_blob"] = blob[:DESCRIPTION_BLOB_LIMIT]
        else:
            company["_description_blob"] = ""

        if "date_posted" in rows.columns:
            dates = pd.to_datetime(rows["date_posted"], errors="coerce").dropna()
            company["earliest_job_posted"] = dates.min().date().isoformat() if len(dates) else None
            company["latest_job_posted"] = dates.max().date().isoformat() if len(dates) else None
        else:
            company["earliest_job_posted"] = None
            company["latest_job_posted"] = None

        return company


class IcpFilteringStage(Stage):
    """Classify each company and drop only direct competitors.

    Large organisations are NOT excluded - they stay in the dataset and
    are scored down later. Companies whose size or type could not be
    established are also kept, marked Unknown: absent data is a gap in
    what the boards returned, not a reason to discard a prospect.
    """

    name = "ICP Filtering"
    description = "Classify company type and size; exclude competitors only"

    def _run(self, ctx: PipelineContext) -> pd.DataFrame:
        if ctx.data.empty:
            return ctx.data

        df = ctx.data.copy()
        dictionary_path = ctx.lead_config.get("excluded_company_dictionary_path", "data/excluded_companies.json")
        known_companies = excluded_company_store.all_companies(dictionary_path)
        known_keys = set()
        for item in known_companies:
            for field in ("company_domain", "company_name"):
                value = str(item.get(field) or "").strip().lower()
                if value:
                    known_keys.add(f"{field}:{value}")

        verdicts = []
        persistent_matches = []
        for row in df.to_dict("records"):
            identity_record = {
                "company_name": row.get("company_name") or row.get("company"),
                "company_domain": row.get("_company_domain") or row.get("company_domain"),
            }
            identity_keys = set()
            for field in ("company_domain", "company_name"):
                value = str(identity_record.get(field) or "").strip().lower()
                if value:
                    identity_keys.add(f"{field}:{value}")

            if identity_keys.intersection(known_keys):
                match = next((x for x in known_companies
                              if identity_keys.intersection({
                                  f"company_domain:{str(x.get('company_domain') or '').strip().lower()}",
                                  f"company_name:{str(x.get('company_name') or '').strip().lower()}"
                              })), None)
                known_category = (match or {}).get("category", "Known excluded company")
                known_reason = (match or {}).get("reason", "Previously identified as a low-probability buyer.")
                verdicts.append({
                    "employee_count": icp.parse_employee_count(row.get("company_num_employees"))[0],
                    "employee_count_estimate": icp.parse_employee_count(row.get("company_num_employees"))[1],
                    "company_size": icp.size_band(icp.parse_employee_count(row.get("company_num_employees"))[1], ctx.lead_config),
                    "company_type": known_category,
                    "company_type_evidence": "persistent exclusion dictionary",
                    "icp_status": "Excluded",
                    "exclusion_reason": known_reason,
                })
                persistent_matches.append({
                    "company_name": identity_record["company_name"] or "Unknown company",
                    "company_domain": identity_record["company_domain"] or "",
                    "company_type": known_category,
                    "exclusion_reason": known_reason,
                })
            else:
                verdicts.append(icp.evaluate_company(row, ctx.lead_config))
        for field in ("employee_count", "employee_count_estimate", "company_size",
                      "company_type", "company_type_evidence", "icp_status",
                      "exclusion_reason"):
            df[field] = [v.get(field) for v in verdicts]

        excluded_mask = df["icp_status"] == "Excluded"
        excluded = df[excluded_mask]
        kept = df[~excluded_mask]

        # Persist every excluded company across runs. This turns the dictionary
        # into a cumulative memory: Run 2 can immediately reject a company
        # already identified in Run 1.
        if not excluded.empty and ctx.lead_config.get("save_excluded_companies", True):
            for record in excluded.to_dict("records"):
                excluded_company_store.add(
                    record,
                    str(record.get("company_type") or "Unknown"),
                    str(record.get("exclusion_reason") or "Excluded by ICP rules."),
                    dictionary_path,
                )

        if persistent_matches:
            ctx.add_artifact("persistent_dictionary_matches", persistent_matches)

        if not excluded.empty:
            ctx.append_excluded(excluded, self.name, "excluded by ICP rules")
            log.info(f"      excluded {len(excluded)} competitor companies")

            # Keep a compact, reviewable company-level watchlist for the
            # UI. This is separate from the reusable rule dictionary in
            # icp.py, so users can see both the rules and the actual
            # companies removed in the current run.
            excluded_rows = []
            for record in excluded.to_dict("records"):
                excluded_rows.append({
                    "company_name": record.get("company_name", ""),
                    "company_domain": record.get("company_domain", ""),
                    "company_type": record.get("company_type", "Unknown"),
                    "exclusion_reason": record.get("exclusion_reason", ""),
                })
            existing_watchlist = ctx.artifacts.get("persistent_dictionary_matches", [])
            combined = existing_watchlist + excluded_rows
            seen = set()
            deduped_watchlist = []
            for item in combined:
                key = (str(item.get("company_domain") or "").lower(), str(item.get("company_name") or "").lower())
                if key not in seen:
                    seen.add(key)
                    deduped_watchlist.append(item)
            ctx.add_artifact("excluded_company_watchlist", deduped_watchlist)
        else:
            ctx.add_artifact("excluded_company_watchlist", ctx.artifacts.get("persistent_dictionary_matches", []))

        ctx.add_artifact("persistent_excluded_companies", excluded_company_store.all_companies(dictionary_path))

        unknown_size = int((kept["company_size"] == icp.UNKNOWN).sum())
        if unknown_size:
            log.info(
                f"      {unknown_size} company(ies) have no reported employee count "
                "- kept and marked Unknown"
            )

        ctx.add_artifact("icp_stats", {
            "companies_in": len(df),
            "excluded": len(excluded),
            "qualified": len(kept),
            "unknown_size": unknown_size,
        })
        return kept.reset_index(drop=True)


class ContactEnrichmentStage(Stage):
    """Attach a real, publicly available contact route per company.

    Sources are the job advert's own contact details and the company's
    public careers page. No address, name or profile URL is ever
    constructed - companies with nothing findable are marked
    "Contact: Not Found" and stay in the list.
    """

    name = "Contact Enrichment"
    description = "Find public contact details from adverts and careers pages"

    def _run(self, ctx: PipelineContext) -> pd.DataFrame:
        if ctx.data.empty:
            return ctx.data

        settings = ctx.lead_config.get("contact_enrichment") or {}
        limit = settings.get("max_companies_to_enrich", 20)

        df = ctx.data.copy()
        records = df.to_dict("records")

        results = []
        for index, record in enumerate(records):
            if index >= limit:
                # Beyond the configured budget, report honestly rather
                # than pretending a lookup happened.
                results.append({
                    "contact_status": contacts.NOT_FOUND_STATUS,
                    "contact_name": "Not Found",
                    "contact_title": "Not Found",
                    "contact_email": "",
                    "contact_email_secondary": "",
                    "contact_linkedin_url": "",
                    "contact_source": "",
                    "contact_confidence": "None",
                    "contact_search_urls": contacts.linkedin_search_urls(
                        str(record.get("company_name") or ""),
                        settings.get("generate_linkedin_search_urls", True),
                    ),
                    "contact_sources_checked": "Not attempted (enrichment budget reached)",
                })
                continue
            results.append(contacts.enrich_company(record, ctx.lead_config))

        for field in ("contact_name", "contact_title", "contact_email",
                      "contact_email_secondary", "contact_linkedin_url",
                      "contact_source", "contact_confidence", "contact_status",
                      "contact_search_urls", "contact_sources_checked"):
            df[field] = [r.get(field, "") for r in results]

        found = int((df["contact_status"] == "Contact: Found").sum())
        log.info(f"      contacts found for {found}/{len(df)} companies")

        ctx.add_artifact("contact_stats", {
            "companies": len(df),
            "contacts_found": found,
            "contacts_not_found": len(df) - found,
        })
        return df


class LeadQualificationStage(Stage):
    """Derive hiring, funding and maturity signals, and the lead summary."""

    name = "Lead Qualification"
    description = "Derive hiring/funding/maturity signals and a lead summary"

    def _run(self, ctx: PipelineContext) -> pd.DataFrame:
        if ctx.data.empty:
            return ctx.data

        df = ctx.data.copy()
        signals = [lead_signals.derive_signals(row) for row in df.to_dict("records")]
        for field in ("team_maturity_signal", "ai_hiring_stage", "funding_signal",
                      "ai_role_count", "hiring_signal_strength", "hiring_signal",
                      "lead_summary"):
            df[field] = [s.get(field) for s in signals]

        with_funding = int((df["funding_signal"] != lead_signals.NOT_FOUND).sum())
        log.info(
            f"      signals derived; {with_funding}/{len(df)} companies state funding "
            "in their postings"
        )
        return df


class LeadPrioritizationStage(Stage):
    """Score every company and cut to the configured final lead count."""

    name = "Lead Prioritization"
    description = "Score, rank and select the final leads"

    def _run(self, ctx: PipelineContext) -> pd.DataFrame:
        if ctx.data.empty:
            return ctx.data

        df = ctx.data.copy()
        scores = [lead_signals.score_company(row, ctx.lead_config)
                  for row in df.to_dict("records")]
        for field in ("priority_score", "lead_priority", "priority_reason"):
            df[field] = [s.get(field) for s in scores]
        df["contact_strength"] = [lead_signals.contact_strength(row)
                                  for row in df.to_dict("records")]

        # Every scored company is kept for audit before the cut.
        ctx.add_artifact("all_company_signals", df.copy())

        selected, not_selected = lead_signals.select_final_leads(df, ctx.lead_config)
        ctx.add_artifact("not_selected_companies", not_selected)

        minimum = ctx.lead_config["min_final_leads"]
        counts = selected["lead_priority"].value_counts().to_dict()
        log.info(
            f"      selected {len(selected)} leads "
            f"(High {counts.get('High', 0)} / Medium {counts.get('Medium', 0)} / "
            f"Low {counts.get('Low', 0)})"
        )
        if len(selected) < minimum:
            log.warning(
                f"      [NOTE] only {len(selected)} qualified companies were available, "
                f"below the configured minimum of {minimum}. Reporting what is real "
                "rather than padding the list."
            )

        ctx.add_artifact("selection_stats", {
            "scored": len(df),
            "selected": len(selected),
            "not_selected": len(not_selected),
            "below_minimum": len(selected) < minimum,
            "minimum": minimum,
            "priority_counts": counts,
        })
        return selected


def build_pipeline(collect_fn, remote_filter_fn, dedupe_fn) -> list:
    """The pipeline, in execution order.

    Excel export is deliberately not a stage: it consumes the pipeline's
    result rather than transforming it, and keeping it outside means a
    half-built pipeline still produces the job spreadsheet that works
    today.
    """
    return [
        JobCollectionStage(collect_fn, dedupe_fn),
        RemoteFilteringStage(remote_filter_fn),
        CompanyNormalizationStage(),
        CompanyDeduplicationStage(),
        IcpFilteringStage(),
        ContactEnrichmentStage(),
        LeadQualificationStage(),
        LeadPrioritizationStage(),
    ]


# ---------------------------------------------------------------------
# Runner
# ---------------------------------------------------------------------

def run_pipeline(ctx: PipelineContext, stages: list) -> PipelineContext:
    """Execute stages in order, recording what each one did.

    A stage that raises is recorded as failed and the pipeline stops
    there, keeping whatever the previous stages produced - so a broken
    late stage still leaves an exportable result rather than losing the
    whole run.
    """
    log.info("\n" + "-" * 70)
    log.info("LEAD PIPELINE")
    log.info("-" * 70)

    for index, stage in enumerate(stages, start=1):
        rows_in = len(ctx.data)
        status = "ok"
        error = None

        if not stage.implemented:
            status = "pending"
            log.info(f"  [{index}/{len(stages)}] {stage.name:<28} -- not implemented yet, passing through")
        else:
            try:
                ctx.data = stage.run(ctx)
            except Exception as exc:
                status = "failed"
                error = f"{type(exc).__name__}: {exc}"
                log.warning(f"  [{index}/{len(stages)}] {stage.name:<28} FAILED - {error}")

        rows_out = len(ctx.data)
        ctx.stage_records.append({
            "stage": stage.name,
            "description": stage.description,
            "status": status,
            "rows_in": rows_in,
            "rows_out": rows_out,
            "error": error,
        })

        if status == "ok":
            delta = rows_out - rows_in
            change = f"{delta:+d}" if delta else "no change"
            log.info(f"  [{index}/{len(stages)}] {stage.name:<28} {rows_in:>4} -> {rows_out:>4}  ({change})")

        if status == "failed":
            log.warning("  Pipeline stopped early; exporting what earlier stages produced.")
            break

    return ctx


def print_pipeline_summary(ctx: PipelineContext) -> None:
    print("\nPipeline stages:")
    for record in ctx.stage_records:
        if record["status"] == "pending":
            mark = "pending"
            counts = "        "
        elif record["status"] == "failed":
            mark = "FAILED"
            counts = f"{record['rows_in']:>4} -> {record['rows_out']:>4}"
        else:
            mark = "done"
            counts = f"{record['rows_in']:>4} -> {record['rows_out']:>4}"
        print(f"  {record['stage']:<28} {counts}  [{mark}]")
        if record["error"]:
            print(f"      {record['error']}")

    pending = [r["stage"] for r in ctx.stage_records if r["status"] == "pending"]
    if pending:
        print(f"\n  Not built yet: {', '.join(pending)}")
        print("  These pass data through unchanged - no placeholder company or contact")
        print("  data is invented for them.")
