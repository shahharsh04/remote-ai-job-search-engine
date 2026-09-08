"""
Contact enrichment
==================
Finds a real, publicly available contact for each qualified company.

What this module will and will not do
-------------------------------------
It uses three sources, in order of reliability:

  1. Emails the job board already extracted from the advert
     (JobSpy's `emails` field) - real text from a public posting.
  2. The company's own public careers/contact page, fetched over plain
     HTTP when enabled in config, honouring robots.txt.
  3. LinkedIn SEARCH URLS for the recruiter-style titles in the brief.

Point 3 needs stating plainly: this produces a *search link*, not a
profile. LinkedIn profiles cannot be read without scraping behind a
login, which the brief forbids and this project does not do. A search URL
is a deterministic query a person can click; it never asserts that a
particular individual exists. Those links go in `contact_search_urls`,
and `contact_linkedin_url` is only ever filled from a real profile URL
found in public text.

Nothing here guesses a name, an email address or a profile. Patterns like
"firstname.lastname@company.com" are never constructed. When no contact
is found, `contact_status` is set to "Contact: Not Found".
"""

import logging
import re
from urllib.parse import quote_plus, urljoin, urlparse
from urllib.robotparser import RobotFileParser

log = logging.getLogger("job_search")

NOT_FOUND_STATUS = "Contact: Not Found"

EMAIL_PATTERN = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
LINKEDIN_PROFILE_PATTERN = re.compile(
    r"https?://(?:[a-z]{2,3}\.)?linkedin\.com/in/[A-Za-z0-9\-_%]+", re.I
)

# Applicant-tracking and HR-platform domains. An address here belongs to
# the software vendor, not the hiring company, so it ranks below a
# company-domain address.
ATS_DOMAINS = {
    "dayforce.com", "myworkday.com", "workday.com", "greenhouse.io",
    "lever.co", "icims.com", "taleo.net", "smartrecruiters.com",
    "workable.com", "bamboohr.com", "jobvite.com", "ashbyhq.com",
    "successfactors.com", "oraclecloud.com", "paylocity.com", "adp.com",
    "indeed.com", "linkedin.com", "ziprecruiter.com", "glassdoor.com",
}

# Legally-required accommodation and compliance mailboxes. Real, but the
# brief explicitly asks to avoid them when anything better exists.
LOW_VALUE_LOCAL_PARTS = {
    "accommodation", "accommodations", "ada", "eeo", "eeoc", "compliance",
    "disability", "accessibility", "privacy", "legal", "dpo",
    "unsubscribe", "noreply", "no-reply", "donotreply", "notify",
}

# Local parts that indicate a hiring mailbox, best first.
RECRUITING_LOCAL_PARTS = [
    "recruiting", "recruitment", "recruiter", "talent", "hiring",
    "jobs", "careers", "career", "apply", "work",
]

# Titles searched on LinkedIn, in the brief's order of preference.
LINKEDIN_SEARCH_TITLES = [
    "Technical Recruiter",
    "Recruiter",
    "Talent Acquisition",
    "Hiring Manager",
]

CAREERS_PATHS = ["/careers", "/careers/", "/jobs", "/company/careers", "/about/careers"]


def _clean(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return "" if text.lower() in ("nan", "none") else text


def _email_domain(email: str) -> str:
    return email.split("@")[-1].lower() if "@" in email else ""


def _registrable(domain: str) -> str:
    parts = domain.split(".")
    return ".".join(parts[-2:]) if len(parts) > 2 else domain


# Machine-generated mailbox names: tracking/relay addresses that are
# real but unusable as a contact. A live run surfaced Dropbox's
# "015d5ce7dd3142cd8fca094a50adbf69@d.dropbox.com" being offered as a
# lead contact, which is noise, not a hiring route.
HASH_LOCAL_PATTERN = re.compile(r"^[0-9a-f]{16,}$|^[0-9a-f]{8}-[0-9a-f-]{20,}$", re.I)


def is_machine_generated(email: str) -> bool:
    local = email.split("@")[0].lower()
    if HASH_LOCAL_PATTERN.match(local):
        return True
    # Long random-looking strings with no separator and plenty of digits.
    if len(local) >= 20 and sum(c.isdigit() for c in local) >= 6 and "." not in local:
        return True
    return False


def score_email(email: str, company_domain: str) -> tuple:
    """Rank an email. Higher is better; the reason is kept for the sheet.

    A company-domain recruiting mailbox beats a company-domain generic
    one, which beats an ATS vendor address, which beats a compliance
    mailbox. Nothing is discarded outright - a low-value address is still
    better than no contact at all.
    """
    local = email.split("@")[0].lower()
    domain = _email_domain(email)
    base_domain = _registrable(domain)

    if is_machine_generated(email):
        return 0, "machine-generated address - not a usable contact"

    if local in LOW_VALUE_LOCAL_PARTS or any(
        local.startswith(p) for p in ("accommodation", "noreply", "no-reply", "donotreply")
    ):
        return 10, "compliance/no-reply mailbox"

    is_ats = base_domain in ATS_DOMAINS
    matches_company = bool(company_domain) and base_domain == _registrable(company_domain)

    for index, part in enumerate(RECRUITING_LOCAL_PARTS):
        if part in local:
            if matches_company:
                return 100 - index, f"recruiting mailbox on the company domain"
            if is_ats:
                return 50 - index, "recruiting mailbox on an ATS vendor domain"
            return 70 - index, "recruiting mailbox"

    if matches_company:
        return 60, "company-domain address"
    if is_ats:
        return 30, "ATS vendor address"
    return 40, "address found in the posting"


def _extract_emails(text: str) -> list:
    if not text:
        return []
    seen = []
    for match in EMAIL_PATTERN.findall(text):
        cleaned = match.strip(".,;:)]}>\"'").lower()
        # Image filenames and the like occasionally survive the regex.
        if cleaned.endswith((".png", ".jpg", ".jpeg", ".gif", ".svg", ".webp")):
            continue
        if cleaned not in seen:
            seen.append(cleaned)
    return seen


def _extract_linkedin_profiles(text: str) -> list:
    if not text:
        return []
    seen = []
    for match in LINKEDIN_PROFILE_PATTERN.findall(text):
        url = match.rstrip(".,;:)]}>\"'")
        if url not in seen:
            seen.append(url)
    return seen


def linkedin_search_urls(company_name: str, enabled: bool) -> str:
    """Deterministic LinkedIn people-search links for a company.

    These are research aids, not discovered profiles - see the module
    docstring. Returned as a single joined string for the spreadsheet.
    """
    if not enabled or not company_name:
        return ""
    links = []
    for title in LINKEDIN_SEARCH_TITLES:
        query = quote_plus(f'{company_name} {title}')
        links.append(f"https://www.linkedin.com/search/results/people/?keywords={query}")
    return " | ".join(links)


def _robots_allows(session, base_url: str, path: str, timeout: int) -> bool:
    """Check robots.txt before fetching. On any doubt, allow - a missing
    or unreadable robots.txt is not a disallow - but a real Disallow is
    obeyed."""
    try:
        parsed = urlparse(base_url)
        robots_url = f"{parsed.scheme}://{parsed.netloc}/robots.txt"
        response = session.get(robots_url, timeout=timeout)
        if response.status_code != 200:
            return True
        parser = RobotFileParser()
        parser.parse(response.text.splitlines())
        return parser.can_fetch("*", urljoin(base_url, path))
    except Exception:
        return True


def fetch_careers_page_contacts(company_website: str, settings: dict) -> dict:
    """Fetch the company's own careers page and read contacts off it.

    Returns {"emails": [...], "linkedin": [...], "source": url} or an
    empty dict. Every failure is swallowed: enrichment is best-effort and
    must never take down a run.
    """
    website = _clean(company_website)
    if not website:
        return {}

    try:
        import requests
    except ImportError:
        log.warning("      [WARN] requests is unavailable - skipping careers-page lookup.")
        return {}

    timeout = settings.get("request_timeout_seconds", 8)
    if "://" not in website:
        website = "https://" + website

    headers = {
        "User-Agent": settings.get(
            "user_agent",
            "RemoteJobSearchEngine/1.0 (lead research; contact via job posting)",
        )
    }

    try:
        session = requests.Session()
        session.headers.update(headers)
    except Exception:
        return {}

    for path in [""] + CAREERS_PATHS:
        url = urljoin(website, path) if path else website
        if settings.get("respect_robots_txt", True):
            if not _robots_allows(session, website, path or "/", timeout):
                continue
        try:
            response = session.get(url, timeout=timeout)
        except Exception:
            continue
        if response.status_code != 200 or not response.text:
            continue

        text = response.text[:400_000]
        emails = _extract_emails(text)
        profiles = _extract_linkedin_profiles(text)
        if emails or profiles:
            return {"emails": emails, "linkedin": profiles, "source": url}

    return {}


def enrich_company(company: dict, lead_config: dict) -> dict:
    """Build the contact fields for one company from real sources only."""
    settings = lead_config.get("contact_enrichment") or {}
    company_name = _clean(company.get("company_name"))
    company_domain = _clean(company.get("company_domain"))

    candidates = []   # (score, email, reason, source)
    profiles = []
    sources_used = []

    # --- source 1: the job advert itself -----------------------------
    advert_text = " ".join([
        _clean(company.get("emails_found")),
        _clean(company.get("_description_blob")),
    ])
    for email in _extract_emails(advert_text):
        score, reason = score_email(email, company_domain)
        candidates.append((score, email, reason, "Job posting"))
    advert_profiles = _extract_linkedin_profiles(advert_text)
    if advert_profiles:
        profiles.extend(advert_profiles)
    if candidates or advert_profiles:
        sources_used.append("Job posting")

    # --- source 2: the company's public careers page ------------------
    if settings.get("enable_careers_page_lookup", True):
        found = fetch_careers_page_contacts(
            company.get("company_url_direct") or company_domain, settings
        )
        if found:
            for email in found.get("emails", []):
                score, reason = score_email(email, company_domain)
                # Slight preference for the company's own site over an
                # address copied into a job board advert.
                candidates.append((score + 5, email, reason, found["source"]))
            for profile in found.get("linkedin", []):
                if profile not in profiles:
                    profiles.append(profile)
            sources_used.append("Careers page")

    # --- assemble -----------------------------------------------------
    candidates.sort(key=lambda item: item[0], reverse=True)
    # Keep up to two distinct addresses, per the brief.
    chosen, seen_emails = [], set()
    for score, email, reason, source in candidates:
        if email in seen_emails:
            continue
        # A zero score means the address is real but unusable (machine
        # generated). Reporting "Not Found" is more honest than handing
        # the CEO a tracking address.
        if score <= 0:
            continue
        seen_emails.add(email)
        chosen.append((score, email, reason, source))
        if len(chosen) == 2:
            break

    result = {
        "contact_name": NOT_FOUND_STATUS.split(": ")[1],
        "contact_title": NOT_FOUND_STATUS.split(": ")[1],
        "contact_email": "",
        "contact_email_secondary": "",
        "contact_linkedin_url": "",
        "contact_source": "",
        "contact_confidence": "None",
        "contact_status": NOT_FOUND_STATUS,
        "contact_search_urls": linkedin_search_urls(
            company_name, settings.get("generate_linkedin_search_urls", True)
        ),
    }

    if profiles:
        result["contact_linkedin_url"] = " | ".join(profiles[:2])

    if chosen:
        top_score, top_email, top_reason, top_source = chosen[0]
        result["contact_email"] = top_email
        if len(chosen) > 1:
            result["contact_email_secondary"] = chosen[1][1]
        result["contact_source"] = top_source
        result["contact_status"] = "Contact: Found"

        # Confidence reflects how well the address identifies a hiring
        # route at THIS company - not a guess about the person.
        if top_score >= 90:
            result["contact_confidence"] = "High"
        elif top_score >= 55:
            result["contact_confidence"] = "Medium"
        else:
            result["contact_confidence"] = "Low"

        # A mailbox is a route, not a person. Names and titles stay
        # "Not Found" unless a real profile turned one up, because
        # inferring a person from "careers@" would be inventing one.
        result["contact_title"] = f"Hiring mailbox ({top_reason})"
    elif profiles:
        result["contact_status"] = "Contact: Found"
        result["contact_source"] = "Public profile link in posting/careers page"
        result["contact_confidence"] = "Medium"

    if sources_used:
        result["contact_sources_checked"] = ", ".join(dict.fromkeys(sources_used))
    else:
        result["contact_sources_checked"] = "Job posting (no contact present)"

    return result
