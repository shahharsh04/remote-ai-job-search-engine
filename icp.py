"""
ICP classification and exclusion filtering
==========================================
Decides, per company, whether it is a prospect or a competitor, and which
size band it falls into.

Every judgement here is derived from text the job boards actually
returned - company name, industry, company description, employee-count
string, and the job descriptions themselves. Nothing is looked up
elsewhere and nothing is guessed. When a field is absent, it is reported
as "Unknown"/"Not Found" and the company is KEPT: missing data is not
evidence of a bad fit.
"""

import re

UNKNOWN = "Unknown"
NOT_FOUND = "Not Found"

# Company types whose primary business competes with an AI engineering
# and services offering. Each entry lists phrases that identify it.
#
# These match against the company's own name, industry and description -
# NOT the body of its job adverts, which mention "consulting" and
# "services" far too often to be safe evidence.
COMPETITOR_TYPE_PATTERNS = {
    "staffing_agency": [
        "staffing", "staff augmentation", "temp agency", "temporary staffing",
        "contract staffing", "manpower", "workforce solutions", "talent solutions",
    ],
    "recruitment_agency": [
        "recruitment agency", "recruiting agency", "recruitment firm",
        "executive search", "headhunt", "talent acquisition firm",
        "recruitment consultancy", "placement agency",
    ],
    "it_services": [
        "it services", "information technology services", "system integrator",
        "systems integrator", "managed services provider", "it consulting",
        "technology services provider", "it solutions provider",
    ],
    "ai_ml_consulting": [
        "ai consulting", "ml consulting", "machine learning consulting",
        "data science consulting", "ai solutions provider", "ai services company",
        "artificial intelligence consultancy", "ai consultancy",
    ],
    "software_outsourcing": [
        "software outsourcing", "outsourcing", "offshore development",
        "nearshore development", "software development agency",
        "product engineering services", "software house", "dev shop",
        "digital agency", "software consultancy", "engineering services provider",
    ],
}

# Phrases in a JOB ADVERT that reliably mark the poster as an
# intermediary hiring on someone else's behalf. Unlike the generic words
# above, these are specific enough to trust from advert text.
INTERMEDIARY_JOB_TEXT_PATTERNS = [
    "our client is seeking", "our client is looking", "on behalf of our client",
    "we are recruiting on behalf", "our client, a", "client is seeking",
    "placement with our client", "contract-to-hire with our client",
]

# Words that mark a name as an agency even without a description.
AGENCY_NAME_PATTERNS = [
    "staffing", "recruit", "talent", "headhunt", "manpower",
    "consultancy", "consulting", "outsourcing", "technologies services",
]


def _clean(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return "" if text.lower() in ("nan", "none") else text


def parse_employee_count(raw):
    """Turn a board's employee-count string into (display, estimate).

    JobSpy returns things like "51 to 200", "1,001 to 5,000", "10,000+".
    `display` keeps the original wording so nothing is misrepresented;
    `estimate` is a single number used only for banding, taken as the
    midpoint of a reported range or the bound of an open-ended one.

    Returns (UNKNOWN, None) when there is nothing to parse - never a
    guessed headcount.
    """
    text = _clean(raw)
    if not text:
        return UNKNOWN, None

    numbers = [int(n.replace(",", "")) for n in re.findall(r"\d[\d,]*", text)]
    if not numbers:
        return text, None

    if len(numbers) >= 2:
        estimate = (numbers[0] + numbers[1]) // 2
    elif "+" in text or "more than" in text.lower() or "over" in text.lower():
        estimate = numbers[0]
    else:
        estimate = numbers[0]

    return text, estimate


def size_band(estimate, lead_config: dict) -> str:
    """Name the company-size band. All cut-offs come from config."""
    if estimate is None:
        return UNKNOWN

    preferred_min = lead_config["preferred_min_employees"]
    preferred_max = lead_config["preferred_max_employees"]
    medium_threshold = lead_config["medium_priority_employee_threshold"]
    low_threshold = lead_config["low_priority_employee_threshold"]

    if estimate < preferred_min:
        return "Below ICP"
    if estimate <= preferred_max:
        return "Preferred ICP"
    if estimate < low_threshold:
        # Between the preferred band and the large-company threshold.
        return "Medium" if estimate >= medium_threshold else "Preferred ICP"
    return "Large"


def classify_company_type(company: dict) -> tuple:
    """Identify the company's business type from its own description.

    Returns (company_type, matched_phrase, evidence_field). The evidence
    is carried so an exclusion can always be explained and challenged.
    """
    name = _clean(company.get("company_name")).lower()
    industry = _clean(company.get("company_industry")).lower()
    description = _clean(company.get("company_description")).lower()

    # Industry and description are the company's own words about itself.
    for field_name, haystack in (("company_industry", industry),
                                 ("company_description", description)):
        if not haystack:
            continue
        for company_type, phrases in COMPETITOR_TYPE_PATTERNS.items():
            for phrase in phrases:
                if phrase in haystack:
                    return company_type, phrase, field_name

    for company_type, phrases in COMPETITOR_TYPE_PATTERNS.items():
        for phrase in phrases:
            if phrase in name:
                return company_type, phrase, "company_name"

    # Adverts written on a client's behalf mark an intermediary.
    job_text = _clean(company.get("_description_blob")).lower()
    if job_text:
        for phrase in INTERMEDIARY_JOB_TEXT_PATTERNS:
            if phrase in job_text:
                return "staffing_agency", phrase, "job_description"

    for phrase in AGENCY_NAME_PATTERNS:
        if re.search(rf"\b{re.escape(phrase)}\b", name):
            return "possible_agency", phrase, "company_name"

    if industry or description:
        return "operating_company", "", "company_industry/description"

    return UNKNOWN, "", ""


def evaluate_company(company: dict, lead_config: dict) -> dict:
    """Full ICP verdict for one company row.

    Adds employee_count, company_size, company_type, icp_status and
    exclusion_reason. Large companies are KEPT (they are simply banded
    "Large" and scored lower later); only competitor types are excluded.
    """
    display_count, estimate = parse_employee_count(company.get("company_num_employees"))
    band = size_band(estimate, lead_config)

    company_type, phrase, evidence_field = classify_company_type(company)

    excluded_types = set(lead_config.get("excluded_company_types") or [])
    is_competitor = company_type in excluded_types

    result = {
        "employee_count": display_count,
        "employee_count_estimate": estimate,
        "company_size": band,
        "company_type": company_type if company_type != UNKNOWN else UNKNOWN,
        "company_type_evidence": (
            f"matched '{phrase}' in {evidence_field}" if phrase else NOT_FOUND
        ),
    }

    if is_competitor:
        result["icp_status"] = "Excluded"
        result["exclusion_reason"] = (
            f"{company_type.replace('_', ' ')} - matched '{phrase}' in {evidence_field}"
        )
        return result

    # Everything that is not a competitor stays in, including large
    # organisations and companies we could not classify at all.
    result["exclusion_reason"] = ""
    if band == "Preferred ICP":
        result["icp_status"] = "Preferred ICP"
    elif band == "Medium":
        result["icp_status"] = "Qualified - medium size"
    elif band == "Large":
        result["icp_status"] = "Qualified - large organisation"
    elif band == "Below ICP":
        result["icp_status"] = "Qualified - below preferred size"
    else:
        result["icp_status"] = "Qualified - size unknown"

    return result
