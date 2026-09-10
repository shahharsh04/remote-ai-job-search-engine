"""
Hiring signals, lead summaries, prioritisation and final selection
==================================================================
Turns a qualified company into a scored, explained lead.

Every signal is derived from data already collected - the number and
wording of the company's own job adverts, its reported employee count,
and any funding language in the advert text. There is no external
enrichment, so anything unavailable is reported as "Unknown" or
"Not Found" rather than filled in.

Scoring is table-driven: the weights and band cut-offs live in
config.yaml, and each contributing factor is a small named function, so
no single function carries the business rules.
"""

import re

UNKNOWN = "Unknown"
NOT_FOUND = "Not Found"

# Funding language that a company sometimes states in its own advert.
# Only an explicit statement counts - "well funded" is marketing, and
# "Series A" without context is not attributed to the company.
FUNDING_PATTERNS = [
    (re.compile(r"\bseries\s+([a-e])\b", re.I), "Series {0} mentioned in job posting"),
    (re.compile(r"\braised\s+\$?\s?([\d.]+)\s*(million|billion|m\b|bn\b)", re.I),
     "Raised {0} {1} (stated in job posting)"),
    (re.compile(r"\b(recently funded|newly funded|just raised|fresh funding)\b", re.I),
     "Recent funding stated in job posting"),
    (re.compile(r"\b(seed[- ]stage|pre[- ]seed)\b", re.I), "Seed stage stated in job posting"),
]

# Seniority wording used to judge whether an AI team already exists.
SENIOR_TITLE_PATTERN = re.compile(
    r"\b(head of|director|vp|vice president|principal|staff|lead|chief)\b", re.I
)
JUNIOR_TITLE_PATTERN = re.compile(r"\b(junior|entry[- ]level|graduate|intern)\b", re.I)

# Wording that indicates an established AI organisation.
ESTABLISHED_TEAM_PATTERNS = [
    "our ai team", "our ml team", "our machine learning team",
    "existing ai team", "our data science team", "join our team of engineers",
    "our platform team", "established ai team", "established ml team",
    "established machine learning team", "mature ai", "existing ml team",
]

# "team of 40 engineers" and similar - an explicit headcount for the
# team the role joins is strong evidence the organisation already exists.
ESTABLISHED_TEAM_SIZE_PATTERN = re.compile(
    r"\bteam of (\d+)\+?\s+(?:engineers|scientists|researchers|developers)\b", re.I
)
NEW_INITIATIVE_PATTERNS = [
    "first ai", "first machine learning", "first ml", "building out",
    "from the ground up", "greenfield", "new team", "founding engineer",
    "establish our", "help us build our first",
]

AI_ROLE_PATTERN = re.compile(
    r"\b(ai|a\.i\.|artificial intelligence|machine learning|ml|llm|genai|"
    r"generative ai|deep learning|nlp|computer vision|data scientist)\b", re.I
)


def _clean(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return "" if text.lower() in ("nan", "none") else text


# ---------------------------------------------------------------------
# Individual signals
# ---------------------------------------------------------------------

def count_ai_roles(titles: str) -> int:
    """How many of the company's matched openings are AI/ML roles."""
    if not titles:
        return 0
    return sum(1 for title in titles.split("|") if AI_ROLE_PATTERN.search(title))


def detect_funding_signal(text: str) -> str:
    """Funding stated in the company's own advert, else Not Found.

    There is no funding data source wired into this project - the boards
    return none - so a company with nothing in its advert text is
    reported as Not Found rather than researched or assumed.
    """
    blob = _clean(text)
    if not blob:
        return NOT_FOUND
    for pattern, template in FUNDING_PATTERNS:
        match = pattern.search(blob)
        if match:
            groups = [g.upper() if g and len(g) == 1 else g for g in match.groups()]
            try:
                return template.format(*groups)
            except (IndexError, KeyError):
                return template
    return NOT_FOUND


def detect_team_maturity(titles: str, text: str) -> str:
    """Whether the adverts suggest an existing AI org or a new one."""
    blob = _clean(text).lower()
    title_text = _clean(titles)

    for phrase in NEW_INITIATIVE_PATTERNS:
        if phrase in blob:
            return "Building new AI capability"
    for phrase in ESTABLISHED_TEAM_PATTERNS:
        if phrase in blob:
            return "Established AI team"

    size_match = ESTABLISHED_TEAM_SIZE_PATTERN.search(blob)
    if size_match and int(size_match.group(1)) >= 10:
        return "Established AI team"

    ai_roles = count_ai_roles(title_text)
    has_senior = bool(SENIOR_TITLE_PATTERN.search(title_text))

    if ai_roles >= 3 and has_senior:
        return "Established AI team"
    if ai_roles >= 2:
        return "Expanding AI team"
    if ai_roles == 1:
        return "Early AI hiring"
    return UNKNOWN


def detect_ai_hiring_stage(titles: str, maturity: str) -> str:
    ai_roles = count_ai_roles(_clean(titles))
    if ai_roles == 0:
        return "No AI/ML roles matched"
    if maturity == "Building new AI capability" or ai_roles == 1:
        return "First / early AI hires"
    if ai_roles >= 3:
        return "Scaling AI team"
    return "Expanding AI team"


def hiring_signal_strength(company: dict) -> tuple:
    """Rate the hiring signal Strong/Moderate/Weak, with the reason.

    More relevant openings mean a stronger signal, as the brief asks.
    """
    positions = int(company.get("matching_position_count") or 0)
    ai_roles = count_ai_roles(_clean(company.get("matching_job_titles")))
    maturity = company.get("team_maturity_signal", UNKNOWN)

    if ai_roles >= 3:
        return "Strong", f"{ai_roles} AI/ML openings"
    if ai_roles == 2:
        return "Moderate", "2 AI/ML openings"
    if ai_roles == 1:
        if maturity == "Building new AI capability":
            return "Strong", "first AI hire for a new capability"
        return "Moderate", "1 AI/ML opening"
    if positions >= 2:
        return "Weak", f"{positions} openings, none clearly AI/ML"
    return "Weak", "single opening, not clearly AI/ML"


def contact_strength(company: dict) -> str:
    confidence = _clean(company.get("contact_confidence"))
    if _clean(company.get("contact_status")).endswith("Not Found"):
        return "None"
    return confidence or "None"


# ---------------------------------------------------------------------
# Lead summary
# ---------------------------------------------------------------------

def build_lead_summary(company: dict) -> str:
    """One sentence on why this company might need AI engineering help.

    Assembled only from signals that were actually established; anything
    unknown is simply left out rather than described.
    """
    parts = []

    size = _clean(company.get("company_size"))
    employees = _clean(company.get("employee_count"))
    if size and size != UNKNOWN and employees and employees != UNKNOWN:
        parts.append(f"{size} company ({employees} employees)")
    elif size and size != UNKNOWN:
        parts.append(f"{size} company")
    else:
        parts.append("Company of unreported size")

    ai_roles = count_ai_roles(_clean(company.get("matching_job_titles")))
    positions = int(company.get("matching_position_count") or 0)
    maturity = _clean(company.get("team_maturity_signal"))

    if maturity == "Building new AI capability":
        parts.append("appears to be starting a new AI initiative")
    elif maturity == "Established AI team":
        parts.append("appears to have an established AI organisation")
    elif maturity == "Expanding AI team":
        parts.append("appears to be expanding an existing AI team")

    if ai_roles >= 2:
        parts.append(f"hiring {ai_roles} AI/ML roles")
    elif ai_roles == 1:
        parts.append("hiring an AI/ML role")
    elif positions:
        parts.append(f"{positions} matched opening(s), none clearly AI/ML")

    funding = _clean(company.get("funding_signal"))
    if funding and funding != NOT_FOUND:
        parts.append(funding.lower())

    summary = "; ".join(parts) + "."

    if maturity == "Established AI team":
        summary += " Lower priority for external augmentation."
    elif maturity in ("Building new AI capability", "Expanding AI team"):
        summary += " Potential need for external AI engineering support."

    return summary


def derive_signals(company: dict) -> dict:
    """All signal fields for one company."""
    titles = _clean(company.get("matching_job_titles"))
    blob = _clean(company.get("_description_blob"))

    maturity = detect_team_maturity(titles, blob)
    enriched = dict(company)
    enriched["team_maturity_signal"] = maturity

    signals = {
        "team_maturity_signal": maturity,
        "ai_hiring_stage": detect_ai_hiring_stage(titles, maturity),
        "funding_signal": detect_funding_signal(blob),
        "ai_role_count": count_ai_roles(titles),
    }

    strength, reason = hiring_signal_strength(enriched)
    signals["hiring_signal_strength"] = strength
    signals["hiring_signal"] = f"{strength} - {reason}"

    enriched.update(signals)
    signals["lead_summary"] = build_lead_summary(enriched)
    return signals


# ---------------------------------------------------------------------
# Scoring
# ---------------------------------------------------------------------

def _score_icp_fit(company, weights):
    band = _clean(company.get("company_size"))
    weight = weights.get("icp_fit", 0)
    if band == "Preferred ICP":
        return weight, "size in the preferred ICP band"
    if band == "Medium":
        return int(weight * 0.6), "medium-sized organisation"
    if band == "Large":
        return int(weight * 0.25), "large organisation - lower priority"
    if band == "Below ICP":
        return int(weight * 0.4), "smaller than the preferred band"
    # Unknown size must not be punished as if it were a bad fit.
    return int(weight * 0.5), "size unknown - confidence reduced, not excluded"


def _score_hiring_signal(company, weights):
    weight = weights.get("hiring_signal", 0)
    strength = _clean(company.get("hiring_signal_strength"))
    factor = {"Strong": 1.0, "Moderate": 0.6, "Weak": 0.2}.get(strength, 0.3)
    return int(weight * factor), f"{strength or UNKNOWN} hiring signal"


def _score_contact(company, weights):
    weight = weights.get("contact_found", 0)
    confidence = contact_strength(company)
    factor = {"High": 1.0, "Medium": 0.7, "Low": 0.4, "None": 0.0}.get(confidence, 0.0)
    if factor == 0.0:
        return 0, "no contact found"
    return int(weight * factor), f"contact found ({confidence.lower()} confidence)"


def _score_multiple_openings(company, weights):
    weight = weights.get("multiple_openings", 0)
    count = int(company.get("matching_position_count") or 0)
    if count >= 3:
        return weight, f"{count} matched openings"
    if count == 2:
        return int(weight * 0.6), "2 matched openings"
    return 0, "single opening"


def _score_funding(company, weights):
    weight = weights.get("funding_signal", 0)
    funding = _clean(company.get("funding_signal"))
    if funding and funding != NOT_FOUND:
        return weight, "funding signal present"
    return 0, "no funding signal available"


def _score_new_initiative(company, weights):
    weight = weights.get("new_ai_initiative", 0)
    maturity = _clean(company.get("team_maturity_signal"))
    if maturity == "Building new AI capability":
        return weight, "starting a new AI initiative"
    if maturity == "Expanding AI team":
        return int(weight * 0.6), "expanding an existing AI team"
    if maturity == "Established AI team":
        return 0, "established AI org - less need for augmentation"
    return int(weight * 0.3), "AI maturity unclear"


# Each scorer is small and independently testable; adding a factor means
# adding a function and a weight, not editing one large rule block.
SCORERS = [
    _score_icp_fit,
    _score_hiring_signal,
    _score_contact,
    _score_multiple_openings,
    _score_funding,
    _score_new_initiative,
]


def score_company(company: dict, lead_config: dict) -> dict:
    """Total the factors and assign a priority band."""
    scoring = lead_config.get("scoring") or {}
    weights = scoring.get("weights") or {}
    bands = scoring.get("priority_bands") or {}

    total = 0
    reasons = []
    for scorer in SCORERS:
        points, reason = scorer(company, weights)
        total += points
        if points:
            reasons.append(f"{reason} (+{points})")
        elif reason:
            reasons.append(reason)

    max_possible = sum(weights.values()) or 1
    normalised = round(100 * total / max_possible)

    high_min = bands.get("high_min_score", 70)
    medium_min = bands.get("medium_min_score", 45)

    if normalised >= high_min:
        priority = "High"
    elif normalised >= medium_min:
        priority = "Medium"
    else:
        priority = "Low"

    # Business-facing priority tiers are intentionally promoted one step:
    # High stays High, Medium becomes High, and Low becomes Medium.
    # This changes only the displayed/selected lead tier; the numeric
    # priority score remains unchanged for transparency.
    priority = {"High": "High", "Medium": "High", "Low": "Medium"}.get(priority, priority)

    return {
        "priority_score": normalised,
        "lead_priority": priority,
        "priority_reason": "; ".join(reasons),
    }


# ---------------------------------------------------------------------
# Final selection
# ---------------------------------------------------------------------

PRIORITY_RANK = {"High": 0, "Medium": 1, "Low": 2}
STRENGTH_RANK = {"Strong": 0, "Moderate": 1, "Weak": 2}


def select_final_leads(companies, lead_config: dict):
    """Rank and cut to the configured lead count.

    Returns (selected, not_selected). Both are kept so the companies that
    just missed the cut remain available for audit.
    """
    if companies is None or companies.empty:
        return companies, companies

    maximum = lead_config["max_final_leads"]

    ranked = companies.copy()
    ranked["_priority_rank"] = ranked["lead_priority"].map(PRIORITY_RANK).fillna(9)
    ranked["_strength_rank"] = ranked["hiring_signal_strength"].map(STRENGTH_RANK).fillna(9)
    ranked = ranked.sort_values(
        by=["_priority_rank", "priority_score", "_strength_rank", "matching_position_count"],
        ascending=[True, False, True, False],
    ).drop(columns=["_priority_rank", "_strength_rank"])

    selected = ranked.head(maximum)
    not_selected = ranked.iloc[maximum:]
    return selected.reset_index(drop=True), not_selected.reset_index(drop=True)
