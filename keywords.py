"""
Similar-job-title keyword expansion
===================================
Turns one job title into a small set of closely-related search keywords
so a search for "AI Engineer" also finds "Machine Learning Engineer",
"ML Engineer", "Applied AI Engineer" and so on.

Three rules shape everything here:

  1. The title the user typed is always keyword #1 and is never
     rewritten. Expansion only ever ADDS terms.
  2. Every candidate is validated before it is searched. A candidate
     that is too broad ("Engineer"), a different kind of role
     ("Data Analyst" for an engineering search) or malformed is
     dropped, whether it came from the built-in table, the config file
     or an AI model.
  3. Nothing here invents job listings. It produces search *terms*; the
     job boards remain the only source of postings.

Reuse:
    from keywords import KeywordExpander, expand_job_title

    expand_job_title("AI Engineer")                    # built-in defaults
    KeywordExpander(settings).expand("Data Scientist") # configured

`settings` is the `keyword_expansion` block of config.yaml, so a search
can be retuned (cap, extra synonyms, blocked terms, AI on/off) without
editing this file.
"""

import logging
import os
import re

log = logging.getLogger("job_search")


# ---------------------------------------------------------------------
# Vocabulary
# ---------------------------------------------------------------------
# Abbreviation <-> long form. Used for two different jobs: normalising
# titles so "ML Engineer" and "Machine Learning Engineer" compare equal,
# and generating the opposite surface form as a real search keyword
# (job boards index the two spellings differently).
ABBREVIATIONS = {
    "ai": "artificial intelligence",
    "ml": "machine learning",
    "dl": "deep learning",
    "nlp": "natural language processing",
    "llm": "large language model",
    "llms": "large language model",
    "genai": "generative artificial intelligence",
    "mlops": "machine learning operations",
    "llmops": "large language model operations",
    "cv": "computer vision",
    "rl": "reinforcement learning",
    "swe": "software engineer",
    "sde": "software development engineer",
    "bi": "business intelligence",
    "qa": "quality assurance",
    "sre": "site reliability engineer",
    "ui": "user interface",
    "ux": "user experience",
    "iot": "internet of things",
}

# Seniority and contract noise. Removed before comparing two titles so
# "Senior AI Engineer" still matches the AI-engineering family.
SENIORITY_TOKENS = {
    "senior", "sr", "junior", "jr", "staff", "principal", "lead", "head",
    "chief", "entry", "mid", "level", "associate", "experienced",
    "i", "ii", "iii", "iv", "v",
}
NOISE_TOKENS = {
    "the", "of", "and", "for", "a", "an", "to", "with",
    "remote", "hybrid", "onsite", "contract", "contractor", "permanent",
    "full", "part", "time", "fulltime", "parttime", "freelance",
    "hiring", "urgent", "immediate", "job", "jobs", "role", "position",
    "opening", "opportunity", "m", "f", "d", "x",
}

# Role-noun equivalence classes. A candidate keyword must sit in the
# SAME class as the searched title, which is what stops an engineering
# search from pulling in analyst, manager or designer roles.
ROLE_NOUN_CLASSES = {
    "engineer": {"engineer", "engineers", "engineering", "developer",
                 "developers", "programmer", "coder"},
    "scientist": {"scientist", "scientists"},
    "researcher": {"researcher", "researchers", "research"},
    "analyst": {"analyst", "analysts"},
    "architect": {"architect", "architects"},
    "manager": {"manager", "managers", "management"},
    "designer": {"designer", "designers"},
    "administrator": {"administrator", "administrators", "admin"},
    "consultant": {"consultant", "consultants"},
    "specialist": {"specialist", "specialists"},
    "technician": {"technician", "technicians"},
    "director": {"director", "directors"},
}

# Interchangeable role nouns used for rule-generated variants. Only
# swaps that name the same job on a job board are listed - "engineer"
# and "developer" are, "engineer" and "manager" are not.
ROLE_NOUN_SYNONYMS = {
    "engineer": ["developer"],
    "developer": ["engineer"],
    "scientist": ["researcher"],
    "researcher": ["scientist"],
}

# Curated families of genuinely interchangeable titles. A family is
# matched when the searched title and one of the family's titles share
# their normalised wording (after abbreviations and seniority are
# stripped), so every entry below is reachable from any other entry.
#
# Add a family, or a title to an existing family, to widen coverage.
# Keep entries specific: a family is a list of titles a recruiter would
# use for the SAME job, not a list of titles in the same department.
ROLE_FAMILIES = {
    "ai_engineering": [
        "AI Engineer",
        "Machine Learning Engineer",
        "Artificial Intelligence Engineer",
        "Applied AI Engineer",
        "ML Engineer",
        "AI/ML Engineer",
        "Applied Machine Learning Engineer",
        "AI Software Engineer",
        "Deep Learning Engineer",
        "Generative AI Engineer",
        "LLM Engineer",
        "AI Developer",
        "Machine Learning Developer",
    ],
    "mlops_engineering": [
        "MLOps Engineer",
        "Machine Learning Operations Engineer",
        "ML Platform Engineer",
        "Machine Learning Infrastructure Engineer",
        "AI Platform Engineer",
        "AI Infrastructure Engineer",
        "LLMOps Engineer",
    ],
    "data_science": [
        "Data Scientist",
        "Machine Learning Scientist",
        "Applied Scientist",
        "Applied Data Scientist",
        "AI Scientist",
        "Decision Scientist",
    ],
    "ai_research": [
        "Research Scientist",
        "AI Research Scientist",
        "Machine Learning Researcher",
        "AI Researcher",
        "Research Engineer",
        "Machine Learning Research Engineer",
    ],
    "nlp_engineering": [
        "NLP Engineer",
        "Natural Language Processing Engineer",
        "Computational Linguistics Engineer",
        "Conversational AI Engineer",
        "LLM Engineer",
    ],
    "computer_vision": [
        "Computer Vision Engineer",
        "CV Engineer",
        "Machine Vision Engineer",
        "Perception Engineer",
        "Image Processing Engineer",
    ],
    "data_engineering": [
        "Data Engineer",
        "Big Data Engineer",
        "Data Platform Engineer",
        "Analytics Engineer",
        "ETL Developer",
        "Data Infrastructure Engineer",
    ],
    "data_analysis": [
        "Data Analyst",
        "Business Intelligence Analyst",
        "BI Analyst",
        "Analytics Analyst",
        "Reporting Analyst",
    ],
    "backend_engineering": [
        "Backend Engineer",
        "Back End Developer",
        "Backend Developer",
        "Server Side Engineer",
        "API Engineer",
    ],
    "frontend_engineering": [
        "Frontend Engineer",
        "Front End Developer",
        "Frontend Developer",
        "UI Engineer",
        "Web Developer",
    ],
    "fullstack_engineering": [
        "Full Stack Engineer",
        "Fullstack Developer",
        "Full Stack Developer",
        "Full Stack Web Developer",
    ],
    "devops_engineering": [
        "DevOps Engineer",
        "Site Reliability Engineer",
        "Platform Engineer",
        "Infrastructure Engineer",
        "Cloud Infrastructure Engineer",
    ],
    "cloud_engineering": [
        "Cloud Engineer",
        "Cloud Solutions Engineer",
        "AWS Engineer",
        "Azure Engineer",
        "Cloud Platform Engineer",
    ],
    "mobile_engineering": [
        "Mobile Engineer",
        "Mobile Developer",
        "iOS Developer",
        "Android Developer",
        "React Native Developer",
    ],
    "qa_engineering": [
        "QA Engineer",
        "Quality Assurance Engineer",
        "Test Engineer",
        "Automation Test Engineer",
        "Software Test Engineer",
    ],
    "security_engineering": [
        "Security Engineer",
        "Cyber Security Engineer",
        "Application Security Engineer",
        "Information Security Engineer",
    ],
    "product_management": [
        "Product Manager",
        "Technical Product Manager",
        "AI Product Manager",
        "Product Owner",
    ],
}

# Terms too broad to search on their own: they would return everything
# and defeat the point of expanding. A candidate equal to one of these
# after normalisation is rejected even if a family or an AI model
# offered it.
DEFAULT_BLOCKED_TERMS = [
    "engineer", "developer", "programmer", "scientist", "analyst",
    "manager", "consultant", "architect", "specialist", "designer",
    "software engineer", "software developer", "technology", "tech",
    "information technology", "computer science", "data",
    "artificial intelligence", "machine learning", "deep learning",
    "intern", "internship", "graduate", "trainee", "remote",
    "remote job", "work from home",
]

DEFAULTS = {
    "enabled": True,
    # Total keywords searched, INCLUDING the user's own title.
    "max_keywords": 6,
    # Generate "Engineer" <-> "Developer" style variants.
    "role_noun_synonyms": True,
    # Generate the opposite spelling of an abbreviation
    # ("AI Engineer" -> "Artificial Intelligence Engineer").
    "abbreviation_variants": True,
    # Carry "Senior"/"Lead" onto generated keywords. Off by default:
    # boards match seniority loosely and it narrows results sharply.
    "preserve_seniority": False,
    # {"searched title": ["extra keyword", ...]} - trusted, format
    # checked only, so a user can force in a term the tables lack.
    "extra_synonyms": {},
    # Added to DEFAULT_BLOCKED_TERMS rather than replacing them.
    "blocked_terms": [],
    "ai": {
        "enabled": False,
        "provider": "anthropic",
        "model": "claude-sonnet-5",
        # Upper bound on what the model may propose. Whatever comes back
        # still goes through validation before it is searched.
        "max_suggestions": 8,
        "timeout_seconds": 20,
        "api_key_env": "ANTHROPIC_API_KEY",
    },
}


def load_keyword_config(config: dict) -> dict:
    """Read the `keyword_expansion` block, filling in every default.

    A missing or partial block is normal, not an error - the built-in
    defaults are a working configuration on their own.
    """
    supplied = (config or {}).get("keyword_expansion") or {}
    if not isinstance(supplied, dict):
        log.warning("[WARN] keyword_expansion in config is not a mapping - using defaults.")
        supplied = {}

    settings = {k: v for k, v in DEFAULTS.items() if k != "ai"}
    settings.update({k: v for k, v in supplied.items() if k != "ai"})

    ai = dict(DEFAULTS["ai"])
    supplied_ai = supplied.get("ai") or {}
    if isinstance(supplied_ai, dict):
        ai.update(supplied_ai)
    settings["ai"] = ai

    try:
        settings["max_keywords"] = max(1, int(settings["max_keywords"]))
    except (TypeError, ValueError):
        log.warning("[WARN] keyword_expansion.max_keywords is not a number - using 6.")
        settings["max_keywords"] = 6

    if not isinstance(settings.get("extra_synonyms"), dict):
        settings["extra_synonyms"] = {}
    if not isinstance(settings.get("blocked_terms"), list):
        settings["blocked_terms"] = []

    return settings


# ---------------------------------------------------------------------
# Normalisation
# ---------------------------------------------------------------------
def _tokenize(title: str) -> list:
    """Lowercase word tokens. Punctuation ('/', '-', ',') becomes a
    separator so "AI/ML Engineer" splits into three tokens."""
    return [t for t in re.split(r"[^a-z0-9+#]+", str(title or "").lower()) if t]


def normalize_title(title: str) -> str:
    """Comparison form of a title: abbreviations expanded, seniority and
    filler removed. "Sr. ML Engineer (Remote)" -> "machine learning engineer".

    Only ever used for comparing; the searched keyword is always the
    original spelling, because the two spellings return different rows.
    """
    tokens = []
    for token in _tokenize(title):
        expanded = ABBREVIATIONS.get(token, token)
        tokens.extend(expanded.split())
    return " ".join(
        t for t in tokens if t not in SENIORITY_TOKENS and t not in NOISE_TOKENS
    )


def _token_set(title: str) -> set:
    return set(normalize_title(title).split())


# Every token that names a role, flattened once for subject-word checks.
ALL_ROLE_NOUNS = {t for members in ROLE_NOUN_CLASSES.values() for t in members}


def _role_noun_class(title: str):
    """Which role class a title names, or None if it names none. The
    LAST matching token wins: in "Machine Learning Engineer" the role is
    the trailing noun."""
    found = None
    for token in normalize_title(title).split():
        for class_name, members in ROLE_NOUN_CLASSES.items():
            if token in members:
                found = class_name
    return found


def _seniority_prefix(title: str) -> str:
    """The leading seniority word as the user typed it, if any."""
    tokens = _tokenize(title)
    if tokens and tokens[0] in SENIORITY_TOKENS and tokens[0] not in ("i", "ii", "iii", "iv", "v"):
        return tokens[0].capitalize()
    return ""


def _relevance(original: str, candidate: str) -> float:
    """Overlap of the two normalised token sets (Jaccard). Used only to
    order candidates, never to accept or reject one."""
    a, b = _token_set(original), _token_set(candidate)
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)


# ---------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------
class KeywordValidator:
    """Decides whether a candidate keyword is safe to search.

    Every source of candidates - the curated table, the rule engine and
    the optional AI model - passes through this, so a model that
    suggests "Chief Happiness Officer" for an AI Engineer search cannot
    reach a job board.
    """

    def __init__(self, original: str, settings: dict):
        self.original = original
        self.original_norm = normalize_title(original)
        self.original_tokens = _token_set(original)
        self.original_class = _role_noun_class(original)
        self.blocked = {
            normalize_title(t)
            for t in list(DEFAULT_BLOCKED_TERMS) + list(settings.get("blocked_terms") or [])
        }
        # Normalised forms of the family the search belongs to. A
        # candidate from the same family is related by definition, even
        # when it shares no words ("AI Engineer" / "Machine Learning
        # Engineer").
        self.family_norms = set()

    def register_family(self, titles) -> None:
        self.family_norms.update(normalize_title(t) for t in titles)

    def check_format(self, candidate: str) -> tuple:
        """Shape checks only. Returns (ok, reason)."""
        text = re.sub(r"\s+", " ", str(candidate or "")).strip()
        if not text:
            return False, "empty"
        if len(text) > 60:
            return False, "too long"
        if not re.fullmatch(r"[A-Za-z0-9 /+#&.'-]+", text):
            return False, "unexpected characters"
        words = text.split()
        if len(words) < 2 or len(words) > 6:
            return False, "not 2-6 words"
        if normalize_title(text) in self.blocked:
            return False, "too broad"
        return True, ""

    def check_relevance(self, candidate: str) -> tuple:
        """Relatedness checks. Returns (ok, reason)."""
        norm = normalize_title(candidate)
        if not norm:
            return False, "nothing left after normalisation"
        if norm == self.original_norm:
            # Not a rejection of the idea, just an equivalent spelling of
            # the searched title - "ML Engineer" for "Machine Learning
            # Engineer". Boards index the two differently, so it is a
            # useful keyword; the caller dedupes on the literal string.
            return True, ""

        candidate_class = _role_noun_class(candidate)
        if self.original_class and candidate_class != self.original_class:
            return False, f"different kind of role ({candidate_class or 'none'})"

        if norm in self.family_norms:
            return True, ""

        # Outside the family, require a shared subject word so we never
        # drift into an unrelated specialism.
        shared = (self.original_tokens & _token_set(candidate)) - ALL_ROLE_NOUNS
        if not shared:
            return False, "no shared subject with the searched title"
        return True, ""

    def validate(self, candidate: str, trusted: bool = False) -> tuple:
        ok, reason = self.check_format(candidate)
        if not ok:
            return False, reason
        if trusted:
            # Explicit config entries: the user asked for them by name,
            # so only the shape is enforced.
            return True, ""
        return self.check_relevance(candidate)


# ---------------------------------------------------------------------
# Expansion
# ---------------------------------------------------------------------
class KeywordExpander:
    """Expands one job title into related search keywords.

    Stateless between calls and configured entirely by `settings`, so
    one instance can serve many searches, and different searches can use
    different settings.
    """

    def __init__(self, settings: dict = None, ai_client=None):
        self.settings = settings if settings is not None else load_keyword_config({})
        # Injectable so the AI path can be exercised without a network.
        self.ai_client = ai_client

    # -- sources ------------------------------------------------------
    def _matching_families(self, title: str) -> list:
        """Families the searched title belongs to, best match first.

        A family matches when one of its titles, normalised, is
        contained in the searched title - so "AI Engineer",
        "Senior AI Engineer" and "Applied AI Engineer" all land on the
        AI engineering family.

        The containment runs one way only, family-title inside searched
        title. The reverse ("AI Engineer" matching the MLOps family
        through "AI Platform Engineer") would pull in narrower
        specialisms the user did not ask for, which is exactly the
        irrelevant-result case this expansion has to avoid.

        Returns [(name, titles, rank)], rank ordering families by how
        specifically they matched: the longest matched family title
        wins, ties broken by declaration order.
        """
        wanted = _token_set(title)
        if not wanted:
            return []

        matches = []
        for order, (name, titles) in enumerate(ROLE_FAMILIES.items()):
            best = 0
            for member in titles:
                tokens = _token_set(member)
                if tokens and tokens <= wanted:
                    best = max(best, len(tokens))
            if best:
                matches.append((name, titles, (-best, order)))

        matches.sort(key=lambda m: m[2])
        return matches

    def _rule_variants(self, title: str) -> list:
        """Surface variants of the searched title itself: the opposite
        abbreviation spelling, and interchangeable role nouns.

        Substitutions are made on the original string so the parts that
        were not swapped keep the user's own capitalisation -
        "NLP Engineer" becomes "NLP Developer", not "Nlp Developer".
        """
        variants = []
        lowered = title.lower()

        def swap(text: str, old: str, new: str) -> str:
            return re.sub(rf"\b{re.escape(old)}\b", new, text, flags=re.IGNORECASE)

        if self.settings.get("abbreviation_variants", True):
            # Abbreviation -> long form ("AI Engineer").
            expanded = title
            for short, long_form in ABBREVIATIONS.items():
                if re.search(rf"\b{re.escape(short)}\b", lowered):
                    expanded = swap(expanded, short, long_form.title())
            if expanded != title:
                variants.append(expanded)

            # Long form -> abbreviation ("Machine Learning Engineer").
            for short, long_form in ABBREVIATIONS.items():
                if " " in long_form and long_form in lowered:
                    variants.append(swap(title, long_form, short.upper()))

        if self.settings.get("role_noun_synonyms", True):
            for noun, synonyms in ROLE_NOUN_SYNONYMS.items():
                if not re.search(rf"\b{re.escape(noun)}\b", lowered):
                    continue
                for synonym in synonyms:
                    variants.append(swap(title, noun, synonym.title()))

        return variants

    def _extra_synonyms(self, title: str) -> list:
        """Config-supplied keywords for this title. Matched on the
        normalised form so casing and "Sr." spelling do not matter."""
        wanted = normalize_title(title)
        out = []
        for key, values in (self.settings.get("extra_synonyms") or {}).items():
            if normalize_title(key) != wanted:
                continue
            if isinstance(values, str):
                values = [values]
            out.extend(str(v) for v in (values or []))
        return out

    def _ai_suggestions(self, title: str) -> list:
        """Ask a model for extra titles. Optional, off by default, and
        every suggestion is validated by the caller before use.

        The model is asked for job TITLES only - it is never asked for,
        and never a source of, job listings.
        """
        ai = self.settings.get("ai") or {}
        if not ai.get("enabled"):
            return []

        limit = int(ai.get("max_suggestions", 8))
        prompt = (
            f'List up to {limit} alternative job titles that recruiters use for the '
            f'same role as "{title}". Include abbreviations and expanded forms. '
            "Only titles for the same kind of role - no broader or adjacent roles, "
            "no seniority-only variants, no company names, no explanations. "
            "Reply with one job title per line and nothing else."
        )

        try:
            if self.ai_client is not None:
                text = self.ai_client(prompt)
            elif ai.get("provider") == "anthropic":
                text = self._call_anthropic(prompt, ai)
            else:
                log.warning(
                    f"[WARN] keyword_expansion.ai.provider='{ai.get('provider')}' is not "
                    "supported - skipping AI keyword suggestions."
                )
                return []
        except Exception as exc:
            # An unreachable model must never fail a search; the curated
            # tables already produced usable keywords.
            log.warning(
                f"[WARN] AI keyword suggestion failed ({exc}) - using built-in tables only."
            )
            return []

        lines = []
        for line in str(text or "").splitlines():
            cleaned = re.sub(r"^\s*(?:[-*]|\d+[.)])\s*", "", line).strip(" \t\"'`")
            if cleaned:
                lines.append(cleaned)
        return lines[:limit]

    def _call_anthropic(self, prompt: str, ai: dict) -> str:
        api_key = os.getenv(ai.get("api_key_env", "ANTHROPIC_API_KEY"), "")
        if not api_key:
            raise RuntimeError(f"{ai.get('api_key_env', 'ANTHROPIC_API_KEY')} is not set")

        import anthropic  # optional dependency - imported only when AI is enabled

        client = anthropic.Anthropic(
            api_key=api_key, timeout=float(ai.get("timeout_seconds", 20))
        )
        message = client.messages.create(
            model=ai.get("model", "claude-sonnet-5"),
            max_tokens=300,
            messages=[{"role": "user", "content": prompt}],
        )
        return "".join(
            block.text for block in message.content
            if getattr(block, "type", "") == "text"
        )

    # -- public API ---------------------------------------------------
    def expand(self, title: str) -> dict:
        """Expand `title` into related keywords.

        Returns a report rather than a bare list so callers can log and
        export exactly what was searched and what was rejected:

            {"original", "keywords", "expanded", "rejected", "sources",
             "families", "enabled"}

        `keywords[0]` is always the original title, unchanged.
        """
        original = re.sub(r"\s+", " ", str(title or "")).strip()
        report = {
            "original": original,
            "keywords": [original] if original else [],
            "expanded": [],
            "rejected": [],
            "sources": {},
            "families": [],
            "enabled": bool(self.settings.get("enabled", True)),
        }
        if not original or not report["enabled"]:
            return report

        validator = KeywordValidator(original, self.settings)

        families = self._matching_families(original)
        report["families"] = [name for name, _, _ in families]
        for _, titles, _ in families:
            validator.register_family(titles)

        # (sort key, source, candidate). The sort key ranks by how much
        # the source is trusted, then - within a curated family - by the
        # hand-ordered list, so the closest alternative titles are the
        # ones that survive the max_keywords cap. Free-form sources
        # (config, AI) fall back to word overlap with the searched title.
        candidates = []
        for family_index, (_, titles, _) in enumerate(families):
            for position, member in enumerate(titles):
                # A family title that normalises to what the user typed
                # is just another spelling of their own search ("ML
                # Engineer" for "Senior Machine Learning Engineer"), so
                # it outranks the rest of the family.
                equivalent = 0 if normalize_title(member) == validator.original_norm else 1
                candidates.append(
                    ((0, family_index, equivalent, position, 0.0), "family", member)
                )
        for position, variant in enumerate(self._rule_variants(original)):
            candidates.append(((1, 0, 0, position, 0.0), "rule", variant))
        for position, extra in enumerate(self._extra_synonyms(original)):
            candidates.append(((2, 0, 0, position, 0.0), "config", extra))
        for suggestion in self._ai_suggestions(original):
            candidates.append(
                ((3, 0, 0, 0, -_relevance(original, suggestion)), "ai", suggestion)
            )

        candidates.sort(key=lambda c: c[0])

        accepted = {}                       # keyword -> source
        seen_literal = {original.lower()}
        ordered = []
        for _, source, candidate in candidates:
            text = re.sub(r"\s+", " ", str(candidate or "")).strip()
            key = text.lower()
            if key in seen_literal:
                continue
            ok, reason = validator.validate(text, trusted=(source == "config"))
            if not ok:
                report["rejected"].append(
                    {"keyword": text, "source": source, "reason": reason}
                )
                continue
            seen_literal.add(key)
            accepted[text] = source
            ordered.append(text)

        if self.settings.get("preserve_seniority"):
            prefix = _seniority_prefix(original)
            if prefix:
                # Prefixing can collide - "AI Engineer" becomes "Senior
                # AI Engineer", which is the searched title itself - so
                # dedupe again afterwards, keeping the first spelling.
                prefixed = []
                seen = {original.lower()}
                for keyword in ordered:
                    text = (
                        keyword if keyword.lower().startswith(prefix.lower())
                        else f"{prefix} {keyword}"
                    )
                    if text.lower() in seen:
                        continue
                    seen.add(text.lower())
                    accepted[text] = accepted.get(keyword, "rule")
                    prefixed.append(text)
                ordered = prefixed

        limit = max(0, int(self.settings["max_keywords"]) - 1)
        kept = ordered[:limit]
        for dropped in ordered[limit:]:
            report["rejected"].append({
                "keyword": dropped,
                "source": accepted.get(dropped, "unknown"),
                "reason": f"over the max_keywords limit of {self.settings['max_keywords']}",
            })

        report["expanded"] = kept
        report["keywords"] = [original] + kept
        report["sources"] = {
            original: "user",
            **{k: accepted.get(k, "unknown") for k in kept},
        }
        return report


def expand_job_title(title: str, settings: dict = None) -> dict:
    """One-shot convenience wrapper around KeywordExpander."""
    return KeywordExpander(settings).expand(title)
