"""
Company identity resolution
===========================
Turning job rows into company rows needs a reliable answer to "are these
two postings from the same company?".

The rule here is deliberately conservative: two rows are the same company
only if they share an EXACT identifier after normalisation. There is no
fuzzy or similarity matching, because "Delta Systems" and "Delta Systems
Group" may well be different businesses, and silently merging them would
put the wrong company in front of the CEO.

Three identifier types are used, all derived from data the job boards
actually return - nothing is invented or guessed:

  1. website domain  - from company_url_direct, the company's own site.
                       The strongest signal: two postings pointing at
                       acme.com are the same Acme.
  2. profile URL     - from company_url, the board's own company page
                       (linkedin.com/company/acme, indeed.com/cmp/Acme).
                       An exact identifier within that board.
  3. normalised name - casing, accents, punctuation, spacing and legal
                       suffixes removed.

Rows sharing any one of these are treated as one company, which is what
lets a domain merge minor name variations (requirement 4) without any
name-similarity guessing (requirement 5).
"""

import re
import unicodedata
from urllib.parse import urlparse

# Boards whose company_url is a profile page rather than a company site.
# The bare domain is useless as an identity (every company shares it), so
# only domain + profile path is used for these.
AGGREGATOR_DOMAINS = {
    "linkedin.com", "indeed.com", "glassdoor.com", "google.com",
    "ziprecruiter.com", "naukri.com", "bayt.com", "bdjobs.com",
    # Dedicated remote-job boards (job_sources.py). None of these are a
    # company's own site, so - same as the boards above - a bare domain
    # match here must never stand in for a real company identity.
    "remoteok.com", "remotive.com", "weworkremotely.com", "jobspresso.co",
}

# Legal-entity suffixes stripped before comparing names. These mark the
# incorporation type, not the business, so "Acme Ltd" and "Acme Limited"
# are the same company. Stripping is exact-token only - no fuzziness.
LEGAL_SUFFIXES = {
    "inc", "incorporated", "llc", "ltd", "limited", "plc",
    "corp", "corporation", "gmbh", "ag", "bv", "nv",
    "sa", "sas", "srl", "spa", "ab", "oy", "kft", "pty", "pte",
    "pvt", "kk", "sarl", "sl", "aps",
}
# Deliberately NOT stripped: "group", "holdings", "company", "co",
# "private", "as". These read like suffixes but are part of the business
# name often enough that removing them merges distinct companies -
# "Delta Systems" and "Delta Systems Group Holdings" are not the same
# firm, and stripping both tokens would have made them identical.


def normalize_company_name(name) -> str:
    """Casing, accents, punctuation, spacing and legal suffixes removed.

    Returns "" for anything unusable, and the caller must treat an empty
    key as "no identity" rather than as a group everyone joins.
    """
    if name is None:
        return ""
    text = str(name)
    if text.strip().lower() in ("", "nan", "none"):
        return ""

    # Strip accents so "Zürich Tech" and "Zurich Tech" agree.
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))

    text = text.lower()
    # Ampersand is a real word in company names ("Johnson & Johnson").
    text = text.replace("&", " and ")
    # Punctuation to spaces, then collapse - this handles "Acme, Inc.",
    # "Acme  Inc" and "ACME-INC" alike.
    text = re.sub(r"[^a-z0-9]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    if not text:
        return ""

    tokens = text.split()
    # Strip trailing legal suffixes only. A leading or middle match may
    # be part of the actual name ("Company of Animals", "Group M").
    while len(tokens) > 1 and tokens[-1] in LEGAL_SUFFIXES:
        tokens.pop()

    return " ".join(tokens)


def _hostname(url) -> str:
    if url is None:
        return ""
    text = str(url).strip()
    if text.lower() in ("", "nan", "none"):
        return ""
    if "://" not in text:
        text = "https://" + text
    try:
        host = (urlparse(text).hostname or "").lower()
    except ValueError:
        return ""
    return host[4:] if host.startswith("www.") else host


def _registrable_domain(host: str) -> str:
    """Reduce a hostname to something stable enough to compare.

    Deliberately simple: the last two labels, plus a third for the common
    two-part public suffixes (co.uk, com.au). This is not a full public
    suffix list, and it does not need to be - it only has to be
    consistent, since it is compared against other rows from the same
    source, never parsed for meaning.
    """
    if not host:
        return ""
    parts = host.split(".")
    if len(parts) <= 2:
        return host
    two_part_tlds = {"co", "com", "org", "net", "gov", "ac", "edu"}
    if len(parts) >= 3 and parts[-2] in two_part_tlds and len(parts[-1]) == 2:
        return ".".join(parts[-3:])
    return ".".join(parts[-2:])


def extract_company_domain(company_url_direct) -> str:
    """The company's own website domain, or "" when the board did not
    return one. Never guessed from the company name."""
    host = _hostname(company_url_direct)
    if not host:
        return ""
    domain = _registrable_domain(host)
    # A company's own URL should not be an aggregator; if it is, the
    # field is unreliable for identity and is ignored here.
    return "" if domain in AGGREGATOR_DOMAINS else domain


def extract_profile_key(company_url) -> str:
    """An exact per-board company identifier, e.g.
    "linkedin.com/company/acme". Empty when the URL is missing or is a
    bare aggregator domain with no company path."""
    if company_url is None:
        return ""
    text = str(company_url).strip()
    if text.lower() in ("", "nan", "none"):
        return ""
    if "://" not in text:
        text = "https://" + text

    try:
        parsed = urlparse(text)
    except ValueError:
        return ""

    host = (parsed.hostname or "").lower()
    if host.startswith("www."):
        host = host[4:]
    if not host:
        return ""

    path = (parsed.path or "").strip("/").lower()
    domain = _registrable_domain(host)

    if domain in AGGREGATOR_DOMAINS:
        # Needs a company path to mean anything.
        if not path:
            return ""
        # linkedin.com/company/acme/jobs -> keep the identifying part
        segments = [s for s in path.split("/") if s][:2]
        return f"{domain}/{'/'.join(segments)}"

    # A non-aggregator company_url is itself a website; the domain alone
    # identifies it.
    return domain


def identity_keys(row) -> list:
    """All exact identifiers for one job row, most reliable first.

    Rows are grouped by shared keys, so returning several keys is what
    allows a website domain to link two spellings of the same name.
    """
    keys = []

    domain = extract_company_domain(row.get("company_url_direct"))
    if domain:
        keys.append(f"domain:{domain}")

    profile = extract_profile_key(row.get("company_url"))
    if profile:
        keys.append(f"profile:{profile}")

    name = normalize_company_name(row.get("company"))
    if name:
        keys.append(f"name:{name}")

    return keys


def assign_company_groups(records) -> list:
    """Return a group id per record; records sharing an id are one company.

    Domain and profile identifiers merge unconditionally - they are exact
    and unambiguous. Name matches are merged too, but only when the rows
    do not contradict each other: if two rows normalise to the same name
    yet advertise DIFFERENT company websites, they are treated as
    different companies and the name match is ignored.

    That guard is what stops a shared name from overriding hard evidence
    of two separate businesses.
    """
    union = UnionFind()

    for index, record in enumerate(records):
        node = f"row:{index}"
        union.find(node)

        domain = extract_company_domain(record.get("company_url_direct"))
        if domain:
            union.union(node, f"domain:{domain}")

        profile = extract_profile_key(record.get("company_url"))
        if profile:
            union.union(node, f"profile:{profile}")

    rows_by_name = {}
    for index, record in enumerate(records):
        name = normalize_company_name(record.get("company"))
        if name:
            rows_by_name.setdefault(name, []).append(index)

    for indexes in rows_by_name.values():
        domains = {
            extract_company_domain(records[i].get("company_url_direct"))
            for i in indexes
        }
        domains.discard("")
        if len(domains) > 1:
            # Same name, different websites - do not merge on the name.
            continue
        first = indexes[0]
        for other in indexes[1:]:
            union.union(f"row:{first}", f"row:{other}")

    return [union.find(f"row:{i}") for i in range(len(records))]


class UnionFind:
    """Groups rows that share any identifier.

    Needed because identity arrives in fragments: row A and row B may
    share a website, B and C may share a normalised name, and all three
    are then one company even though A and C have nothing in common
    directly.
    """

    def __init__(self):
        self.parent = {}

    def find(self, item):
        self.parent.setdefault(item, item)
        root = item
        while self.parent[root] != root:
            root = self.parent[root]
        # Path compression keeps repeated lookups cheap.
        while self.parent[item] != root:
            self.parent[item], item = root, self.parent[item]
        return root

    def union(self, a, b):
        root_a, root_b = self.find(a), self.find(b)
        if root_a != root_b:
            self.parent[root_b] = root_a
