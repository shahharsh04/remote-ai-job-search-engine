# Lead Generation System — Business Overview

## Purpose
The system searches remote AI/ML job postings across the web, identifies the companies behind them, filters out companies unlikely to buy our services, and produces a ranked, contact-ready list of leads in a single Excel file.

---

## 1. How the Priority Score Is Calculated

Every qualified company is scored on a **0–100 scale** using six weighted factors ([lead_signals.py](lead_signals.py:317)):

| Factor | Max Points | What It Measures |
|---|---|---|
| ICP Fit (company size) | 30 | How closely the company's employee count matches our preferred customer size |
| Hiring Signal Strength | 25 | How strong the AI/ML hiring activity is (based on number of matching roles) |
| Contact Found | 15 | Whether a usable contact was found, and how confident we are in it |
| Multiple Openings | 10 | Whether the company has more than one matching job opening |
| Funding Signal | 10 | Whether the job posting itself mentions funding (e.g., "Series A") |
| New AI Initiative | 10 | Whether the company appears to be building or expanding an AI team, vs. already having a mature one |

**Rules for each factor:**
- **ICP Fit:** Full points for "Preferred ICP" size band; 60% for Medium; 25% for Large; 40% for Below-ICP; 50% if size is unknown (unknown is never treated as a bad sign).
- **Hiring Signal:** Full points for "Strong" signal, 60% for "Moderate," 20% for "Weak."
- **Contact Found:** 100% of points for High-confidence contact, 70% for Medium, 40% for Low, 0 if no contact found.
- **Multiple Openings:** Full points for 3+ matching openings, 60% for exactly 2, 0 for a single opening.
- **Funding Signal:** Full points if the job posting itself states funding language; 0 otherwise (no external funding research is performed).
- **New AI Initiative:** Full points if the company is "Building new AI capability," 60% if "Expanding AI team," 0 if it already has an "Established AI team," 30% if unclear.

All raw points are totaled and **normalized to a 0–100 score** (`100 × total points earned / total points possible`). Every reason contributing to the score is recorded and stored for auditability.

---

## 2. How Leads Are Classified: High / Medium / Low

Based on the normalized 0–100 score ([lead_signals.py:346-360](lead_signals.py:346)):

- **Score ≥ 70** → High Priority
- **Score ≥ 45** → Medium Priority
- **Score < 45** → Low Priority

**Business-facing adjustment:** To keep the delivered list action-oriented, the displayed priority tier is promoted one step before it reaches the final spreadsheet:
- High stays **High**
- Medium becomes **High**
- Low becomes **Medium**

(The underlying numeric score is left unchanged in the export for transparency — only the labeled tier is promoted.)

---

## 3. Filters and Qualification Criteria

Before a company is counted as a lead, it passes through several checks ([icp.py](icp.py), [pipeline.py](pipeline.py:586)):

**a) Remote-only check** — Every posting is re-verified as genuinely remote before it enters the pipeline.

**b) Competitor exclusion** — Companies are automatically excluded if their own name, industry, or description identifies them as one of the following (a direct competitor to our AI engineering/services offering), based on matched keyword phrases:
- Staffing / staff augmentation agencies
- Recruitment / executive search firms
- IT services / systems integrators
- AI/ML consulting firms
- Software outsourcing / development agencies
- BPO (business process outsourcing) providers

A company is also excluded if the job posting text itself indicates it was placed on behalf of a client (e.g., "our client is seeking…"), which marks the poster as a recruiting intermediary rather than the hiring company itself.

**c) Company size banding** — Companies are NOT excluded for size; instead they are banded and scored accordingly:
- **Preferred ICP:** 20–500 employees
- **Medium:** 500 employees up to the "low priority" threshold (5,000)
- **Large:** 5,000+ employees (kept, but scored lower)
- **Below ICP:** under 20 employees
- **Unknown:** kept and treated neutrally, never penalized as a bad fit

**d) Persistent exclusion memory** — Companies excluded in a previous run are remembered in a saved dictionary (`data/excluded_companies.json`) and automatically rejected in future runs without re-evaluating them.

**e) Missing data policy** — Any field the job boards did not supply (size, funding, contact, etc.) is reported as "Unknown" or "Not Found" rather than guessed or fabricated, and is never treated as disqualifying.

---

## 4. Data Sources / Websites Used

Job postings and company information are collected from the following platforms ([config.yaml](config.yaml:29), [job_sources.py](job_sources.py)):

| Source | Access Method |
|---|---|
| Indeed | Via JobSpy library |
| LinkedIn | Via JobSpy library |
| ZipRecruiter | Via JobSpy library |
| Glassdoor | Via JobSpy library |
| RemoteOK | Public JSON API (remoteok.com/api) |
| Remotive | Public JSON API (remotive.com/api/remote-jobs) |
| We Work Remotely | Public per-category RSS feeds |
| Jobspresso | Public RSS feed (jobspresso.co/jobs/feed) |

**Deliberately excluded** (no legitimate public/free access exists): Wellfound/AngelList Talent (requires authenticated login; blocked by robots.txt) and FlexJobs (subscription paywall, terms prohibit automated collection).

Only postings from the **last 7 days (168 hours)** are considered by default.

---

## 5. APIs Used

| Stage | API / Method |
|---|---|
| Job search & collection | JobSpy library (wraps Indeed, LinkedIn, ZipRecruiter, Glassdoor); RemoteOK public JSON API; Remotive public JSON API; We Work Remotely & Jobspresso RSS feeds |
| Company/industry enrichment | None external — all company size, industry, and description data comes directly from what the job boards already return with each posting |
| Contact extraction | No third-party contact-finder API is used. Contacts are extracted from: (1) emails already captured in the job posting text, (2) the company's own public careers page (fetched directly over HTTP, respecting robots.txt) |
| Contact "validation" / scoring | Rule-based scoring logic (not an external API) — ranks found emails by whether they sit on the company's own domain, contain recruiting-related wording (e.g., "careers@", "talent@"), or belong to an ATS vendor/compliance mailbox |
| LinkedIn contacts | Not scraped. The system generates **LinkedIn people-search links** (deep links a person can click), never an assumed or scraped profile |

**Note:** There is no funding-data API, no email-verification API, and no third-party enrichment service integrated — every data point is sourced only from what the job boards and the company's own public web pages provide.

---

## 6. Overall Pipeline Flow

```
1. Job Collection
   → Search all configured platforms (Indeed, LinkedIn, ZipRecruiter,
     Glassdoor, RemoteOK, Remotive, WWR, Jobspresso) for the given job
     title and location, with similar-title keyword expansion.

2. Remote Filtering
   → Re-verify every posting is genuinely remote.

3. Company Identification & Deduplication
   → Group postings by company (via website domain, board profile URL,
     or normalized name) so each company appears once, with all its
     matching job openings combined.

4. ICP Filtering (Qualification)
   → Exclude direct competitors (staffing/recruiting/IT-services/AI-
     consulting/outsourcing/BPO firms) and previously-excluded companies.
   → Band remaining companies by size (Preferred / Medium / Large /
     Below ICP / Unknown) — none of these are dropped.

5. Contact Enrichment
   → Extract contact emails from the job posting and the company's own
     careers page; generate LinkedIn search links; score and rank
     contacts by likely usefulness.

6. Lead Qualification (Signal Derivation)
   → Derive hiring-signal strength, AI-team maturity, funding mentions,
     and a plain-English lead summary for each company.

7. Lead Prioritization (Scoring)
   → Score every company 0–100 using the six weighted factors, assign
     High/Medium/Low priority, and rank all companies.

8. Final Selection & Excel Export
   → Cut the ranked list to the configured target number of leads
     (default: 50) and write the results to a single Excel workbook
     with these tabs:
       • Qualified Leads       — the final prioritized lead list
       • All Company Signals   — every scored company, for audit
       • Excluded Companies    — companies filtered out, with reasons
       • Pipeline Summary      — a stage-by-stage run report
       • Remote Jobs           — the underlying job-posting detail
```

---

## Key Principles Behind the Design

- **No fabricated data.** If a data point (size, funding, contact) isn't available from a source, it is reported as "Unknown"/"Not Found" — never guessed or invented.
- **Full auditability.** Every score carries the specific reasons behind it, and every exclusion carries the matched phrase and source field that caused it.
- **Config-driven, not hard-coded.** All weights, size thresholds, and priority cut-offs live in [config.yaml](config.yaml), so they can be retuned by a non-technical user without changing code.
