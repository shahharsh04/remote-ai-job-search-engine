"""
HTTP API for the Remote AI Job Search Engine
============================================
A thin wrapper around the existing pipeline. It contains no search,
filtering, enrichment or export logic of its own - every request calls
`main.execute_pipeline()`, the same function the CLI uses.

Runs take minutes (job-board pagination plus careers-page lookups), far
longer than a browser will wait on one request. So a search starts in a
background thread and returns a run id immediately; the page polls
/api/status/{run_id} until it finishes, then downloads the workbook.

Run locally:
    uvicorn api:app --reload --port 8000

Deploy: any host that allows long-running processes (Render, Railway,
Fly.io, a VM). NOT Netlify Functions - see README for why.
"""

import logging
import os
import threading
import traceback
import uuid
from datetime import datetime
from pathlib import Path

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field

import main as engine
import pipeline
import icp
import excluded_company_store

logging.basicConfig(level=logging.INFO, format="%(message)s")
log = logging.getLogger("job_search")

app = FastAPI(title="Remote AI Job Search Engine", version="1.0")

# The frontend is hosted separately (Netlify), so its origin must be
# allowed explicitly. Set ALLOWED_ORIGINS to your Netlify URL in
# production; the default is permissive for local development only.
allowed_origins = [
    origin.strip()
    for origin in os.getenv("ALLOWED_ORIGINS", "*").split(",")
    if origin.strip()
]
app.add_middleware(
    CORSMiddleware,
    allow_origins=allowed_origins,
    allow_credentials=False,
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)

# Run state lives in memory: this is a single-worker tool, and a restart
# losing a run's status is acceptable. A database would be the answer if
# runs ever needed to survive a redeploy.
RUNS = {}
RUNS_LOCK = threading.Lock()
LAST_EXCLUDED_COMPANIES = []

VALID_REGIONS = engine.REGION_ORDER


class SearchRequest(BaseModel):
    job_title: str = Field(min_length=1, max_length=120)
    region: str


def _set(run_id: str, **fields) -> None:
    with RUNS_LOCK:
        if run_id in RUNS:
            RUNS[run_id].update(fields)


def _run_search(run_id: str, job_title: str, region: str) -> None:
    """Background worker: one pipeline run, then record the outcome."""
    try:
        _set(run_id, status="running", stage="Loading configuration")

        config = engine.load_config(os.getenv("CONFIG_PATH", "config.yaml"))
        lead_config = pipeline.load_lead_config(config)

        _set(run_id, stage="Searching job boards", target_leads=lead_config.get("target_leads", 50), current_leads=0, regions_searched=[])

        def progress(update):
            # Keep the browser informed while the long-running search
            # moves from one location to the next and then through the
            # lead-qualification stages.
            if isinstance(update, dict):
                _set(run_id, **update)

        result = engine.execute_pipeline(
            {"job_title": job_title, "region": region, "remote_only": True},
            config,
            lead_config,
            progress_callback=progress,
        )

        # The keywords actually searched, so the page can show why a job
        # it did not literally ask for is in the file.
        search_summary = result.get("summary") or {}
        expansion = search_summary.get("keyword_expansion") or {}

        ctx = result.get("ctx")
        excluded_watchlist = []
        if ctx is not None:
            excluded_watchlist = list(ctx.artifacts.get("excluded_company_watchlist", []))

        global LAST_EXCLUDED_COMPANIES
        with RUNS_LOCK:
            LAST_EXCLUDED_COMPANIES = excluded_watchlist
            if run_id in RUNS:
                RUNS[run_id]["excluded_company_watchlist"] = excluded_watchlist

        _set(
            run_id,
            status="completed",
            stage="Done",
            keywords_searched=expansion.get("keywords", [job_title]),
            similar_keywords=expansion.get("expanded", []),
            keyword_counts=search_summary.get("keyword_counts", {}),
            finished_at=datetime.now().isoformat(timespec="seconds"),
            output_path=result["output_path"],
            jobs_found=result["jobs_count"],
            unique_companies=result["unique_companies"],
            qualified_companies=result["qualified_companies"],
            excluded_companies=result["excluded_companies"],
            leads_selected=result["leads_count"],
            contacts_found=result["contacts_found"],
            contacts_not_found=result["contacts_not_found"],
            priority_counts=result["priority_counts"],
            target_leads=search_summary.get("lead_target", lead_config.get("target_leads", 50)),
            target_reached=search_summary.get("target_reached", False),
            regions_searched=search_summary.get("regions_searched", []),
            regions_skipped=search_summary.get("regions_skipped", []),
        )
    except Exception as exc:
        log.warning(f"[ERROR] run {run_id} failed: {exc}")
        traceback.print_exc()
        _set(
            run_id,
            status="failed",
            stage="Failed",
            error=f"{type(exc).__name__}: {exc}",
            finished_at=datetime.now().isoformat(timespec="seconds"),
        )


@app.api_route("/", methods=["GET", "HEAD"])
def root():
    # Render's default health check probes "/" unless a service is
    # configured (via render.yaml or the dashboard) to use a specific
    # path. Without this route, that probe - and every bot that scans
    # the bare URL - gets a 404, which is noisy and can make an
    # otherwise-healthy service look broken. This just points anyone
    # landing here at the real endpoints; it duplicates no pipeline logic.
    return {
        "service": "Remote AI Job Search Engine API",
        "status": "ok",
        "docs": "/docs",
        "health": "/api/health",
    }


@app.get("/api/health")
def health():
    # allowed_origins is echoed back so a CORS misconfiguration can be
    # diagnosed by opening this URL in a browser - no Render dashboard
    # access needed. These are frontend URLs, not secrets, so exposing
    # the list here is safe.
    return {
        "status": "ok",
        "regions": VALID_REGIONS,
        "allowed_origins": allowed_origins,
    }


@app.get("/api/exclusion-dictionary")
def exclusion_dictionary():
    """Return reusable rules and the permanently accumulated JSON company dictionary."""
    config_path = os.getenv("CONFIG_PATH", "config.yaml")
    try:
        config = engine.load_config(config_path)
        lead_config = pipeline.load_lead_config(config)
        dictionary_path = lead_config.get("excluded_company_dictionary_path", "data/excluded_companies.json")
    except Exception:
        dictionary_path = os.getenv("EXCLUDED_COMPANY_DICTIONARY_PATH", "data/excluded_companies.json")

    permanent = excluded_company_store.all_companies(dictionary_path)
    return {
        "categories": icp.exclusion_dictionary(),
        "persistent_excluded_companies": permanent,
        "persistent_company_count": len(permanent),
        "last_run_excluded_companies": list(LAST_EXCLUDED_COMPANIES),
        "dictionary_path": dictionary_path,
        "note": (
            "Companies in the persistent dictionary accumulate across runs. They are prospecting filters, "
            "not absolute claims about future buying behavior. A company is excluded when its identity is already "
            "in the dictionary or its current company data matches a configured exclusion rule."
        ),
    }


@app.get("/api/regions")
def regions():
    """Region list, so the frontend never hardcodes its own copy."""
    return {
        "regions": [
            {
                "name": name,
                "countries": [e["country_indeed"] for e in engine.REGION_DEFINITIONS[name]],
            }
            for name in VALID_REGIONS
        ]
    }


@app.post("/api/search")
def start_search(request: SearchRequest):
    job_title = request.job_title.strip()
    if not job_title:
        raise HTTPException(status_code=400, detail="Job title is required.")
    if request.region not in VALID_REGIONS:
        raise HTTPException(
            status_code=400,
            detail=f"Region must be one of: {', '.join(VALID_REGIONS)}",
        )

    run_id = uuid.uuid4().hex
    with RUNS_LOCK:
        RUNS[run_id] = {
            "run_id": run_id,
            "status": "queued",
            "stage": "Queued",
            "job_title": job_title,
            "region": request.region,
            "started_at": datetime.now().isoformat(timespec="seconds"),
        }

    threading.Thread(
        target=_run_search, args=(run_id, job_title, request.region), daemon=True
    ).start()

    return {"run_id": run_id, "status": "queued"}


@app.get("/api/status/{run_id}")
def status(run_id: str):
    with RUNS_LOCK:
        run = RUNS.get(run_id)
    if run is None:
        raise HTTPException(status_code=404, detail="Unknown run id.")

    # The absolute path is internal; the browser only needs to know a
    # file is available and what it is called.
    payload = {k: v for k, v in run.items() if k != "output_path"}
    if run.get("output_path"):
        payload["download_available"] = True
        payload["file_name"] = Path(run["output_path"]).name
    return payload


@app.get("/api/download/{run_id}")
def download(run_id: str):
    with RUNS_LOCK:
        run = RUNS.get(run_id)
    if run is None:
        raise HTTPException(status_code=404, detail="Unknown run id.")
    if run.get("status") != "completed":
        raise HTTPException(status_code=409, detail="This run has not finished yet.")

    path = Path(run.get("output_path", ""))
    if not path.is_file():
        raise HTTPException(status_code=404, detail="The generated file is no longer on disk.")

    return FileResponse(
        path,
        filename=path.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
