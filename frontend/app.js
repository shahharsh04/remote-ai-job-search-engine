/*
 * Frontend for the Remote AI Job Search Engine.
 *
 * The API base URL comes from config.js, which Netlify regenerates at
 * build time from the API_BASE_URL environment variable. No URL is
 * hardcoded here.
 */

const API_BASE = (window.APP_CONFIG && window.APP_CONFIG.API_BASE_URL
  ? window.APP_CONFIG.API_BASE_URL
  : ""
).replace(/\/+$/, "");

// A run takes minutes, so the page polls rather than waiting on one
// long request that the browser would abandon.
const POLL_INTERVAL_MS = 3000;
const MAX_POLL_MINUTES = 20;

const FALLBACK_REGIONS = [
  { name: "USA", countries: ["USA"] },
  { name: "UK", countries: ["UK"] },
  { name: "Australia", countries: ["Australia"] },
  { name: "Europe", countries: ["Germany", "Netherlands", "Ireland", "France", "Spain"] },
];

const form = document.getElementById("search-form");
const jobTitleInput = document.getElementById("job-title");
const regionsBox = document.getElementById("regions");
const startButton = document.getElementById("start-button");
const statusCard = document.getElementById("status-card");
const statusTitle = document.getElementById("status-title");
const statusDetail = document.getElementById("status-detail");
const spinner = document.getElementById("spinner");
const results = document.getElementById("results");
const downloadButton = document.getElementById("download-button");
const apiTarget = document.getElementById("api-target");

let currentRunId = null;
let pollTimer = null;

function setStatus(state, title, detail) {
  statusCard.hidden = false;
  statusCard.classList.toggle("is-success", state === "success");
  statusCard.classList.toggle("is-error", state === "error");
  spinner.hidden = state !== "loading";
  statusTitle.textContent = title;
  statusDetail.textContent = detail || "";
}

function setBusy(busy) {
  startButton.disabled = busy;
  startButton.textContent = busy ? "Searching…" : "Start job search";
}

function renderRegions(regions) {
  regionsBox.innerHTML = regions
    .map((region, index) => {
      const countries =
        region.countries.length > 1 ? region.countries.join(", ") : "&nbsp;";
      return `
        <label class="region">
          <input type="radio" name="region" value="${region.name}" ${
        index === 0 ? "checked" : ""
      } />
          <span class="region-name">${region.name}</span>
          <span class="region-countries">${countries}</span>
        </label>`;
    })
    .join("");
}

async function loadRegions() {
  // The API owns the region list; this only falls back so the form still
  // works if the backend is briefly unreachable.
  try {
    const response = await fetch(`${API_BASE}/api/regions`);
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const data = await response.json();
    renderRegions(data.regions && data.regions.length ? data.regions : FALLBACK_REGIONS);
  } catch (error) {
    renderRegions(FALLBACK_REGIONS);
  }
}

function selectedRegion() {
  const checked = regionsBox.querySelector('input[name="region"]:checked');
  return checked ? checked.value : null;
}

async function readError(response) {
  try {
    const data = await response.json();
    if (typeof data.detail === "string") return data.detail;
    if (Array.isArray(data.detail) && data.detail.length) {
      return data.detail[0].msg || `Request failed (${response.status}).`;
    }
  } catch (error) {
    /* fall through to the generic message */
  }
  return `Request failed (${response.status}).`;
}

function stopPolling() {
  if (pollTimer) {
    clearInterval(pollTimer);
    pollTimer = null;
  }
}

function showResults(run) {
  document.getElementById("stat-jobs").textContent = run.jobs_found ?? "—";
  document.getElementById("stat-companies").textContent =
    run.qualified_companies ?? "—";
  document.getElementById("stat-leads").textContent = run.leads_selected ?? "—";
  document.getElementById("stat-contacts").textContent = run.contacts_found ?? "—";
  results.hidden = false;
}

function describeCounts(run) {
  const counts = run.priority_counts || {};
  const parts = [];
  if (counts.High) parts.push(`${counts.High} high`);
  if (counts.Medium) parts.push(`${counts.Medium} medium`);
  if (counts.Low) parts.push(`${counts.Low} low`);
  const priority = parts.length ? ` (${parts.join(", ")} priority)` : "";
  return `${run.unique_companies ?? 0} unique companies found, ${
    run.excluded_companies ?? 0
  } excluded${priority}.`;
}

async function pollStatus() {
  if (!currentRunId) return;

  let response;
  try {
    response = await fetch(`${API_BASE}/api/status/${currentRunId}`);
  } catch (error) {
    stopPolling();
    setBusy(false);
    setStatus("error", "Lost connection to the backend", "The search may still be running on the server.");
    return;
  }

  if (!response.ok) {
    stopPolling();
    setBusy(false);
    setStatus("error", "Could not read the run status", await readError(response));
    return;
  }

  const run = await response.json();

  if (run.status === "completed") {
    stopPolling();
    setBusy(false);
    showResults(run);
    setStatus("success", "Search complete", describeCounts(run));
    downloadButton.hidden = false;
    return;
  }

  if (run.status === "failed") {
    stopPolling();
    setBusy(false);
    setStatus("error", "The search failed", run.error || "No further detail was reported.");
    return;
  }

  setStatus("loading", run.stage || "Working…", "This can take several minutes.");
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  const jobTitle = jobTitleInput.value.trim();
  const region = selectedRegion();

  if (!jobTitle) {
    setStatus("error", "Job title is required", "Enter a job title such as “AI Engineer”.");
    jobTitleInput.focus();
    return;
  }
  if (!region) {
    setStatus("error", "Location is required", "Choose one of the four locations.");
    return;
  }
  if (!API_BASE) {
    setStatus(
      "error",
      "Backend URL is not configured",
      "Set the API_BASE_URL environment variable in Netlify and redeploy."
    );
    return;
  }

  stopPolling();
  results.hidden = true;
  downloadButton.hidden = true;
  setBusy(true);
  setStatus("loading", "Starting the search…", "");

  let response;
  try {
    response = await fetch(`${API_BASE}/api/search`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ job_title: jobTitle, region }),
    });
  } catch (error) {
    setBusy(false);
    setStatus(
      "error",
      "Could not reach the backend",
      `No response from ${API_BASE}. Check that the backend is running and that this site's URL is in its ALLOWED_ORIGINS.`
    );
    return;
  }

  if (!response.ok) {
    setBusy(false);
    setStatus("error", "The search could not be started", await readError(response));
    return;
  }

  const data = await response.json();
  currentRunId = data.run_id;
  setStatus("loading", "Search queued…", "This can take several minutes.");

  pollStatus();
  pollTimer = setInterval(pollStatus, POLL_INTERVAL_MS);

  // Stop polling rather than hammering the API forever if a run hangs.
  setTimeout(() => {
    if (pollTimer) {
      stopPolling();
      setBusy(false);
      setStatus(
        "error",
        "Stopped waiting for the search",
        `No result after ${MAX_POLL_MINUTES} minutes. The run may still be going on the server.`
      );
    }
  }, MAX_POLL_MINUTES * 60 * 1000);
});

downloadButton.addEventListener("click", () => {
  if (currentRunId) {
    window.location.href = `${API_BASE}/api/download/${currentRunId}`;
  }
});

apiTarget.textContent = API_BASE ? `Backend: ${API_BASE}` : "Backend URL not configured";
loadRegions();
