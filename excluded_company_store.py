"""Persistent JSON storage for low-probability prospect companies."""

import json
import logging
import os
import threading
from datetime import datetime, timezone
from pathlib import Path

log = logging.getLogger("job_search")
_LOCK = threading.RLock()


def _default_path() -> Path:
    return Path(os.getenv("EXCLUDED_COMPANY_DICTIONARY_PATH", "data/excluded_companies.json"))


def _normalise(value) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def _empty_store() -> dict:
    return {"version": 1, "companies": []}


def load(path: str | Path | None = None) -> dict:
    file_path = Path(path or _default_path())
    with _LOCK:
        if not file_path.exists():
            return _empty_store()
        try:
            data = json.loads(file_path.read_text(encoding="utf-8"))
            if not isinstance(data, dict):
                return _empty_store()
            companies = data.get("companies")
            if not isinstance(companies, list):
                data["companies"] = []
            data.setdefault("version", 1)
            return data
        except (OSError, json.JSONDecodeError) as exc:
            log.warning("[WARN] Could not read excluded-company dictionary %s: %s", file_path, exc)
            return _empty_store()


def _identity_keys(record: dict) -> set[str]:
    keys = set()
    for field in ("company_domain", "company_name"):
        value = _normalise(record.get(field))
        if value:
            keys.add(f"{field}:{value}")
    return keys


def contains(record: dict, path: str | Path | None = None) -> bool:
    wanted = _identity_keys(record)
    if not wanted:
        return False
    for item in load(path).get("companies", []):
        if wanted.intersection(_identity_keys(item)):
            return True
    return False


def _atomic_write(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = path.with_suffix(path.suffix + ".tmp")
    tmp_path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    tmp_path.replace(path)


def add(record: dict, category: str, reason: str, path: str | Path | None = None) -> dict | None:
    company_name = str(record.get("company_name") or record.get("company") or "").strip()
    domain = str(record.get("company_domain") or record.get("_company_domain") or "").strip()
    if not company_name and not domain:
        return None

    file_path = Path(path or _default_path())
    now = datetime.now(timezone.utc).isoformat(timespec="seconds")

    with _LOCK:
        data = load(file_path)
        companies = data.setdefault("companies", [])
        match = None
        incoming = {"company_name": company_name, "company_domain": domain}
        wanted = _identity_keys(incoming)
        for item in companies:
            if wanted.intersection(_identity_keys(item)):
                match = item
                break

        if match is None:
            match = {
                "company_name": company_name,
                "company_domain": domain,
                "category": category or "Unknown",
                "reason": reason or "Excluded by configured low-probability-buyer rule.",
                "first_seen": now,
                "last_seen": now,
                "times_seen": 1,
            }
            companies.append(match)
            changed = True
        else:
            match["last_seen"] = now
            match["times_seen"] = int(match.get("times_seen") or 0) + 1
            if company_name and not match.get("company_name"):
                match["company_name"] = company_name
            if domain and not match.get("company_domain"):
                match["company_domain"] = domain
            if category and not match.get("category"):
                match["category"] = category
            if reason and not match.get("reason"):
                match["reason"] = reason
            changed = True

        if changed:
            _atomic_write(file_path, data)
        return dict(match)


def all_companies(path: str | Path | None = None) -> list[dict]:
    companies = load(path).get("companies", [])
    return sorted(
        [dict(item) for item in companies if isinstance(item, dict)],
        key=lambda item: (
            _normalise(item.get("category")),
            _normalise(item.get("company_name")),
        ),
    )
