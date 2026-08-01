#!/usr/bin/env python3
"""Refresh public metadata conservatively without paid APIs or secrets."""

from __future__ import annotations

import argparse
import datetime as dt
import gzip
import hashlib
import html
import io
import json
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any, Iterable

ROOT = Path(__file__).resolve().parents[1]
CONFIG_PATH = ROOT / "config" / "workbooks.json"
MANIFEST_PATH = ROOT / "data" / "manifest.json"
HEALTH_DIR = ROOT / "data" / "health"
CANDIDATE_DIR = ROOT / "data" / "candidates"
USER_AGENT = (
    "NathanExcelCatalogues/1.0 "
    "(public metadata maintenance; https://github.com/Nathan20202/excel-catalogues-auto-update)"
)
URL_RE = re.compile(r"https?://[^\s<>\"]+", re.IGNORECASE)
TITLE_RE = re.compile(r"<title[^>]*>(.*?)</title>", re.IGNORECASE | re.DOTALL)
IMDB_RE = re.compile(r"^tt\d{7,9}$", re.IGNORECASE)

CHECK_LIMITS = {
    "promo": 100,
    "pokemon": 140,
    "fashion": 220,
    "cinema": 0,
    "gcdl": 120,
    "tech": 260,
    "activities": 240,
}


def now_utc() -> dt.datetime:
    return dt.datetime.now(dt.timezone.utc).replace(microsecond=0)


def iso_now() -> str:
    return now_utc().isoformat().replace("+00:00", "Z")


def french_date() -> str:
    return now_utc().strftime("%d/%m/%Y")


def load_json(path: Path, default: Any | None = None) -> Any:
    if not path.exists() and default is not None:
        return default
    return json.loads(path.read_text(encoding="utf-8"))


def save_json(path: Path, value: Any, *, compact: bool = False) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    if compact:
        text = json.dumps(value, ensure_ascii=False, separators=(",", ":")) + "\n"
    else:
        text = json.dumps(value, ensure_ascii=False, indent=2) + "\n"
    temporary.write_text(text, encoding="utf-8")
    temporary.replace(path)


def url_from_record(record: dict[str, Any]) -> str | None:
    preferred = (
        "Site officiel",
        "Lien officiel / vérification",
        "Lien",
        "URL",
        "Source officielle",
        "Source principale",
        "Site",
        "Preuve / source",
        "Source / preuve FR",
    )
    values: list[Any] = [record.get(key) for key in preferred]
    values.extend(record.values())
    for value in values:
        if not isinstance(value, str):
            continue
        match = URL_RE.search(value)
        if match:
            return match.group(0).rstrip(").,;]")
    return None


def fetch_probe(url: str) -> dict[str, Any]:
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": USER_AGENT,
            "Accept": "text/html,application/xhtml+xml,application/json,application/pdf;q=0.8,*/*;q=0.5",
            "Range": "bytes=0-131071",
        },
        method="GET",
    )
    try:
        with urllib.request.urlopen(request, timeout=18) as response:
            body = response.read(131072)
            status = int(getattr(response, "status", 200) or 200)
            final_url = response.geturl()
            content_type = response.headers.get("Content-Type", "")
            modified = response.headers.get("Last-Modified")
    except urllib.error.HTTPError as exc:
        body = exc.read(32768) if exc.fp else b""
        status = int(exc.code)
        final_url = exc.geturl() or url
        content_type = exc.headers.get("Content-Type", "") if exc.headers else ""
        modified = exc.headers.get("Last-Modified") if exc.headers else None
    except Exception as exc:
        return {
            "url": url,
            "ok": False,
            "status": 0,
            "checkedAt": iso_now(),
            "error": f"{type(exc).__name__}: {exc}"[:300],
        }

    text = body.decode("utf-8", errors="ignore")
    title_match = TITLE_RE.search(text)
    title = ""
    if title_match:
        title = re.sub(r"\s+", " ", html.unescape(title_match.group(1))).strip()[:240]
    digest = hashlib.sha256(body).hexdigest() if body else None
    return {
        "url": url,
        "finalUrl": final_url,
        "ok": 200 <= status < 400,
        "status": status,
        "checkedAt": iso_now(),
        "contentType": content_type,
        "lastModified": modified,
        "contentHash": digest,
        "title": title,
    }


def rotate(values: list[Any], cursor: int, limit: int) -> tuple[list[Any], int]:
    if not values or limit <= 0:
        return [], cursor
    limit = min(limit, len(values))
    selected = [values[(cursor + offset) % len(values)] for offset in range(limit)]
    return selected, (cursor + limit) % len(values)


def iter_catalog_datasets(config: dict[str, Any], catalog: str) -> Iterable[tuple[dict, Path, dict]]:
    workbook = config["workbooks"][catalog]
    for dataset in workbook["datasets"]:
        path = ROOT / dataset["file"]
        yield dataset, path, load_json(path)


def refresh_health(config: dict[str, Any], catalog: str) -> dict[str, Any]:
    state_path = HEALTH_DIR / f"{catalog}.json"
    state = load_json(
        state_path,
        {
            "schemaVersion": 1,
            "catalog": catalog,
            "cursor": 0,
            "entries": {},
            "changes": [],
        },
    )
    queue: list[tuple[dict, Path, dict, dict, str, str]] = []
    seen_urls: set[str] = set()
    for dataset, path, payload in iter_catalog_datasets(config, catalog):
        id_column = dataset["idColumn"]
        for record in payload["records"]:
            url = url_from_record(record)
            identifier = str(record.get(id_column, "")).strip()
            if not url or not identifier or url in seen_urls:
                continue
            seen_urls.add(url)
            queue.append((dataset, path, payload, record, identifier, url))

    selected, next_cursor = rotate(queue, int(state.get("cursor", 0)), CHECK_LIMITS[catalog])
    checked = 0
    ok = 0
    failures = 0
    changes: list[dict[str, Any]] = []
    touched_payloads: dict[Path, dict] = {}

    for dataset, path, payload, record, identifier, url in selected:
        probe = fetch_probe(url)
        checked += 1
        ok += int(bool(probe.get("ok")))
        failures += int(not probe.get("ok"))
        old = state.get("entries", {}).get(identifier, {})
        if (
            probe.get("ok")
            and old.get("contentHash")
            and probe.get("contentHash")
            and old["contentHash"] != probe["contentHash"]
        ):
            changes.append(
                {
                    "id": identifier,
                    "url": url,
                    "detectedAt": probe["checkedAt"],
                    "oldHash": old["contentHash"],
                    "newHash": probe["contentHash"],
                    "title": probe.get("title", ""),
                }
            )
        state.setdefault("entries", {})[identifier] = probe
        if probe.get("ok"):
            for header in ("Vérifié le", "Dernière vérification"):
                if header in record:
                    record[header] = french_date()
                    touched_payloads[path] = payload

    for path, payload in touched_payloads.items():
        payload["updatedAt"] = iso_now()
        payload["recordCount"] = len(payload["records"])
        save_json(path, payload, compact=True)

    state.update(
        {
            "updatedAt": iso_now(),
            "cursor": next_cursor,
            "totalUrls": len(queue),
            "checkedThisRun": checked,
            "okThisRun": ok,
            "failedThisRun": failures,
            "changes": (changes + state.get("changes", []))[:250],
        }
    )
    save_json(state_path, state)
    return state


def wikidata_recent_candidates() -> list[dict[str, Any]]:
    cutoff = (now_utc().date().replace(day=1) - dt.timedelta(days=760)).isoformat()
    query = f"""
SELECT ?item ?itemLabel ?imdb ?date ?kind WHERE {{
  VALUES ?kind {{ wd:Q11424 wd:Q5398426 }}
  ?item wdt:P31/wdt:P279* ?kind ;
        wdt:P345 ?imdb ;
        wdt:P577 ?date .
  FILTER(?date >= "{cutoff}T00:00:00Z"^^xsd:dateTime)
  SERVICE wikibase:label {{ bd:serviceParam wikibase:language "fr,en". }}
}}
ORDER BY DESC(?date)
LIMIT 300
"""
    endpoint = "https://query.wikidata.org/sparql?" + urllib.parse.urlencode(
        {"query": query, "format": "json"}
    )
    request = urllib.request.Request(
        endpoint,
        headers={"User-Agent": USER_AGENT, "Accept": "application/sparql-results+json"},
    )
    try:
        with urllib.request.urlopen(request, timeout=45) as response:
            payload = json.load(response)
    except Exception as exc:
        print(f"Wikidata indisponible: {exc}", file=sys.stderr)
        return []

    results: list[dict[str, Any]] = []
    for binding in payload.get("results", {}).get("bindings", []):
        imdb = binding.get("imdb", {}).get("value", "").lower()
        if not IMDB_RE.match(imdb):
            continue
        item_url = binding.get("item", {}).get("value", "")
        results.append(
            {
                "title": binding.get("itemLabel", {}).get("value", ""),
                "imdbId": imdb,
                "releaseDate": binding.get("date", {}).get("value", "")[:10],
                "kind": "film"
                if binding.get("kind", {}).get("value", "").endswith("Q11424")
                else "series",
                "wikidata": item_url,
            }
        )
    return results


def jikan_anime_candidates() -> list[dict[str, Any]]:
    url = "https://api.jikan.moe/v4/top/anime?filter=airing&limit=25"
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT, "Accept": "application/json"})
    try:
        with urllib.request.urlopen(request, timeout=35) as response:
            payload = json.load(response)
    except Exception as exc:
        print(f"Jikan indisponible: {exc}", file=sys.stderr)
        return []
    candidates = []
    for item in payload.get("data", []):
        candidates.append(
            {
                "malId": item.get("mal_id"),
                "title": item.get("title_english") or item.get("title"),
                "originalTitle": item.get("title"),
                "year": item.get("year"),
                "type": item.get("type"),
                "score": item.get("score"),
                "members": item.get("members"),
                "status": item.get("status"),
                "url": item.get("url"),
                "synopsis": item.get("synopsis"),
            }
        )
    return candidates


def update_imdb_ratings(
    config: dict[str, Any], external_candidates: list[dict[str, Any]]
) -> tuple[int, dict[str, tuple[float, int]]]:
    targets: set[str] = {
        candidate["imdbId"] for candidate in external_candidates if candidate.get("imdbId")
    }
    payloads: list[tuple[Path, dict]] = []
    for _dataset, path, payload in iter_catalog_datasets(config, "cinema"):
        payloads.append((path, payload))
        for record in payload["records"]:
            imdb = str(record.get("_imdb_id", "")).lower()
            if IMDB_RE.match(imdb):
                targets.add(imdb)
    if not targets:
        return 0, {}

    url = "https://datasets.imdbws.com/title.ratings.tsv.gz"
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    ratings: dict[str, tuple[float, int]] = {}
    try:
        with urllib.request.urlopen(request, timeout=90) as response:
            with gzip.GzipFile(fileobj=response) as compressed:
                text_stream = io.TextIOWrapper(compressed, encoding="utf-8")
                next(text_stream, None)
                for line in text_stream:
                    tconst, rating, votes = line.rstrip("\n").split("\t")
                    if tconst in targets:
                        ratings[tconst] = (float(rating), int(votes))
                        if len(ratings) == len(targets):
                            break
    except Exception as exc:
        print(f"IMDb indisponible: {exc}", file=sys.stderr)
        return 0, {}

    updated = 0
    for path, payload in payloads:
        touched = False
        for record in payload["records"]:
            imdb = str(record.get("_imdb_id", "")).lower()
            if imdb not in ratings:
                continue
            rating, votes = ratings[imdb]
            if "Note IMDb" in record and record.get("Note IMDb") != rating:
                record["Note IMDb"] = rating
                touched = True
            if "Votes IMDb" in record and record.get("Votes IMDb") != votes:
                record["Votes IMDb"] = votes
                touched = True
            updated += 1
        if touched:
            payload["updatedAt"] = iso_now()
            save_json(path, payload, compact=True)
    return updated, ratings


def refresh_cinema(config: dict[str, Any]) -> dict[str, Any]:
    candidates = wikidata_recent_candidates()
    updated, ratings = update_imdb_ratings(config, candidates)
    known: set[str] = set()
    for _dataset, _path, payload in iter_catalog_datasets(config, "cinema"):
        for record in payload["records"]:
            imdb = str(record.get("_imdb_id", "")).lower()
            if imdb:
                known.add(imdb)
    filtered = []
    for candidate in candidates:
        imdb = candidate["imdbId"]
        if imdb in known:
            continue
        rating, votes = ratings.get(imdb, (None, None))
        candidate["imdbRating"] = rating
        candidate["imdbVotes"] = votes
        candidate["reason"] = "Candidat récent issu de Wikidata, à enrichir avant intégration."
        if votes is None or votes >= 5_000:
            filtered.append(candidate)
    save_json(
        CANDIDATE_DIR / "cinema.json",
        {
            "schemaVersion": 1,
            "updatedAt": iso_now(),
            "policy": "Candidates are never injected automatically without enough metadata.",
            "count": len(filtered),
            "records": filtered[:250],
        },
    )
    anime = jikan_anime_candidates()
    save_json(
        CANDIDATE_DIR / "anime.json",
        {
            "schemaVersion": 1,
            "updatedAt": iso_now(),
            "source": "Jikan / MyAnimeList public API",
            "count": len(anime),
            "records": anime,
        },
    )
    state = {
        "schemaVersion": 1,
        "catalog": "cinema",
        "updatedAt": iso_now(),
        "ratingsMatched": updated,
        "cinemaCandidates": len(filtered),
        "animeCandidates": len(anime),
    }
    save_json(HEALTH_DIR / "cinema.json", state)
    return state


def write_change_candidates(catalog: str, state: dict[str, Any]) -> None:
    changes = state.get("changes", [])
    save_json(
        CANDIDATE_DIR / f"{catalog}-changes.json",
        {
            "schemaVersion": 1,
            "updatedAt": iso_now(),
            "catalog": catalog,
            "count": len(changes),
            "records": changes,
        },
    )


def touch_catalog_payloads(config: dict[str, Any], catalog: str) -> None:
    stamp = iso_now()
    for _dataset, path, payload in iter_catalog_datasets(config, catalog):
        payload["updatedAt"] = stamp
        payload["recordCount"] = len(payload["records"])
        save_json(path, payload, compact=True)


def update_manifest(catalogs: list[str], states: dict[str, dict[str, Any]]) -> None:
    manifest = load_json(MANIFEST_PATH)
    manifest["generatedAt"] = iso_now()
    for catalog in catalogs:
        workbook = manifest["workbooks"][catalog]
        workbook["lastAutomatedRun"] = states[catalog].get("updatedAt", iso_now())
        workbook["healthFile"] = f"data/health/{catalog}.json"
        workbook["candidateFile"] = (
            "data/candidates/cinema.json"
            if catalog == "cinema"
            else f"data/candidates/{catalog}-changes.json"
        )
    save_json(MANIFEST_PATH, manifest)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--catalog",
        required=True,
        choices=["promo", "pokemon", "fashion", "cinema", "gcdl", "all"],
    )
    args = parser.parse_args()
    config = load_json(CONFIG_PATH)
    catalogs = list(config["workbooks"]) if args.catalog == "all" else [args.catalog]
    states: dict[str, dict[str, Any]] = {}
    for catalog in catalogs:
        touch_catalog_payloads(config, catalog)
        if catalog == "cinema":
            state = refresh_cinema(config)
        else:
            state = refresh_health(config, catalog)
            write_change_candidates(catalog, state)
        states[catalog] = state
        print(
            f"{catalog}: actualisé — "
            f"{state.get('checkedThisRun', state.get('ratingsMatched', 0))} contrôles/métadonnées"
        )
    update_manifest(catalogs, states)


if __name__ == "__main__":
    main()
