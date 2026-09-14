#!/usr/bin/env python3
"""Discover and select recent films and series for Nathan's cinema catalogue.

Films released in France are read from AlloCiné's public weekly agenda. A film
is accepted when:
- at least one principal actor is already associated with a good work in the
  catalogue and the AlloCiné press score is at least 2.5/5; or
- no actor is recognised yet and the press score is at least 3.5/5.

For series, where a comparable weekly French press score is not consistently
available, the equivalent IMDb thresholds are 5.0/10 and 7.0/10, with minimum
vote counts to avoid treating a tiny sample as a breakout.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import html as html_lib
import json
import re
import time
import unicodedata
import urllib.error
import urllib.request
from html.parser import HTMLParser
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
CONFIG_PATH = ROOT / "config" / "workbooks.json"
STATE_PATH = ROOT / "data" / "health" / "cinema-discovery.json"
ALLOCINE_AGENDA = "https://www.allocine.fr/film/agenda/"
ALLOCINE_WEEK = "https://www.allocine.fr/film/agenda/sem-{week}/"
USER_AGENT = (
    "NathanExcelCatalogues/1.1 "
    "(weekly public cinema metadata; "
    "https://github.com/Nathan20202/excel-catalogues-auto-update)"
)
FILM_LINK_RE = re.compile(
    r"""(?:https://www\.allocine\.fr)?/film/fichefilm_gen_cfilm=(\d+)\.html""",
    re.IGNORECASE,
)
BLOCK_TAGS = {
    "address", "article", "aside", "blockquote", "br", "div", "footer",
    "h1", "h2", "h3", "h4", "header", "li", "main", "nav", "p",
    "section", "td", "th", "tr",
}
IGNORED_TAGS = {"script", "style", "noscript", "svg"}
FRENCH_MONTHS = {
    "janvier": 1,
    "fevrier": 2,
    "mars": 3,
    "avril": 4,
    "mai": 5,
    "juin": 6,
    "juillet": 7,
    "aout": 8,
    "septembre": 9,
    "octobre": 10,
    "novembre": 11,
    "decembre": 12,
}


class VisibleTextParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []
        self.metas: dict[str, str] = {}
        self.ignored_depth = 0

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.casefold()
        if tag in IGNORED_TAGS:
            self.ignored_depth += 1
            return
        if self.ignored_depth:
            return
        if tag == "meta":
            values = {str(key).casefold(): value or "" for key, value in attrs}
            name = (values.get("property") or values.get("name") or "").casefold()
            content = html_lib.unescape(values.get("content", "")).strip()
            if name and content:
                self.metas[name] = content
        if tag in BLOCK_TAGS:
            self.parts.append("\n")

    def handle_endtag(self, tag: str) -> None:
        tag = tag.casefold()
        if tag in IGNORED_TAGS and self.ignored_depth:
            self.ignored_depth -= 1
            return
        if not self.ignored_depth and tag in BLOCK_TAGS:
            self.parts.append("\n")

    def handle_data(self, data: str) -> None:
        if not self.ignored_depth:
            self.parts.append(data)

    def visible_text(self) -> str:
        raw = html_lib.unescape("".join(self.parts))
        lines = [re.sub(r"\s+", " ", line).strip() for line in raw.splitlines()]
        return "\n".join(line for line in lines if line)


def now_utc() -> dt.datetime:
    return dt.datetime.now(dt.timezone.utc).replace(microsecond=0)


def iso_now() -> str:
    return now_utc().isoformat().replace("+00:00", "Z")


def normalize_key(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = re.sub(r"[^a-z0-9]+", " ", text.casefold())
    return re.sub(r"\s+", " ", text).strip()


def split_people(value: Any) -> list[str]:
    people = []
    for item in re.split(r"\s*[•|]\s*|\s*,\s*", str(value or "")):
        item = item.strip()
        if len(item) >= 3:
            people.append(item)
    return people


def as_list(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item).strip()]
    if isinstance(value, str):
        return [item.strip() for item in re.split(r"\s*[•,|]\s*", value) if item.strip()]
    return [str(value).strip()]


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


def fetch_html(url: str, timeout: int = 35) -> str | None:
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": USER_AGENT,
            "Accept": "text/html,application/xhtml+xml;q=0.9,*/*;q=0.5",
            "Accept-Language": "fr-FR,fr;q=0.9,en;q=0.6",
            "Cache-Control": "no-cache",
        },
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            charset = response.headers.get_content_charset() or "utf-8"
            return response.read(1_500_000).decode(charset, errors="replace")
    except Exception as exc:
        print(f"AlloCiné indisponible pour {url}: {type(exc).__name__}: {exc}")
        return None


def fetch_json(url: str, timeout: int = 30) -> dict[str, Any] | None:
    request = urllib.request.Request(
        url,
        headers={"User-Agent": USER_AGENT, "Accept": "application/json"},
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            return json.load(response)
    except Exception as exc:
        print(f"Métadonnées indisponibles pour {url}: {type(exc).__name__}: {exc}")
        return None


def parse_french_date(day: str, month: str, year: str) -> dt.date | None:
    month_number = FRENCH_MONTHS.get(normalize_key(month))
    if not month_number:
        return None
    try:
        return dt.date(int(year), month_number, int(day))
    except ValueError:
        return None


def line_value(text: str, label: str) -> str | None:
    match = re.search(
        rf"(?:^|\n){re.escape(label)}\s+([^\n]+)",
        text,
        re.IGNORECASE,
    )
    return match.group(1).strip() if match else None


def extract_agenda_links(document: str) -> dict[str, str]:
    links: dict[str, str] = {}
    for film_id in FILM_LINK_RE.findall(document):
        links[film_id] = (
            "https://www.allocine.fr/film/"
            f"fichefilm_gen_cfilm={film_id}.html"
        )
    return links


def parse_allocine_film(document: str, film_id: str, url: str) -> dict[str, Any] | None:
    parser = VisibleTextParser()
    try:
        parser.feed(document)
    except Exception:
        return None
    text = parser.visible_text()

    title = parser.metas.get("og:title", "").strip()
    title = re.sub(r"\s*-\s*Film\s+\d{4}.*$", "", title, flags=re.IGNORECASE)
    title = re.sub(r"\s*-\s*AlloCiné\s*$", "", title, flags=re.IGNORECASE)
    if not title:
        title_match = re.search(r"(?:^|\n)([^\n]+)\n.*?\ben salle\b", text)
        title = title_match.group(1).strip() if title_match else ""
    if not title:
        return None

    release_match = re.search(
        r"\b(\d{1,2})\s+([A-Za-zÀ-ÿ]+)\s+(\d{4})\s+en salle\b",
        text,
        re.IGNORECASE,
    )
    release_date = (
        parse_french_date(*release_match.groups()) if release_match else None
    )

    press_match = re.search(
        r"(?:^|\n)Presse\s*(?:\n|\s)+"
        r"([0-5](?:[,.]\d))\s+(\d+)\s+critiques?",
        text,
        re.IGNORECASE,
    )
    press_score = (
        float(press_match.group(1).replace(",", ".")) if press_match else None
    )
    press_reviews = int(press_match.group(2)) if press_match else 0

    release_line = ""
    if release_match:
        for line in text.splitlines():
            if release_match.group(0).casefold() in line.casefold():
                release_line = line
                break

    duration = None
    genres: list[str] = []
    if release_line:
        parts = [part.strip() for part in release_line.split("|")]
        for part in parts[1:]:
            if re.fullmatch(r"\d+h(?:\s*\d+min)?|\d+\s*min", part, re.IGNORECASE):
                duration = part
            elif part and "séance" not in normalize_key(part):
                genres.extend(
                    genre.strip() for genre in part.split(",") if genre.strip()
                )

    director = line_value(text, "De")
    cast_line = line_value(text, "Avec")
    actors = split_people(cast_line)
    original_title = line_value(text, "Titre original")

    meta_description = (
        parser.metas.get("description")
        or parser.metas.get("og:description")
        or ""
    )
    synopsis_match = re.search(
        r"\bSynopsis\s*:\s*(.+)$",
        meta_description,
        re.IGNORECASE | re.DOTALL,
    )
    synopsis = synopsis_match.group(1).strip() if synopsis_match else ""
    if not synopsis:
        lines = text.splitlines()
        for index, line in enumerate(lines):
            if normalize_key(line) == "synopsis":
                for candidate in lines[index + 1 : index + 8]:
                    if len(candidate) >= 40 and not candidate.lower().startswith("interdit"):
                        synopsis = candidate.strip()
                        break
                break

    return {
        "allocineId": film_id,
        "url": url,
        "title": title,
        "originalTitle": original_title,
        "releaseDate": release_date.isoformat() if release_date else None,
        "year": release_date.year if release_date else None,
        "duration": duration,
        "genres": genres,
        "director": director,
        "actors": actors,
        "synopsis": synopsis,
        "pressScore": press_score,
        "pressReviews": press_reviews,
    }


def cinema_datasets(config: dict[str, Any], sheet: str) -> list[dict[str, Any]]:
    return [
        dataset
        for dataset in config["workbooks"]["cinema"]["datasets"]
        if dataset.get("sheet") == sheet
    ]


def catalogue_index(
    config: dict[str, Any],
) -> tuple[set[str], set[tuple[str, int]], set[str], set[str]]:
    known_actors: set[str] = set()
    known_titles: set[tuple[str, int]] = set()
    known_imdb: set[str] = set()
    known_allocine: set[str] = set()

    for sheet in ("02 - Films", "03 - Séries"):
        for dataset in cinema_datasets(config, sheet):
            payload = load_json(ROOT / dataset["file"])
            for record in payload.get("records", []):
                title = normalize_key(record.get("Titre français / usuel"))
                try:
                    year = int(record.get("Année") or 0)
                except (TypeError, ValueError):
                    year = 0
                if title:
                    known_titles.add((title, year))

                imdb = str(record.get("_imdb_id", "")).casefold()
                if imdb:
                    known_imdb.add(imdb)
                allocine_id = str(record.get("_allocine_id", ""))
                if allocine_id:
                    known_allocine.add(allocine_id)

                try:
                    rating = float(record.get("Note IMDb") or 0)
                except (TypeError, ValueError):
                    rating = 0
                level = normalize_key(record.get("Niveau"))
                is_good_work = rating >= 6.0 or level.startswith(("1 ", "2 ", "3 ", "4 "))
                if not is_good_work:
                    continue
                for actor in split_people(record.get("Casting principal")):
                    key = normalize_key(actor)
                    if key:
                        known_actors.add(key)

    return known_actors, known_titles, known_imdb, known_allocine


def film_record(
    item: dict[str, Any],
    known_actor_names: list[str],
) -> dict[str, Any]:
    score = float(item["pressScore"])
    reviews = int(item["pressReviews"])
    year = int(item.get("year") or now_utc().year)

    if known_actor_names:
        reason = (
            f"Ajout automatique : note presse AlloCiné {score:.1f}/5 "
            f"({reviews} critiques) et acteur reconnu : "
            f"{', '.join(known_actor_names[:3])}."
        )
    else:
        reason = (
            f"Ajout automatique : nouveaux interprètes avec une forte réception, "
            f"note presse AlloCiné {score:.1f}/5 ({reviews} critiques)."
        )

    digest = hashlib.sha1(
        f"allocine-film|{item['allocineId']}".encode("utf-8")
    ).hexdigest()[:12].upper()

    return {
        "Titre français / usuel": item["title"],
        "Titre original": item.get("originalTitle"),
        "Année": year,
        "Époque": f"Années {(year // 10) * 10}",
        "Format": "Film",
        "Genres": " • ".join(item.get("genres") or []) or None,
        "Pays": None,
        "Réalisation / création": item.get("director"),
        "Casting principal": " • ".join(item.get("actors") or []) or None,
        "Durée / saisons": item.get("duration"),
        "Synopsis": item.get("synopsis") or "Synopsis à compléter.",
        "Pourquoi le voir": reason,
        "Niveau": "3 • Très recommandé" if score >= 3.5 else "4 • À voir",
        "Saga / univers": None,
        "Ordre": None,
        "Note IMDb": None,
        "Votes IMDb": None,
        "Distinctions": None,
        "Sortie": "Sorti",
        "Dans ta liste": None,
        "Source principale": item["url"],
        "ID": f"CIN-FILM-{digest}",
        "_allocine_id": str(item["allocineId"]),
        "_press_score": score,
        "_press_reviews": reviews,
        "_selection_policy": "known-actor-2.5" if known_actor_names else "new-cast-3.5",
    }


def refresh_allocine_films(config: dict[str, Any]) -> dict[str, Any]:
    state = load_json(
        STATE_PATH,
        {
            "schemaVersion": 1,
            "catalog": "cinema-discovery",
            "entries": {},
        },
    )
    entries: dict[str, dict[str, Any]] = state.setdefault("entries", {})
    known_actors, known_titles, _known_imdb, known_allocine = catalogue_index(config)

    today = now_utc().date()
    last_wednesday = today - dt.timedelta(days=(today.weekday() - 2) % 7)
    history_weeks = 12 if not entries else 5
    agenda_requests: list[tuple[str, dt.date | None]] = [(ALLOCINE_AGENDA, None)]
    for offset in range(history_weeks):
        week = last_wednesday - dt.timedelta(days=7 * offset)
        agenda_requests.append((ALLOCINE_WEEK.format(week=week.isoformat()), week))

    scheduled: dict[str, tuple[str, dt.date | None]] = {}
    failures = 0
    for agenda_url, expected_week in agenda_requests:
        document = fetch_html(agenda_url)
        if not document:
            failures += 1
            continue
        for film_id, url in extract_agenda_links(document).items():
            scheduled[film_id] = (url, expected_week)
        time.sleep(0.35)

    retry_cutoff = today - dt.timedelta(days=56)
    for film_id, entry in list(entries.items()):
        if entry.get("status") in {"accepted", "already-present"}:
            continue
        release_text = str(entry.get("releaseDate") or "")
        try:
            release_date = dt.date.fromisoformat(release_text)
        except ValueError:
            release_date = today
        if release_date >= retry_cutoff and entry.get("url"):
            scheduled.setdefault(film_id, (str(entry["url"]), release_date))

    accepted: list[dict[str, Any]] = []
    candidate_records: list[dict[str, Any]] = []
    checked = 0

    for film_id, (url, expected_week) in list(scheduled.items())[:160]:
        if film_id in known_allocine:
            entries[film_id] = {
                **entries.get(film_id, {}),
                "url": url,
                "status": "already-present",
                "checkedAt": iso_now(),
            }
            continue

        document = fetch_html(url)
        if not document:
            failures += 1
            entries[film_id] = {
                **entries.get(film_id, {}),
                "url": url,
                "status": "fetch-error",
                "checkedAt": iso_now(),
            }
            continue
        checked += 1
        item = parse_allocine_film(document, film_id, url)
        time.sleep(0.35)
        if not item:
            failures += 1
            entries[film_id] = {
                **entries.get(film_id, {}),
                "url": url,
                "status": "parse-error",
                "checkedAt": iso_now(),
            }
            continue

        release_text = str(item.get("releaseDate") or "")
        release_date = None
        try:
            release_date = dt.date.fromisoformat(release_text)
        except ValueError:
            pass
        if (
            expected_week is not None
            and release_date is not None
            and abs((release_date - expected_week).days) > 3
        ):
            continue

        title_key = (normalize_key(item["title"]), int(item.get("year") or 0))
        if title_key in known_titles:
            entries[film_id] = {
                **item,
                "status": "already-present",
                "checkedAt": iso_now(),
            }
            continue

        actor_lookup = {normalize_key(actor): actor for actor in item.get("actors", [])}
        known_hits = [
            original
            for key, original in actor_lookup.items()
            if key and key in known_actors
        ]
        score = item.get("pressScore")
        reviews = int(item.get("pressReviews") or 0)

        status = "pending"
        reason = "Note presse indisponible."
        if score is not None and reviews < 3:
            reason = "Moins de 3 critiques presse : attente d'un échantillon plus fiable."
        elif score is not None and known_hits and float(score) >= 2.5:
            status = "accepted"
            reason = "Acteur reconnu et note presse d'au moins 2,5/5."
        elif score is not None and not known_hits and reviews >= 5 and float(score) >= 3.5:
            status = "accepted"
            reason = "Nouveaux interprètes et note presse d'au moins 3,5/5."
        elif score is not None and known_hits:
            status = "rejected"
            reason = "Note presse inférieure à 2,5/5 malgré un acteur reconnu."
        elif score is not None:
            status = "rejected"
            reason = "Aucun acteur reconnu et note presse inférieure à 3,5/5."

        entries[film_id] = {
            **item,
            "knownActors": known_hits,
            "status": status,
            "reason": reason,
            "checkedAt": iso_now(),
        }

        if status == "accepted":
            record = film_record(item, known_hits)
            accepted.append(record)
            known_titles.add(title_key)
            known_allocine.add(film_id)
            for actor in item.get("actors", []):
                key = normalize_key(actor)
                if key:
                    known_actors.add(key)
        else:
            candidate_records.append(
                {
                    "source": "AlloCiné",
                    "kind": "film",
                    "title": item["title"],
                    "releaseDate": item.get("releaseDate"),
                    "actors": item.get("actors", []),
                    "knownActors": known_hits,
                    "pressScore": score,
                    "pressReviews": reviews,
                    "status": status,
                    "reason": reason,
                    "url": url,
                    "allocineId": film_id,
                }
            )

    if accepted:
        datasets = cinema_datasets(config, "02 - Films")
        target_dataset = datasets[-1]
        target_path = ROOT / target_dataset["file"]
        payload = load_json(target_path)
        payload["records"].extend(accepted)
        payload["recordCount"] = len(payload["records"])
        payload["updatedAt"] = iso_now()
        target_dataset["recordCount"] = payload["recordCount"]
        save_json(target_path, payload, compact=True)
        save_json(CONFIG_PATH, config)

    state.update(
        {
            "updatedAt": iso_now(),
            "checkedThisRun": checked,
            "addedThisRun": len(accepted),
            "pendingThisRun": sum(
                1 for record in candidate_records if record["status"] == "pending"
            ),
            "rejectedThisRun": sum(
                1 for record in candidate_records if record["status"] == "rejected"
            ),
            "failuresThisRun": failures,
            "criteria": {
                "knownActorPressMinimum": 2.5,
                "newCastPressMinimum": 3.5,
                "minimumPressReviews": 3,
                "minimumNewCastPressReviews": 5,
            },
        }
    )
    save_json(STATE_PATH, state)

    return {
        "added": len(accepted),
        "checked": checked,
        "failures": failures,
        "candidates": candidate_records,
    }


def series_record(
    candidate: dict[str, Any],
    meta: dict[str, Any],
    rating: float,
    votes: int,
    known_actor_names: list[str],
) -> dict[str, Any]:
    imdb = str(candidate["imdbId"]).casefold()
    actors = as_list(meta.get("cast"))[:5]
    genres = as_list(meta.get("genres"))
    directors = as_list(meta.get("director"))
    release_date = str(candidate.get("releaseDate") or "")
    year_match = re.search(r"\b(19|20)\d{2}\b", release_date)
    if not year_match:
        year_match = re.search(r"\b(19|20)\d{2}\b", str(meta.get("releaseInfo") or ""))
    year = int(year_match.group(0)) if year_match else now_utc().year

    if known_actor_names:
        reason = (
            f"Ajout automatique : acteur reconnu ({', '.join(known_actor_names[:3])}) "
            f"et note IMDb {rating:.1f}/10 ({votes:,} votes)."
        )
        policy = "known-actor-imdb-5.0"
    else:
        reason = (
            f"Ajout automatique : nouveaux interprètes en forte progression, "
            f"note IMDb {rating:.1f}/10 ({votes:,} votes)."
        )
        policy = "new-cast-imdb-7.0"

    digest = hashlib.sha1(f"series|{imdb}".encode("utf-8")).hexdigest()[:12].upper()
    return {
        "Titre français / usuel": meta.get("name") or candidate.get("title"),
        "Titre original": None,
        "Année": year,
        "Époque": f"Années {(year // 10) * 10}",
        "Format": "Série",
        "Genres": " • ".join(genres) or None,
        "Pays": meta.get("country"),
        "Réalisation / création": " • ".join(directors) or None,
        "Casting principal": " • ".join(actors) or None,
        "Durée / saisons": meta.get("runtime"),
        "Synopsis": meta.get("description") or "Synopsis à compléter.",
        "Pourquoi le voir": reason,
        "Niveau": "3 • Très recommandé" if rating >= 7.0 else "4 • À voir",
        "Saga / univers": None,
        "Ordre": None,
        "Note IMDb": rating,
        "Votes IMDb": votes,
        "Distinctions": None,
        "Sortie": "Sorti",
        "Dans ta liste": None,
        "Source principale": f"https://www.imdb.com/title/{imdb}/",
        "ID": f"CIN-SERIE-{digest}",
        "_imdb_id": imdb,
        "_selection_policy": policy,
    }


def add_recent_series(
    config: dict[str, Any],
    candidates: list[dict[str, Any]],
    ratings: dict[str, tuple[float, int]],
) -> dict[str, Any]:
    known_actors, known_titles, known_imdb, _known_allocine = catalogue_index(config)
    accepted: list[dict[str, Any]] = []
    pending: list[dict[str, Any]] = []
    checked = 0

    for candidate in candidates:
        if candidate.get("kind") != "series":
            continue
        imdb = str(candidate.get("imdbId") or "").casefold()
        if not imdb or imdb in known_imdb:
            continue
        rating, votes = ratings.get(imdb, (None, None))
        if rating is None or votes is None or rating < 5.0 or votes < 500:
            pending.append(
                {
                    **candidate,
                    "status": "pending",
                    "reason": "Réception IMDb encore insuffisante ou indisponible.",
                }
            )
            continue

        metadata = fetch_json(
            f"https://v3-cinemeta.strem.io/meta/series/{imdb}.json"
        )
        time.sleep(0.2)
        if not metadata or not isinstance(metadata.get("meta"), dict):
            pending.append(
                {
                    **candidate,
                    "imdbRating": rating,
                    "imdbVotes": votes,
                    "status": "pending",
                    "reason": "Métadonnées de distribution indisponibles.",
                }
            )
            continue

        checked += 1
        meta = metadata["meta"]
        actors = as_list(meta.get("cast"))[:5]
        known_hits = [
            actor for actor in actors if normalize_key(actor) in known_actors
        ]
        accepted_by_known_actor = bool(known_hits) and rating >= 5.0 and votes >= 500
        accepted_by_breakout = not known_hits and rating >= 7.0 and votes >= 5_000

        year_text = str(candidate.get("releaseDate") or meta.get("releaseInfo") or "")
        year_match = re.search(r"\b(19|20)\d{2}\b", year_text)
        year = int(year_match.group(0)) if year_match else 0
        title = meta.get("name") or candidate.get("title")
        title_key = (normalize_key(title), year)
        if title_key in known_titles:
            known_imdb.add(imdb)
            continue

        if accepted_by_known_actor or accepted_by_breakout:
            record = series_record(
                candidate,
                meta,
                float(rating),
                int(votes),
                known_hits,
            )
            accepted.append(record)
            known_imdb.add(imdb)
            known_titles.add(title_key)
            for actor in actors:
                key = normalize_key(actor)
                if key:
                    known_actors.add(key)
        else:
            pending.append(
                {
                    **candidate,
                    "actors": actors,
                    "knownActors": known_hits,
                    "imdbRating": rating,
                    "imdbVotes": votes,
                    "status": "rejected",
                    "reason": (
                        "Seuil IMDb non atteint : 5,0/10 avec acteur reconnu "
                        "ou 7,0/10 et 5 000 votes pour de nouveaux interprètes."
                    ),
                }
            )

        if checked >= 60:
            break

    if accepted:
        datasets = cinema_datasets(config, "03 - Séries")
        target_dataset = datasets[-1]
        target_path = ROOT / target_dataset["file"]
        payload = load_json(target_path)
        payload["records"].extend(accepted)
        payload["recordCount"] = len(payload["records"])
        payload["updatedAt"] = iso_now()
        target_dataset["recordCount"] = payload["recordCount"]
        save_json(target_path, payload, compact=True)
        save_json(CONFIG_PATH, config)

    return {
        "added": len(accepted),
        "checked": checked,
        "acceptedImdbIds": [
            str(record.get("_imdb_id", "")) for record in accepted
        ],
        "candidates": pending,
    }
