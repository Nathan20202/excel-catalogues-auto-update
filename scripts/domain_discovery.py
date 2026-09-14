#!/usr/bin/env python3
"""Discover trustworthy additions for every non-cinema catalogue.

The collectors deliberately prefer official or structured public sources. A row is
inserted only when its identity and a usable source URL are available. Incomplete
items remain in a candidate file and are retried on later runs.
"""

from __future__ import annotations

import datetime as dt
import email.utils
import hashlib
import html
import json
import math
import re
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, Callable

ROOT = Path(__file__).resolve().parents[1]
CONFIG_PATH = ROOT / "config" / "workbooks.json"
HEALTH_DIR = ROOT / "data" / "health"
USER_AGENT = (
    "NathanExcelCatalogues/2.0 "
    "(public catalogue maintenance; "
    "https://github.com/Nathan20202/excel-catalogues-auto-update)"
)
TODAY = dt.datetime.now(dt.timezone.utc).date()


def iso_now() -> str:
    return dt.datetime.now(dt.timezone.utc).replace(microsecond=0).isoformat().replace(
        "+00:00", "Z"
    )


def french_date() -> str:
    return TODAY.strftime("%d/%m/%Y")


def normalize(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(character for character in text if not unicodedata.combining(character))
    return re.sub(r"[^a-z0-9]+", " ", text.casefold()).strip()


def stable_id(prefix: str, value: str) -> str:
    digest = hashlib.sha1(value.encode("utf-8")).hexdigest()[:12].upper()
    return f"{prefix}-{digest}"


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


def request_bytes(
    url: str,
    *,
    data: bytes | None = None,
    timeout: int = 45,
    attempts: int = 3,
    headers: dict[str, str] | None = None,
) -> bytes:
    merged_headers = {
        "User-Agent": USER_AGENT,
        "Accept": "application/json,application/xml,text/xml,text/html,*/*;q=0.7",
    }
    if headers:
        merged_headers.update(headers)
    last_error: Exception | None = None
    for attempt in range(attempts):
        request = urllib.request.Request(
            url,
            data=data,
            headers=merged_headers,
            method="POST" if data is not None else "GET",
        )
        try:
            with urllib.request.urlopen(request, timeout=timeout) as response:
                return response.read(2_500_000)
        except Exception as exc:
            last_error = exc
            if attempt + 1 < attempts:
                time.sleep(1.2 * (attempt + 1))
    if last_error is not None:
        raise last_error
    raise RuntimeError(f"Réponse vide pour {url}")


def fetch_text(url: str, *, timeout: int = 40, attempts: int = 2) -> str:
    return request_bytes(url, timeout=timeout, attempts=attempts).decode(
        "utf-8", errors="ignore"
    )


def fetch_json(
    url: str,
    *,
    data: bytes | None = None,
    timeout: int = 45,
    attempts: int = 3,
) -> dict[str, Any]:
    return json.loads(
        request_bytes(
            url,
            data=data,
            timeout=timeout,
            attempts=attempts,
            headers={"Content-Type": "application/x-www-form-urlencoded"}
            if data is not None
            else None,
        ).decode("utf-8")
    )


def strip_html(value: Any) -> str:
    text = re.sub(r"<[^>]+>", " ", str(value or ""))
    return re.sub(r"\s+", " ", html.unescape(text)).strip()


def parse_date(value: Any) -> dt.date | None:
    text = str(value or "").strip()
    if not text:
        return None
    try:
        parsed = email.utils.parsedate_to_datetime(text)
        return parsed.date()
    except (TypeError, ValueError, OverflowError):
        pass
    clean = text.replace("Z", "+00:00")
    try:
        return dt.datetime.fromisoformat(clean).date()
    except ValueError:
        pass
    match = re.search(r"\b(20\d{2})-(\d{2})-(\d{2})\b", text)
    if match:
        try:
            return dt.date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
        except ValueError:
            return None
    return None


def child_value(node: ET.Element, names: tuple[str, ...]) -> str:
    for child in list(node):
        local = child.tag.rsplit("}", 1)[-1].casefold()
        if local not in names:
            continue
        if local == "link" and child.attrib.get("href"):
            return child.attrib["href"].strip()
        return "".join(child.itertext()).strip()
    return ""


def parse_feed(url: str) -> list[dict[str, Any]]:
    root = ET.fromstring(request_bytes(url, timeout=45, attempts=3))
    entries = [
        node
        for node in root.iter()
        if node.tag.rsplit("}", 1)[-1].casefold() in {"item", "entry"}
    ]
    results: list[dict[str, Any]] = []
    for node in entries:
        title = strip_html(child_value(node, ("title",)))
        link = child_value(node, ("link", "guid", "id"))
        published = child_value(
            node, ("pubdate", "published", "updated", "dc:date", "date")
        )
        summary = strip_html(
            child_value(node, ("description", "summary", "content", "encoded"))
        )
        if title and link.startswith("http"):
            results.append(
                {
                    "title": title[:300],
                    "url": link,
                    "published": published,
                    "date": parse_date(published),
                    "summary": summary[:1200],
                }
            )
    return results


def workbook_datasets(
    config: dict[str, Any], catalog: str, sheet: str | None = None
) -> list[dict[str, Any]]:
    datasets = config["workbooks"][catalog]["datasets"]
    if sheet is None:
        return list(datasets)
    return [dataset for dataset in datasets if dataset.get("sheet") == sheet]


def records_for(
    config: dict[str, Any], catalog: str, sheet: str | None = None
) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for dataset in workbook_datasets(config, catalog, sheet):
        payload = load_json(ROOT / dataset["file"])
        records.extend(payload.get("records", []))
    return records


def append_records(
    config: dict[str, Any],
    catalog: str,
    sheet: str,
    proposed: list[dict[str, Any]],
    *,
    identity: Callable[[dict[str, Any]], str],
    limit: int = 40,
) -> list[dict[str, Any]]:
    datasets = workbook_datasets(config, catalog, sheet)
    if not datasets or not proposed:
        return []
    known_ids: set[str] = set()
    known_identity: set[str] = set()
    for dataset in datasets:
        payload = load_json(ROOT / dataset["file"])
        for record in payload.get("records", []):
            identifier = str(record.get(dataset["idColumn"], "")).strip()
            if identifier:
                known_ids.add(identifier)
            key = identity(record)
            if key:
                known_identity.add(key)

    accepted: list[dict[str, Any]] = []
    for record in proposed:
        identifier = str(record.get(datasets[-1]["idColumn"], "")).strip()
        key = identity(record)
        if not identifier or identifier in known_ids or not key or key in known_identity:
            continue
        accepted.append(record)
        known_ids.add(identifier)
        known_identity.add(key)
        if len(accepted) >= limit:
            break
    if not accepted:
        return []

    target = datasets[-1]
    path = ROOT / target["file"]
    payload = load_json(path)
    payload["records"].extend(accepted)
    payload["recordCount"] = len(payload["records"])
    payload["updatedAt"] = iso_now()
    target["recordCount"] = payload["recordCount"]
    save_json(path, payload, compact=True)
    save_json(CONFIG_PATH, config)
    return accepted


def discovery_state(catalog: str) -> dict[str, Any]:
    return load_json(
        HEALTH_DIR / f"{catalog}-discovery.json",
        {"schemaVersion": 1, "catalog": f"{catalog}-discovery", "cursor": 0},
    )


def save_discovery_state(catalog: str, state: dict[str, Any]) -> None:
    state["schemaVersion"] = 1
    state["catalog"] = f"{catalog}-discovery"
    state["updatedAt"] = iso_now()
    save_json(HEALTH_DIR / f"{catalog}-discovery.json", state)


def wikidata_search(query: str, limit: int = 8) -> list[dict[str, Any]]:
    params = urllib.parse.urlencode(
        {
            "action": "wbsearchentities",
            "search": query,
            "language": "fr",
            "uselang": "fr",
            "type": "item",
            "limit": limit,
            "format": "json",
            "origin": "*",
        }
    )
    payload = fetch_json(f"https://www.wikidata.org/w/api.php?{params}")
    return list(payload.get("search", []))


def wikidata_entity(qid: str) -> dict[str, Any]:
    payload = fetch_json(f"https://www.wikidata.org/wiki/Special:EntityData/{qid}.json")
    return payload.get("entities", {}).get(qid, {})


def entity_label(entity: dict[str, Any], fallback: str) -> str:
    labels = entity.get("labels", {})
    for language in ("fr", "en"):
        value = labels.get(language, {}).get("value")
        if value:
            return str(value)
    return fallback


def entity_description(entity: dict[str, Any], fallback: str = "") -> str:
    descriptions = entity.get("descriptions", {})
    for language in ("fr", "en"):
        value = descriptions.get(language, {}).get("value")
        if value:
            return str(value)
    return fallback


def entity_websites(entity: dict[str, Any]) -> list[str]:
    results: list[str] = []
    for claim in entity.get("claims", {}).get("P856", []):
        value = (
            claim.get("mainsnak", {})
            .get("datavalue", {})
            .get("value")
        )
        if isinstance(value, str) and value.startswith("http"):
            results.append(value)
    return results


def discover_promo(config: dict[str, Any]) -> dict[str, Any]:
    queries = (
        "comparateur de prix",
        "price comparison website",
        "service cashback",
        "cashback website",
        "site de codes promotionnels",
        "coupon website",
    )
    state = discovery_state("promo")
    cursor = int(state.get("cursor", 0)) % len(queries)
    selected = [queries[cursor], queries[(cursor + 1) % len(queries)]]
    state["cursor"] = (cursor + 2) % len(queries)
    failures = 0
    checked = 0
    candidates: list[dict[str, Any]] = []
    comparison_rows: list[dict[str, Any]] = []
    promo_rows: list[dict[str, Any]] = []

    seen_qids: set[str] = set()
    for query in selected:
        try:
            hits = wikidata_search(query)
        except Exception as exc:
            failures += 1
            candidates.append({"source": "Wikidata", "query": query, "error": str(exc)[:240]})
            continue
        for hit in hits:
            qid = str(hit.get("id", ""))
            if not qid or qid in seen_qids:
                continue
            seen_qids.add(qid)
            try:
                entity = wikidata_entity(qid)
            except Exception:
                failures += 1
                continue
            checked += 1
            label = entity_label(entity, str(hit.get("label", "")))
            description = entity_description(entity, str(hit.get("description", "")))
            combined = normalize(f"{label} {description} {query}")
            relevant = any(
                token in combined
                for token in (
                    "comparateur prix",
                    "price comparison",
                    "cashback",
                    "code promotionnel",
                    "coupon website",
                    "coupon service",
                )
            )
            websites = entity_websites(entity)
            if not relevant or not websites:
                candidates.append(
                    {
                        "source": "Wikidata",
                        "qid": qid,
                        "title": label,
                        "description": description,
                        "status": "pending",
                        "reason": "Site officiel ou catégorie suffisamment précise indisponible.",
                    }
                )
                continue
            website = websites[0]
            source = f"https://www.wikidata.org/wiki/{qid}"
            if "compar" in combined:
                comparison_rows.append(
                    {
                        "Service": label,
                        "Portée": "Service découvert automatiquement — portée à confirmer",
                        "Zone / langue": description or "À vérifier",
                        "Catégories": "Comparaison de prix",
                        "Volume annoncé": "Non publié",
                        "Actualisation": "Selon le service",
                        "Garantie fraîcheur": "À vérifier sur le site officiel",
                        "Historique prix": "À vérifier",
                        "Alertes": "À vérifier",
                        "Frais de port": "À vérifier au panier",
                        "Classement / monétisation": "À vérifier dans les conditions du service",
                        "Idéal pour": "Compléter une comparaison avant achat",
                        "Limites": "Nouvelle entrée : contrôler couverture française et modèle économique.",
                        "Couverture /5": None,
                        "Fraîcheur /5": None,
                        "Fiabilité /5": None,
                        "Fonctions /5": None,
                        "France /5": None,
                        "Site": website,
                        "Source": source,
                        "Vérifié le": french_date(),
                        "ID": stable_id("COMP-AUTO", qid),
                        "_wikidata_id": qid,
                    }
                )
            else:
                promo_rows.append(
                    {
                        "Service": label,
                        "Famille": "Cashback / coupons — nouvelle entrée",
                        "Zone / langue": description or "À vérifier",
                        "Accès": "Site officiel",
                        "Actualisation": "Selon le service",
                        "Taille / couverture": "À confirmer",
                        "Codes testés": "À confirmer",
                        "Cashback": "À confirmer",
                        "Bons d’achat": "À confirmer",
                        "Extension": "À confirmer",
                        "Idéal pour": "Comparer une économie potentielle avant achat",
                        "Limites / vigilance": "Nouvelle entrée : vérifier conditions, paiement et réputation.",
                        "Couverture /5": None,
                        "Fraîcheur /5": None,
                        "Fiabilité /5": None,
                        "Fonctions /5": None,
                        "France /5": None,
                        "Site officiel": website,
                        "Preuve / source": source,
                        "Vérifié le": french_date(),
                        "ID": stable_id("PROMO-AUTO", qid),
                        "_wikidata_id": qid,
                    }
                )

    accepted_comparison = append_records(
        config,
        "promo",
        "04 - Comparateurs",
        comparison_rows,
        identity=lambda row: normalize(row.get("Service")),
        limit=8,
    )
    accepted_promo = append_records(
        config,
        "promo",
        "02 - Codes & cashback",
        promo_rows,
        identity=lambda row: normalize(row.get("Service")),
        limit=8,
    )
    state.update(
        {
            "checkedThisRun": checked,
            "addedThisRun": len(accepted_comparison) + len(accepted_promo),
            "failuresThisRun": failures,
            "queriesThisRun": selected,
        }
    )
    save_discovery_state("promo", state)
    return {
        "added": len(accepted_comparison) + len(accepted_promo),
        "checked": checked,
        "failures": failures,
        "candidates": candidates,
    }


POKEMON_AREAS = (
    ("Île-de-France", "FR", "France", 48.8566, 2.3522, 90000),
    ("Hauts-de-France", "FR", "France", 50.6292, 3.0573, 80000),
    ("Auvergne-Rhône-Alpes", "FR", "France", 45.7640, 4.8357, 90000),
    ("Provence-Alpes-Côte d’Azur", "FR", "France", 43.2965, 5.3698, 90000),
    ("Occitanie", "FR", "France", 43.6047, 1.4442, 100000),
    ("Nouvelle-Aquitaine", "FR", "France", 44.8378, -0.5792, 100000),
    ("Grand Est", "FR", "France", 48.5734, 7.7521, 90000),
    ("Pays de la Loire", "FR", "France", 47.2184, -1.5536, 90000),
    ("Belgique", "BE", "Belgique", 50.8503, 4.3517, 100000),
    ("Suisse romande", "CH", "Suisse", 46.2044, 6.1432, 100000),
    ("Luxembourg", "LU", "Luxembourg", 49.6116, 6.1319, 70000),
    ("Rhénanie", "DE", "Allemagne", 50.9375, 6.9603, 100000),
    ("Catalogne", "ES", "Espagne", 41.3874, 2.1686, 100000),
    ("Italie du Nord", "IT", "Italie", 45.4642, 9.1900, 100000),
    ("Pays-Bas", "NL", "Pays-Bas", 52.3676, 4.9041, 100000),
)


def overpass(query: str) -> dict[str, Any]:
    body = urllib.parse.urlencode({"data": query}).encode("utf-8")
    errors: list[str] = []
    for endpoint in (
        "https://overpass-api.de/api/interpreter",
        "https://overpass.kumi.systems/api/interpreter",
    ):
        try:
            return fetch_json(endpoint, data=body, timeout=55, attempts=1)
        except Exception as exc:
            errors.append(f"{endpoint}: {exc}")
    raise RuntimeError(" ; ".join(errors))


def discover_pokemon(config: dict[str, Any]) -> dict[str, Any]:
    state = discovery_state("pokemon")
    cursor = int(state.get("cursor", 0)) % len(POKEMON_AREAS)
    area_name, country_code, country_name, latitude, longitude, radius = (
        POKEMON_AREAS[cursor]
    )
    state["cursor"] = (cursor + 1) % len(POKEMON_AREAS)
    failures = 0
    checked = 0
    proposed: list[dict[str, Any]] = []
    candidates: list[dict[str, Any]] = []

    query = f"""
[out:json][timeout:45];
(
  nwr(around:{radius},{latitude},{longitude})["name"~"(pokemon|pokémon|tcg|trading cards?)",i]["shop"];
  nwr(around:{radius},{latitude},{longitude})["name"~"(pokemon|pokémon|tcg|trading cards?)",i]["amenity"];
  nwr(around:{radius},{latitude},{longitude})["shop"~"^(games|collector|toys|gift)$"]["name"~"(pokemon|pokémon|tcg|trading cards?)",i];
);
out center tags 80;
"""
    try:
        payload = overpass(query)
    except Exception as exc:
        failures = 1
        payload = {"elements": []}
        candidates.append(
            {
                "source": "OpenStreetMap / Overpass",
                "area": area_name,
                "status": "source-error",
                "reason": str(exc)[:300],
            }
        )

    for element in payload.get("elements", []):
        tags = element.get("tags", {})
        name = str(tags.get("name") or "").strip()
        website = str(
            tags.get("website")
            or tags.get("contact:website")
            or tags.get("url")
            or ""
        ).strip()
        if not name or not website.startswith("http"):
            continue
        checked += 1
        osm_type = str(element.get("type", "node"))
        osm_id = str(element.get("id", ""))
        osm_url = f"https://www.openstreetmap.org/{osm_type}/{osm_id}"
        page = ""
        try:
            page = fetch_text(website, timeout=18, attempts=1)[:350_000]
        except Exception:
            pass
        page_key = normalize(page)
        name_key = normalize(name)
        proves_pokemon = any(
            token in f"{name_key} {page_key}"
            for token in ("pokemon", "tcg", "trading card")
        )
        proves_french = country_code == "FR" or any(
            token in page_key
            for token in (
                "version francaise",
                "cartes francaises",
                "francais",
                "livraison france",
            )
        )
        if not proves_pokemon or not proves_french:
            candidates.append(
                {
                    "source": "OpenStreetMap + site de l’enseigne",
                    "title": name,
                    "country": country_name,
                    "website": website,
                    "status": "pending",
                    "reason": "Preuve Pokémon ou disponibilité française insuffisante.",
                }
            )
            continue

        city = str(
            tags.get("addr:city")
            or tags.get("addr:place")
            or tags.get("addr:suburb")
            or area_name
        )
        proposed.append(
            {
                "Enseigne / site": name,
                "Pays": country_name,
                "Ville / région": city,
                "Canal": "Magasin + en ligne",
                "Type": "Spécialiste TCG / boutique découverte",
                "Pérennité": "À évaluer",
                "Priorité": "C — complément",
                "Statut du français": "FR confirmé sur la source",
                "Offre principale": "Pokémon TCG ; assortiment exact à vérifier",
                "Scellé FR": "À vérifier",
                "À l'unité FR": "À vérifier",
                "Gradées / vintage": "À vérifier",
                "Précommandes": "À vérifier",
                "Accessoires": "À vérifier",
                "Rachat / revente": "À vérifier",
                "Tournois / League": "À vérifier",
                "Livraison France": "À vérifier au panier",
                "Douane / TVA": "Pas de douane si expédition depuis l’UE",
                "Retrait magasin": "Oui / à confirmer",
                "Position prix": "À comparer",
                "Idéal pour": "Élargir les recherches de stock et de précommandes Pokémon",
                "Points de vigilance": "Nouvelle enseigne : vérifier avis, langue, stock, frais et conditions avant paiement.",
                "Confiance structurelle /5": None,
                "Clarté de la preuve FR /5": 4,
                "Site officiel": website,
                "Source / preuve FR": osm_url,
                "Vérifié le": french_date(),
                "ID": stable_id("PKM-AUTO", f"{osm_type}:{osm_id}"),
                "_osm_id": f"{osm_type}/{osm_id}",
            }
        )

    accepted = append_records(
        config,
        "pokemon",
        "02 - Annuaire",
        proposed,
        identity=lambda row: normalize(row.get("Enseigne / site")),
        limit=20,
    )
    scanned = set(state.get("scannedAreas", []))
    if not failures:
        scanned.add(area_name)
    state.update(
        {
            "checkedThisRun": checked,
            "addedThisRun": len(accepted),
            "failuresThisRun": failures,
            "areaThisRun": area_name,
            "scannedAreas": sorted(scanned),
        }
    )
    save_discovery_state("pokemon", state)
    return {
        "added": len(accepted),
        "checked": checked,
        "failures": failures,
        "candidates": candidates,
    }

FASHION_QUERIES = (
    "marque de mode durable",
    "marque de vêtements homme",
    "fashion label",
    "streetwear brand",
    "menswear brand",
    "sustainable fashion brand",
    "footwear brand",
    "outdoor clothing brand",
)


def discover_fashion(config: dict[str, Any]) -> dict[str, Any]:
    state = discovery_state("fashion")
    cursor = int(state.get("cursor", 0)) % len(FASHION_QUERIES)
    selected = [
        FASHION_QUERIES[cursor],
        FASHION_QUERIES[(cursor + 1) % len(FASHION_QUERIES)],
    ]
    state["cursor"] = (cursor + 2) % len(FASHION_QUERIES)
    failures = 0
    checked = 0
    proposed: list[dict[str, Any]] = []
    candidates: list[dict[str, Any]] = []
    seen: set[str] = set()

    for query in selected:
        try:
            hits = wikidata_search(query, limit=12)
        except Exception as exc:
            failures += 1
            candidates.append({"source": "Wikidata", "query": query, "error": str(exc)[:240]})
            continue
        for hit in hits:
            qid = str(hit.get("id", ""))
            if not qid or qid in seen:
                continue
            seen.add(qid)
            try:
                entity = wikidata_entity(qid)
            except Exception:
                failures += 1
                continue
            checked += 1
            brand = entity_label(entity, str(hit.get("label", "")))
            description = entity_description(entity, str(hit.get("description", "")))
            combined = normalize(f"{brand} {description} {query}")
            websites = entity_websites(entity)
            relevant = any(
                token in combined
                for token in (
                    "mode",
                    "vetement",
                    "fashion",
                    "clothing",
                    "streetwear",
                    "footwear",
                    "chaussure",
                    "apparel",
                )
            )
            if not relevant or not websites:
                candidates.append(
                    {
                        "source": "Wikidata",
                        "qid": qid,
                        "title": brand,
                        "description": description,
                        "status": "pending",
                        "reason": "Marque ou site officiel insuffisamment établi.",
                    }
                )
                continue
            website = websites[0]
            proposed.append(
                {
                    "Marque": brand,
                    "Type d’entrée": "Marque / maison",
                    "Niveau de gamme": "À évaluer",
                    "Repère prix": "À relever sur le site officiel",
                    "Budget min (€)": None,
                    "Budget max (€)": None,
                    "Priorité": "4 • Nouveauté à évaluer",
                    "Style principal": "À classifier",
                    "Styles associés": description or "À classifier",
                    "Catégories fortes": "À confirmer",
                    "Pièces à acheter en priorité": "À déterminer après analyse de la collection",
                    "À choisir pour": "Découvrir une marque absente du catalogue",
                    "Points de vigilance": "Nouvelle entrée : vérifier matières, retours, production et rapport qualité-prix.",
                    "Public": "À vérifier",
                    "Pays / ADN": description or "À vérifier",
                    "Coupe / fit": "Consulter le guide des tailles officiel.",
                    "Matières / savoir-faire": "À vérifier sur les fiches produits.",
                    "Qualité /5": None,
                    "Rapport Q/P /5": None,
                    "Créativité /5": None,
                    "Transparence /5": None,
                    "Exclusivité /5": None,
                    "Canal conseillé": "Site officiel / revendeur agréé",
                    "Seconde main": "À vérifier",
                    "Maison / groupe": "À vérifier",
                    "Disponibilité": "Site officiel actif ; livraison France à confirmer",
                    "Lien officiel / vérification": website,
                    "Dernière vérification": french_date(),
                    "ID": stable_id("BR-AUTO", qid),
                    "_wikidata_id": qid,
                }
            )

    accepted = append_records(
        config,
        "fashion",
        "02 - Marques",
        proposed,
        identity=lambda row: normalize(row.get("Marque")),
        limit=16,
    )
    state.update(
        {
            "checkedThisRun": checked,
            "addedThisRun": len(accepted),
            "failuresThisRun": failures,
            "queriesThisRun": selected,
        }
    )
    save_discovery_state("fashion", state)
    return {
        "added": len(accepted),
        "checked": checked,
        "failures": failures,
        "candidates": candidates,
    }


ACTIVITY_AREAS = (
    ("Cabourg", 49.2918, -0.1130, 25000, "CAB"),
    ("Paris", 48.8566, 2.3522, 25000, "PARIS"),
    ("Orsay", 48.6992, 2.1875, 25000, "ORS"),
    ("Mende", 44.5180, 3.5006, 40000, "MEN"),
    ("Saint-Cyprien", 42.6182, 3.0067, 30000, "STC"),
    ("Boulouris", 43.4154, 6.8066, 30000, "BOU"),
)


def haversine(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    radius = 6371.0
    p1 = math.radians(lat1)
    p2 = math.radians(lat2)
    delta_p = math.radians(lat2 - lat1)
    delta_l = math.radians(lon2 - lon1)
    value = (
        math.sin(delta_p / 2) ** 2
        + math.cos(p1) * math.cos(p2) * math.sin(delta_l / 2) ** 2
    )
    return 2 * radius * math.asin(math.sqrt(value))


def activity_category(tags: dict[str, Any]) -> tuple[str, str, str, str]:
    tourism = str(tags.get("tourism", ""))
    leisure = str(tags.get("leisure", ""))
    historic = str(tags.get("historic", ""))
    if tourism in {"museum", "gallery"}:
        return "Musée & exposition", "Musée / galerie", "Intérieur", "Idéal pluie / chaleur"
    if tourism in {"zoo", "aquarium"}:
        return "Nature & animaux", tourism.capitalize(), "Mixte", "Tous temps"
    if tourism == "theme_park" or leisure in {"water_park", "amusement_arcade"}:
        return "Loisirs", "Parc / divertissement", "Mixte", "Temps sec de préférence"
    if tourism == "viewpoint":
        return "Nature & panorama", "Point de vue", "Extérieur", "Temps clair"
    if leisure in {"sports_centre", "escape_game", "bowling_alley"}:
        return "Sport & loisirs", leisure.replace("_", " ").title(), "Intérieur", "Tous temps"
    if leisure in {"park", "garden", "marina"}:
        return "Nature & plein air", leisure.replace("_", " ").title(), "Extérieur", "Temps sec"
    if historic:
        return "Patrimoine & culture", "Site historique", "Mixte", "Tous temps"
    return "Découverte locale", "Lieu d’intérêt", "Mixte", "Tous temps"


def discover_activities(config: dict[str, Any]) -> dict[str, Any]:
    state = discovery_state("activities")
    cursor = int(state.get("cursor", 0)) % len(ACTIVITY_AREAS)
    selected = [
        ACTIVITY_AREAS[cursor],
        ACTIVITY_AREAS[(cursor + 1) % len(ACTIVITY_AREAS)],
    ]
    state["cursor"] = (cursor + 2) % len(ACTIVITY_AREAS)
    failures = 0
    checked = 0
    added = 0
    candidates: list[dict[str, Any]] = []

    for sheet, center_lat, center_lon, radius, prefix in selected:
        query = f"""
[out:json][timeout:45];
(
  nwr(around:{radius},{center_lat},{center_lon})["name"]["tourism"~"^(attraction|museum|gallery|theme_park|zoo|aquarium|viewpoint)$"];
  nwr(around:{radius},{center_lat},{center_lon})["name"]["leisure"~"^(water_park|escape_game|sports_centre|park|garden|marina|bowling_alley|amusement_arcade)$"];
  nwr(around:{radius},{center_lat},{center_lon})["name"]["historic"~"^(castle|monument|memorial|archaeological_site|ruins|manor|fort)$"];
);
out center tags 120;
"""
        try:
            payload = overpass(query)
        except Exception as exc:
            failures += 1
            candidates.append(
                {
                    "source": "OpenStreetMap / Overpass",
                    "area": sheet,
                    "status": "source-error",
                    "reason": str(exc)[:300],
                }
            )
            continue

        ranked: list[tuple[int, float, dict[str, Any]]] = []
        for element in payload.get("elements", []):
            tags = element.get("tags", {})
            name = str(tags.get("name") or "").strip()
            website = str(
                tags.get("website")
                or tags.get("contact:website")
                or tags.get("wikidata")
                or tags.get("wikipedia")
                or ""
            ).strip()
            if not name:
                continue
            center = element.get("center") or element
            try:
                latitude = float(center.get("lat"))
                longitude = float(center.get("lon"))
                distance = haversine(center_lat, center_lon, latitude, longitude)
            except (TypeError, ValueError):
                distance = 0.0
            score = 0
            score += 3 if website else 0
            score += 2 if tags.get("wikidata") else 0
            score += 1 if tags.get("opening_hours") else 0
            score += 1 if tags.get("tourism") else 0
            score += 1 if tags.get("historic") else 0
            if score >= 3:
                ranked.append((score, distance, element))
        ranked.sort(
            key=lambda pair: (
                -pair[0],
                pair[1],
                str(pair[2].get("tags", {}).get("name", "")),
            )
        )

        proposed: list[dict[str, Any]] = []
        for score, distance, element in ranked[:40]:
            tags = element.get("tags", {})
            name = str(tags.get("name") or "").strip()
            checked += 1
            osm_type = str(element.get("type", "node"))
            osm_id = str(element.get("id", ""))
            city = str(
                tags.get("addr:city")
                or tags.get("addr:place")
                or tags.get("addr:suburb")
                or sheet
            )
            official = str(tags.get("website") or tags.get("contact:website") or "")
            osm_url = f"https://www.openstreetmap.org/{osm_type}/{osm_id}"
            source = official if official.startswith("http") else osm_url
            category, subcategory, setting, weather = activity_category(tags)
            free = str(tags.get("fee", "")).casefold() == "no"
            proposed.append(
                {
                    "ID": stable_id(f"ACT-{prefix}-AUTO", f"{osm_type}:{osm_id}"),
                    "Activité": name,
                    "Catégorie": category,
                    "Sous-catégorie": subcategory,
                    "Commune": city,
                    "Lieu / zone": city,
                    "Distance approx. (km)": round(distance, 1),
                    "Temps depuis centre": "À calculer selon le moyen de transport",
                    "Durée conseillée": "1 h à 3 h",
                    "Budget": "Gratuit" if free else "À vérifier",
                    "Prix indicatif (€)": 0 if free else None,
                    "Réservation": "À vérifier sur la source officielle",
                    "Saison / période": "Toute l’année — horaires à vérifier",
                    "Cadre": setting,
                    "Météo idéale": weather,
                    "Public": "Tous publics — vérifier les restrictions",
                    "Intensité": "Faible à modérée",
                    "Âge minimum": "À vérifier",
                    "Accessibilité PMR": "À vérifier",
                    "Chien": "À vérifier",
                    "Transport / stationnement": "Consulter l’itinéraire avant le départ",
                    "Description": f"{subcategory} référencé près de {sheet}.",
                    "Pourquoi ça vaut le coup": "Nouvelle activité locale issue d’une base cartographique structurée et reliée à une source vérifiable.",
                    "Conseil pratique": "Vérifier horaires, tarifs et réservation sur le site officiel avant de partir.",
                    "À combiner avec": f"Une autre activité proche de {city}",
                    "Priorité /10": min(9, 5 + score),
                    "Source officielle": source,
                    "Source horaires / tarifs": source,
                    "Vérifié le": french_date(),
                    "_osm_id": f"{osm_type}/{osm_id}",
                }
            )

        accepted = append_records(
            config,
            "activities",
            sheet,
            proposed,
            identity=lambda row: normalize(
                f"{row.get('Activité', '')}|{row.get('Commune', '')}"
            ),
            limit=15,
        )
        added += len(accepted)

    state.update(
        {
            "checkedThisRun": checked,
            "addedThisRun": added,
            "failuresThisRun": failures,
            "areasThisRun": [area[0] for area in selected],
        }
    )
    save_discovery_state("activities", state)
    return {
        "added": added,
        "checked": checked,
        "failures": failures,
        "candidates": candidates,
    }

ENGLISH_FEEDS = (
    {
        "url": "https://feeds.bbci.co.uk/learningenglish/english/features/6-minute-english/rss.xml",
        "provider": "BBC Learning English",
        "skill": "Listening",
        "type": "Podcast",
        "level": "C1",
        "task": "Listen once without subtitles; note 8 expressions; give a 90-second oral summary.",
        "duration": 25,
    },
    {
        "url": "https://www.microsoft.com/en-us/security/blog/feed/",
        "provider": "Microsoft Security Blog",
        "skill": "Cybersecurity English",
        "type": "Professional article",
        "level": "C1–C2",
        "task": "Read the article; extract 10 technical collocations; write a 100-word executive summary.",
        "duration": 35,
    },
    {
        "url": "https://cloudblog.withgoogle.com/rss/",
        "provider": "Google Cloud Blog",
        "skill": "Cloud English",
        "type": "Professional article",
        "level": "C1–C2",
        "task": "Read one section; explain the business value and risk in a two-minute briefing.",
        "duration": 30,
    },
    {
        "url": "https://www.cisa.gov/cybersecurity-advisories/all.xml",
        "provider": "CISA",
        "skill": "Cybersecurity reading",
        "type": "Official advisory",
        "level": "C2",
        "task": "Identify the threat, affected assets and mitigations; brief them in plain English.",
        "duration": 30,
    },
)


def discover_english(config: dict[str, Any]) -> dict[str, Any]:
    failures = 0
    checked = 0
    proposed: list[dict[str, Any]] = []
    candidates: list[dict[str, Any]] = []
    cutoff = TODAY - dt.timedelta(days=120)

    for feed in ENGLISH_FEEDS:
        try:
            entries = parse_feed(str(feed["url"]))
        except Exception as exc:
            failures += 1
            candidates.append(
                {
                    "source": feed["url"],
                    "status": "source-error",
                    "reason": str(exc)[:300],
                }
            )
            continue
        accepted_from_feed = 0
        for entry in entries:
            published = entry.get("date")
            if published is not None and published < cutoff:
                continue
            checked += 1
            proposed.append(
                {
                    "ID": stable_id("RES-AUTO", str(entry["url"])),
                    "Skill": feed["skill"],
                    "Type": feed["type"],
                    "Level": feed["level"],
                    "Provider": feed["provider"],
                    "Resource": entry["title"],
                    "Task": feed["task"],
                    "Duration (min)": feed["duration"],
                    "URL": entry["url"],
                    "Verified on": french_date(),
                    "_published": published.isoformat() if published else None,
                }
            )
            accepted_from_feed += 1
            if accepted_from_feed >= 6:
                break

    accepted = append_records(
        config,
        "english",
        "12 - Reading & Listening",
        proposed,
        identity=lambda row: normalize(row.get("URL")),
        limit=24,
    )
    state = discovery_state("english")
    state.update(
        {
            "checkedThisRun": checked,
            "addedThisRun": len(accepted),
            "failuresThisRun": failures,
            "feedsThisRun": [feed["url"] for feed in ENGLISH_FEEDS],
        }
    )
    save_discovery_state("english", state)
    return {
        "added": len(accepted),
        "checked": checked,
        "failures": failures,
        "candidates": candidates,
    }


TECH_FEEDS = (
    ("Apple", "https://www.apple.com/newsroom/rss-feed.rss"),
    ("Samsung", "https://news.samsung.com/global/feed"),
    ("NVIDIA", "https://blogs.nvidia.com/feed/"),
    ("Microsoft", "https://news.microsoft.com/source/feed/"),
    ("PlayStation", "https://blog.playstation.com/feed/"),
    ("Google", "https://blog.google/rss/"),
)

ANNOUNCEMENT_RE = re.compile(
    r"\b(introduc(?:e|es|ing)|unveil(?:s|ed)?|announce(?:s|d)?|launch(?:es|ed)?|"
    r"reveal(?:s|ed)?|debut(?:s|ed)?|present(?:s|ed)?|new|nouveau|nouvelle)\b",
    re.IGNORECASE,
)


def tech_destination(title: str) -> tuple[str, str, str, str] | None:
    value = normalize(title)
    routes = (
        (("iphone", "galaxy phone", "pixel phone", "smartphone", "wearable", "smartwatch", "watch", "earbuds", "ring"), "06 - Mobiles & wearables", "Mobiles & wearables", "Nouveauté mobile / wearable", "Mobile"),
        (("macbook", "laptop", "notebook", "surface pro", "surface laptop", "copilot pc", "desktop pc"), "05 - Ordinateurs", "Ordinateurs", "Nouvel ordinateur", "Ordinateur"),
        (("geforce", "radeon", "ryzen", "processor", "gpu", "graphics card", "cpu", "ssd"), "04 - Composants PC", "Composants PC", "Nouveau composant", "Composant"),
        (("monitor", "display", "television", "projector", "oled tv", "microled"), "07 - Écrans & projection", "Écrans & projection", "Nouvel écran / projecteur", "Affichage"),
        (("headphone", "headset", "speaker", "soundbar", "audio", "airpods"), "08 - Audio", "Audio", "Nouveau produit audio", "Audio"),
        (("playstation", "xbox", "nintendo", "gaming", "game console", "virtual reality", "vr headset", "meta quest"), "09 - Gaming & VR", "Gaming & VR", "Nouveauté gaming / VR", "Gaming"),
        (("camera", "lens", "drone", "gimbal"), "10 - Photo vidéo drones", "Photo vidéo drones", "Nouveauté photo / vidéo", "Image"),
        (("router", "wi fi", "wifi", "nas", "network storage"), "11 - Réseau & stockage", "Réseau & stockage", "Nouveauté réseau / stockage", "Réseau"),
        (("smart home", "homepod", "nest ", "smartthings", "connected home"), "12 - Maison connectée", "Maison connectée", "Nouveauté maison connectée", "Maison"),
        (("keyboard", "mouse", "dock", "webcam", "office accessory"), "13 - Bureau & accessoires", "Bureau & accessoires", "Nouvel accessoire", "Accessoire"),
        (("battery", "solar", "power station", "3d printer", "maker"), "14 - Makers & énergie", "Makers & énergie", "Nouveauté maker / énergie", "Énergie"),
        (("automotive", "electric vehicle", "mobility", "car platform"), "15 - Mobilité & auto", "Mobilité & auto", "Nouveauté mobilité", "Mobilité"),
    )
    for terms, sheet, domain, category, subcategory in routes:
        if any(term in value for term in terms):
            return sheet, domain, category, subcategory
    return None


def discover_tech(config: dict[str, Any]) -> dict[str, Any]:
    failures = 0
    checked = 0
    candidates: list[dict[str, Any]] = []
    by_sheet: dict[str, list[dict[str, Any]]] = {}
    source_rows: list[dict[str, Any]] = []
    cutoff = TODAY - dt.timedelta(days=150)

    for vendor, feed_url in TECH_FEEDS:
        try:
            entries = parse_feed(feed_url)
        except Exception as exc:
            failures += 1
            candidates.append(
                {
                    "source": feed_url,
                    "vendor": vendor,
                    "status": "source-error",
                    "reason": str(exc)[:300],
                }
            )
            continue
        kept = 0
        for entry in entries:
            published = entry.get("date")
            if published is not None and published < cutoff:
                continue
            checked += 1
            route = tech_destination(str(entry["title"]))
            if route is None or not ANNOUNCEMENT_RE.search(str(entry["title"])):
                continue
            sheet, domain, category, subcategory = route
            identifier = stable_id("TECH-AUTO", str(entry["url"]))
            record = {
                "ID": identifier,
                "Domaine": domain,
                "Catégorie": category,
                "Sous-catégorie": subcategory,
                "Besoin": "Suivre une nouveauté constructeur avant comparaison",
                "Niveau": "🆕 Nouveauté officielle à évaluer",
                "Produit": entry["title"],
                "Marque": vendor,
                "Prix repère (€)": None,
                "Bon prix ≤ (€)": None,
                "Spécifications clés": entry.get("summary") or "Consulter l’annonce officielle.",
                "Pourquoi ce choix": "Annonce officielle récente détectée automatiquement.",
                "Points forts": "À confirmer par des tests indépendants.",
                "Limites / compromis": "Prix français, disponibilité et performances réelles à vérifier.",
                "Compatibilité / prérequis": "Vérifier la fiche officielle et la compatibilité avant achat.",
                "Acheter maintenant ?": "Attendre les tests indépendants et le prix français.",
                "Alternative utile": "Comparer au modèle précédent et aux concurrents directs.",
                "Source officielle": entry["url"],
                "Vérifié le": french_date(),
                "Performance /10": None,
                "Qualité /10": None,
                "Rapport Q/P /10": None,
                "Durabilité /10": None,
                "_published": published.isoformat() if published else None,
            }
            by_sheet.setdefault(sheet, []).append(record)
            source_rows.append(
                {
                    "ID": stable_id("SRC-AUTO", str(entry["url"])),
                    "Domaine": domain,
                    "Catégorie": category,
                    "Type": "Annonce constructeur officielle",
                    "Source": entry["url"],
                    "Vérifié le": french_date(),
                    "Cadence": "Automatique — flux officiel",
                    "Utilisation": "Détection de nouveauté ; spécifications et disponibilité à confirmer",
                    "Note": f"Annonce publiée par {vendor}; les tests indépendants restent nécessaires.",
                }
            )
            kept += 1
            if kept >= 5:
                break

    added = 0
    for sheet, rows in by_sheet.items():
        accepted = append_records(
            config,
            "tech",
            sheet,
            rows,
            identity=lambda row: normalize(row.get("Source officielle")),
            limit=8,
        )
        added += len(accepted)
    append_records(
        config,
        "tech",
        "16 - Sources & méthode",
        source_rows,
        identity=lambda row: normalize(row.get("Source")),
        limit=40,
    )

    state = discovery_state("tech")
    state.update(
        {
            "checkedThisRun": checked,
            "addedThisRun": added,
            "failuresThisRun": failures,
            "feedsThisRun": [url for _vendor, url in TECH_FEEDS],
        }
    )
    save_discovery_state("tech", state)
    return {
        "added": added,
        "checked": checked,
        "failures": failures,
        "candidates": candidates,
    }


GCDL_FEED = "https://cloud.google.com/feeds/gcp-release-notes.xml"
GCDL_KEYWORDS = (
    "artificial intelligence",
    "generative ai",
    "machine learning",
    "vertex ai",
    "data analytics",
    "bigquery",
    "security",
    "identity",
    "cloud storage",
    "compute engine",
    "kubernetes",
    "cloud run",
    "network",
    "operations",
    "observability",
    "sustainability",
    "cost",
    "billing",
)


def discover_gcdl(config: dict[str, Any]) -> dict[str, Any]:
    failures = 0
    checked = 0
    candidates: list[dict[str, Any]] = []
    proposed: list[dict[str, Any]] = []
    cutoff = TODAY - dt.timedelta(days=75)
    try:
        entries = parse_feed(GCDL_FEED)
    except Exception as exc:
        entries = []
        failures = 1
        candidates.append(
            {
                "source": GCDL_FEED,
                "status": "source-error",
                "reason": str(exc)[:300],
            }
        )

    for entry in entries:
        published = entry.get("date")
        if published is not None and published < cutoff:
            continue
        checked += 1
        searchable = normalize(f"{entry['title']} {entry.get('summary', '')}")
        if not any(normalize(keyword) in searchable for keyword in GCDL_KEYWORDS):
            continue
        summary = entry.get("summary") or "Consulter la note de version officielle."
        proposed.append(
            {
                "Domaine / Domain": "7 · Veille officielle\n7 · Official updates",
                "Sous-thème (FR)": "Nouveauté Google Cloud à replacer dans le programme",
                "Official term / service (EN)": entry["title"],
                "Exam keywords / definition (EN)": summary[:900],
                "Explication claire (FR)": "Évolution officielle récente de Google Cloud. Elle devient matière d’examen uniquement si elle apparaît dans le guide officiel.",
                "Quand l’utiliser / exemple (FR)": "Veille produit et mise à jour des connaissances après vérification du guide d’examen.",
                "Valeur métier (FR)": "Maintient le classeur aligné sur l’évolution réelle de Google Cloud.",
                "À ne pas confondre / piège (FR)": "Une release note récente n’est pas automatiquement au programme de la certification.",
                "Priorité": "Veille — hors programme à confirmer",
                "Version de l’examen": "Veille officielle",
                "Source officielle": entry["url"],
                "ID": stable_id("GCDL-AUTO", str(entry["url"])),
                "_published": published.isoformat() if published else None,
            }
        )
        if len(proposed) >= 18:
            break

    accepted = append_records(
        config,
        "gcdl",
        "01 Cours complet",
        proposed,
        identity=lambda row: normalize(row.get("Source officielle")),
        limit=18,
    )
    state = discovery_state("gcdl")
    state.update(
        {
            "checkedThisRun": checked,
            "addedThisRun": len(accepted),
            "failuresThisRun": failures,
            "feed": GCDL_FEED,
        }
    )
    save_discovery_state("gcdl", state)
    return {
        "added": len(accepted),
        "checked": checked,
        "failures": failures,
        "candidates": candidates,
    }


DISCOVERERS: dict[str, Callable[[dict[str, Any]], dict[str, Any]]] = {
    "promo": discover_promo,
    "pokemon": discover_pokemon,
    "fashion": discover_fashion,
    "tech": discover_tech,
    "activities": discover_activities,
    "english": discover_english,
    "gcdl": discover_gcdl,
}


def discover_catalog(config: dict[str, Any], catalog: str) -> dict[str, Any]:
    discoverer = DISCOVERERS.get(catalog)
    if discoverer is None:
        return {"added": 0, "checked": 0, "failures": 0, "candidates": []}
    try:
        return discoverer(config)
    except Exception as exc:
        state = discovery_state(catalog)
        state.update(
            {
                "checkedThisRun": 0,
                "addedThisRun": 0,
                "failuresThisRun": 1,
                "fatalError": f"{type(exc).__name__}: {exc}"[:500],
            }
        )
        save_discovery_state(catalog, state)
        return {
            "added": 0,
            "checked": 0,
            "failures": 1,
            "candidates": [
                {
                    "source": "moteur de découverte",
                    "status": "source-error",
                    "reason": state["fatalError"],
                }
            ],
        }
