#!/usr/bin/env python3
"""Validate catalogue schemas and privacy rules with the Python standard library."""

from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CONFIG_PATH = ROOT / "config" / "workbooks.json"
MANIFEST_PATH = ROOT / "data" / "manifest.json"

FORBIDDEN_FRAGMENTS = (
    "github_pat_",
    "ghp_",
)


def fail(message: str) -> None:
    print(f"ERREUR: {message}", file=sys.stderr)
    raise SystemExit(1)


def load_json(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:  # pragma: no cover - diagnostic path
        fail(f"{path.relative_to(ROOT)} n'est pas un JSON valide: {exc}")


def main() -> None:
    config = load_json(CONFIG_PATH)
    manifest = load_json(MANIFEST_PATH)
    if config.get("schemaVersion") != 1:
        fail("schemaVersion de config/workbooks.json non pris en charge")
    if manifest.get("schemaVersion") != 1:
        fail("schemaVersion de data/manifest.json non pris en charge")

    checked_files = 0
    checked_records = 0
    for key, workbook in config.get("workbooks", {}).items():
        datasets = workbook.get("datasets", [])
        if not datasets:
            fail(f"{key}: aucun jeu de données")
        seen_by_table: dict[str, set[str]] = {}
        for dataset in datasets:
            relative = dataset.get("file")
            path = ROOT / str(relative)
            if not path.is_file():
                fail(f"{key}: fichier absent: {relative}")
            payload = load_json(path)
            records = payload.get("records")
            if not isinstance(records, list):
                fail(f"{relative}: records doit être une liste")
            if payload.get("recordCount") != len(records):
                fail(f"{relative}: recordCount ne correspond pas au nombre de lignes")
            id_column = dataset.get("idColumn")
            table_key = f"{dataset.get('sheet')}|{dataset.get('table')}"
            seen = seen_by_table.setdefault(table_key, set())
            for index, record in enumerate(records, start=1):
                if not isinstance(record, dict):
                    fail(f"{relative}: ligne {index} non objet")
                identifier = str(record.get(id_column, "")).strip()
                if not identifier:
                    fail(f"{relative}: ID vide à la ligne {index}")
                if identifier in seen:
                    fail(f"{relative}: ID dupliqué {identifier}")
                seen.add(identifier)
                for protected in dataset.get("preserveColumns", []):
                    if protected in record:
                        fail(f"{relative}: colonne personnelle publiée: {protected}")
                for formula in dataset.get("formulaColumns", []):
                    if formula in record:
                        fail(f"{relative}: colonne de formule publiée: {formula}")
            checked_files += 1
            checked_records += len(records)

    for path in ROOT.rglob("*"):
        if not path.is_file() or ".git" in path.parts:
            continue
        if path.resolve() == Path(__file__).resolve():
            continue
        if path.stat().st_size > 8_000_000:
            continue
        try:
            content = path.read_text(encoding="utf-8", errors="ignore")
        except OSError:
            continue
        lowered = content.casefold()
        for fragment in FORBIDDEN_FRAGMENTS:
            if fragment.casefold() in lowered:
                fail(f"information sensible détectée dans {path.relative_to(ROOT)}")

    print(f"OK — {checked_files} fichiers et {checked_records} enregistrements validés.")


if __name__ == "__main__":
    main()
