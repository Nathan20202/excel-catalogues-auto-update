#!/usr/bin/env python3
"""Write a monthly activity heartbeat so scheduled workflows remain active."""

from __future__ import annotations

import datetime as dt
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PATH = ROOT / "data" / "heartbeat.json"
now = dt.datetime.now(dt.timezone.utc).replace(microsecond=0)
payload = {
    "schemaVersion": 1,
    "updatedAt": now.isoformat().replace("+00:00", "Z"),
    "purpose": "Monthly activity marker for the free scheduled update infrastructure.",
}
PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
print(payload["updatedAt"])

