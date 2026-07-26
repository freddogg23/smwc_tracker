#!/usr/bin/env python3
"""Export every currently moderated SMW hack from SMWCentral to CSV and JSON.

Uses only Python's standard library. It does not download ROMs or patch files.
"""
from __future__ import annotations

import csv
import json
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import urlencode
from urllib.request import Request, urlopen

API_URL = "https://www.smwcentral.net/ajax.php"
DETAIL_URL = "https://www.smwcentral.net/?a=details&id={id}&p=section"
DIFFICULTY = {
    "diff_1": "Newcomer",
    "diff_2": "Casual",
    "diff_3": "Intermediate",
    "diff_4": "Advanced",
    "diff_5": "Expert",
    "diff_6": "Master",
    "diff_7": "Grandmaster",
}


def preferred_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (str, int, float)):
        return str(value).strip()
    if isinstance(value, dict):
        for key in ("name", "username", "display_name", "label", "title", "text", "value"):
            candidate = value.get(key)
            if candidate is not None and str(candidate).strip():
                return str(candidate).strip()
    return str(value).strip()


def joined(value: Any) -> str:
    if isinstance(value, list):
        return ", ".join(filter(None, (preferred_text(item) for item in value)))
    return preferred_text(value)


def type_text(value: Any) -> str:
    values = value if isinstance(value, list) else re.split(r"[,;]", str(value or ""))
    mapping = {
        "standard": "Standard",
        "kaizo": "Kaizo",
        "puzzle": "Puzzle",
        "tool_assisted": "Tool-Assisted",
        "toolassisted": "Tool-Assisted",
        "pit": "Pit",
        "troll": "Troll",
    }
    result: list[str] = []
    for item in values:
        text = preferred_text(item)
        if not text:
            continue
        key = re.sub(r"[- ]", "_", text.lower())
        display = mapping.get(key, text.replace("_", " ").title())
        if display not in result:
            result.append(display)
    return ", ".join(result)


def normalize_url(value: Any) -> str:
    url = str(value or "").strip()
    if url.startswith("//"):
        return "https:" + url
    if url.startswith("/"):
        return "https://www.smwcentral.net" + url
    return url


def fetch_json(page: int) -> dict[str, Any]:
    query = urlencode({"a": "getsectionlist", "s": "smwhacks", "n": page, "u": "0"})
    request = Request(
        f"{API_URL}?{query}",
        headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/150.0 Safari/537.36",
            "Accept": "application/json,text/plain,*/*",
            "Referer": "https://www.smwcentral.net/?p=section&s=smwhacks",
        },
    )

    last_error: Exception | None = None
    for attempt in range(1, 5):
        try:
            with urlopen(request, timeout=60) as response:
                return json.loads(response.read().decode("utf-8"))
        except Exception as exc:  # noqa: BLE001 - retry all network/HTTP failures
            last_error = exc
            if attempt == 4:
                break
            time.sleep(2**attempt)
    raise RuntimeError(f"Page {page} could not be retrieved: {last_error}") from last_error


def main() -> int:
    output_dir = Path(__file__).resolve().parent
    rows_by_id: dict[str, dict[str, Any]] = {}
    page = 1
    last_page = 1
    reported_total = 0

    while page <= last_page:
        payload = fetch_json(page)
        last_page = int(payload.get("last_page") or 1)
        reported_total = int(payload.get("total") or 0)

        for hack in payload.get("data") or []:
            hack_id = str(hack.get("id") or "").strip()
            title = str(hack.get("name") or "").strip()
            if not hack_id or not title or hack_id in rows_by_id:
                continue

            raw = hack.get("raw_fields") or {}
            length_match = re.search(r"-?\d+", str(raw.get("length", hack.get("length", 0))))
            diff_raw = joined(raw.get("difficulty", hack.get("difficulty")))
            timestamp = hack.get("time")
            try:
                added_date = datetime.fromtimestamp(int(timestamp)).strftime("%Y-%m-%d") if timestamp else ""
            except (TypeError, ValueError, OSError, OverflowError):
                added_date = ""

            rows_by_id[hack_id] = {
                "SMWC ID": hack_id,
                "Title": title,
                "Created By": joined(hack.get("authors")) or "Unknown",
                "Exits": int(length_match.group()) if length_match else 0,
                "Difficulty": DIFFICULTY.get(
                    diff_raw,
                    diff_raw.removeprefix("diff_").replace("_", " ").title() if diff_raw else "No Difficulty",
                ),
                "Type": type_text(raw.get("type", hack.get("type"))),
                "Added Date": added_date,
                "SMWCentral Detail URL": DETAIL_URL.format(id=hack_id),
                "Direct Download URL": normalize_url(hack.get("download_url")),
            }

        print(f"Fetched page {page}/{last_page}: {len(rows_by_id)}/{reported_total or '?'}", flush=True)
        page += 1
        if page <= last_page:
            time.sleep(2.5)

    rows = sorted(rows_by_id.values(), key=lambda row: (row["Title"].casefold(), int(row["SMWC ID"])))
    csv_path = output_dir / "SMWCentral_All_Moderated_Hacks.csv"
    json_path = output_dir / "SMWCentral_All_Moderated_Hacks.json"

    fieldnames = [
        "SMWC ID", "Title", "Created By", "Exits", "Difficulty", "Type",
        "Added Date", "SMWCentral Detail URL", "Direct Download URL",
    ]
    with csv_path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    with json_path.open("w", encoding="utf-8") as handle:
        json.dump(rows, handle, ensure_ascii=False, indent=2)

    print(f"Saved {len(rows)} unique moderated hacks to:")
    print(csv_path)
    print(json_path)
    if reported_total and len(rows) != reported_total:
        print(f"Warning: API reported {reported_total}, but {len(rows)} unique records were saved.")
        return 2
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\nCancelled.", file=sys.stderr)
        raise SystemExit(130)
    except Exception as exc:  # noqa: BLE001
        print(f"\nERROR: {exc}", file=sys.stderr)
        raise SystemExit(1)
