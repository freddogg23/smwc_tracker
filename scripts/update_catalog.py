#!/usr/bin/env python3
"""Update the public SMWCentral catalog used by the Community Tracker."""
from __future__ import annotations

import csv
import hashlib
import io
import json
import re
import sys
import time
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

API_URL = "https://www.smwcentral.net/ajax.php"
DETAIL_URL = "https://www.smwcentral.net/?a=details&id={id}&p=section"
PAGE_DELAY_SECONDS = 2.5
MAX_ATTEMPTS = 10

DIFFICULTY = {
    "diff_1": "Newcomer",
    "diff_2": "Casual",
    "diff_3": "Intermediate",
    "diff_4": "Advanced",
    "diff_5": "Expert",
    "diff_6": "Master",
    "diff_7": "Grandmaster",
}

HEADERS = [
    "Dropdown Selection",
    "ROM Hack Title",
    "Created By",
    "Exits",
    "Difficulty",
    "Type",
    "Added Date",
    "SMWC ID",
    "SMWCentral Page URL",
    "Direct Download URL",
]


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


def request_page(page: int) -> dict[str, Any]:
    query = urlencode({"a": "getsectionlist", "s": "smwhacks", "n": page, "u": "0"})
    request = Request(
        f"{API_URL}?{query}",
        headers={
            "User-Agent": "SMW-Community-Tracker-Catalog-Updater/1.0",
            "Accept": "application/json,text/plain,*/*",
            "Referer": "https://www.smwcentral.net/?p=section&s=smwhacks",
        },
    )

    last_error: Exception | None = None
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            with urlopen(request, timeout=75) as response:
                return json.loads(response.read().decode("utf-8"))
        except HTTPError as exc:
            last_error = exc
            wait = 75 if exc.code in (403, 429) else min(60, 2**attempt)
        except (URLError, TimeoutError, json.JSONDecodeError) as exc:
            last_error = exc
            wait = min(60, 2**attempt)

        if attempt == MAX_ATTEMPTS:
            break
        print(
            f"Page {page} failed on attempt {attempt}/{MAX_ATTEMPTS}; "
            f"retrying in {wait} seconds...",
            flush=True,
        )
        time.sleep(wait)

    raise RuntimeError(f"Page {page} could not be retrieved: {last_error}") from last_error


def collect_catalog() -> tuple[list[dict[str, Any]], int]:
    rows_by_id: dict[str, dict[str, Any]] = {}
    page = 1
    last_page = 1
    reported_total = 0

    while page <= last_page:
        payload = request_page(page)
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
                added_date = (
                    datetime.fromtimestamp(int(timestamp)).strftime("%Y-%m-%d")
                    if timestamp
                    else ""
                )
            except (TypeError, ValueError, OSError, OverflowError):
                added_date = ""

            rows_by_id[hack_id] = {
                "ROM Hack Title": title,
                "Created By": joined(hack.get("authors")) or "Unknown",
                "Exits": int(length_match.group()) if length_match else 0,
                "Difficulty": DIFFICULTY.get(
                    diff_raw,
                    diff_raw.removeprefix("diff_").replace("_", " ").title()
                    if diff_raw
                    else "Unranked",
                ),
                "Type": type_text(raw.get("type", hack.get("type"))),
                "Added Date": added_date,
                "SMWC ID": hack_id,
                "SMWCentral Page URL": DETAIL_URL.format(id=hack_id),
                "Direct Download URL": normalize_url(hack.get("download_url")),
            }

        print(
            f"Fetched page {page}/{last_page}: {len(rows_by_id):,}/{reported_total or '?'}",
            flush=True,
        )
        page += 1
        if page <= last_page:
            time.sleep(PAGE_DELAY_SECONDS)

    ordered = sorted(
        rows_by_id.values(),
        key=lambda row: (row["ROM Hack Title"].casefold(), int(row["SMWC ID"])),
    )
    duplicate_counts = Counter(row["ROM Hack Title"].casefold() for row in ordered)

    rows: list[dict[str, Any]] = []
    for row in ordered:
        title = row["ROM Hack Title"]
        label = title
        if duplicate_counts[title.casefold()] > 1:
            label = f'{title} [SMWC #{row["SMWC ID"]}]'
        rows.append({"Dropdown Selection": label, **row})

    return rows, reported_total


def make_csv(rows: list[dict[str, Any]]) -> bytes:
    buffer = io.StringIO(newline="")
    writer = csv.DictWriter(buffer, fieldnames=HEADERS, lineterminator="\r\n")
    writer.writeheader()
    writer.writerows(rows)
    return b"\xef\xbb\xbf" + buffer.getvalue().encode("utf-8")


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    data_dir = root / "data"
    data_dir.mkdir(parents=True, exist_ok=True)

    rows, reported_total = collect_catalog()
    if reported_total and len(rows) != reported_total:
        raise RuntimeError(
            f"SMWCentral reported {reported_total:,} hacks, but {len(rows):,} unique rows were collected."
        )

    new_csv = make_csv(rows)
    csv_path = data_dir / "SMWCentral_All_Moderated_Hacks.csv"
    json_path = data_dir / "SMWCentral_All_Moderated_Hacks.json"
    version_path = data_dir / "version.json"

    old_csv = csv_path.read_bytes() if csv_path.exists() else b""
    if old_csv == new_csv and json_path.exists() and version_path.exists():
        print("The catalog has not changed; no files were rewritten.")
        return 0

    csv_path.write_bytes(new_csv)
    json_path.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")

    now_utc = datetime.now(timezone.utc).replace(microsecond=0).isoformat()
    version = {
        "catalog_version": now_utc[:10].replace("-", "."),
        "generated_at_utc": now_utc,
        "hack_count": len(rows),
        "sha256": hashlib.sha256(new_csv).hexdigest(),
        "source": "SMWCentral moderated Super Mario World hacks catalog",
    }
    version_path.write_text(json.dumps(version, indent=2), encoding="utf-8")
    print(f"Saved {len(rows):,} unique moderated hacks.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        raise SystemExit(130)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1)
