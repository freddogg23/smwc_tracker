#!/usr/bin/env python3
"""Update the SMWCentral catalog and publish incremental delta files.

The GitHub runner still checks the complete moderated catalog so it can detect
all additions, metadata updates, and removals. Downloaded Excel workbooks do
NOT download the whole catalog on each refresh. They download only version.json
and any missing data/deltas/########.csv files.
"""
from __future__ import annotations

import csv
import json
import re
import sys
import time
import traceback
from collections import Counter
from datetime import datetime, timezone
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
    "SMWCentral Rating",
]
DELTA_HEADERS = ["Operation", *HEADERS]


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


def clean_rating(value: Any) -> str:
    """Return a compact rating value if the API exposes one."""
    if value is None:
        return ""
    if isinstance(value, dict):
        for key in ("average", "avg", "value", "rating", "score", "stars"):
            if key in value and value[key] not in (None, ""):
                return clean_rating(value[key])
        return ""
    if isinstance(value, (int, float)):
        return f"{float(value):.2f}".rstrip("0").rstrip(".")
    text = str(value).strip()
    if not text:
        return ""
    match = re.search(r"\d+(?:\.\d+)?", text)
    return match.group(0) if match else text


def extract_rating(hack: dict[str, Any]) -> str:
    containers = [
        hack,
        hack.get("raw_fields") or {},
        hack.get("fields") or {},
        hack.get("stats") or {},
        hack.get("metadata") or {},
    ]
    for container in containers:
        if not isinstance(container, dict):
            continue
        for key in (
            "rating",
            "ratings",
            "rating_average",
            "rating_avg",
            "average_rating",
            "avg_rating",
            "score",
            "stars",
        ):
            if key in container:
                rating = clean_rating(container.get(key))
                if rating:
                    return rating
    return ""


def fetch_json(page: int) -> dict[str, Any]:
    query = urlencode({"a": "getsectionlist", "s": "smwhacks", "n": page, "u": "0"})
    request = Request(
        f"{API_URL}?{query}",
        headers={
            "User-Agent": "SMW-Hack-Tracker-Catalog-Updater/2.0",
            "Accept": "application/json,text/plain,*/*",
            "Referer": "https://www.smwcentral.net/?p=section&s=smwhacks",
        },
    )
    last_error: Exception | None = None
    for attempt in range(1, 6):
        try:
            with urlopen(request, timeout=60) as response:
                return json.loads(response.read().decode("utf-8"))
        except Exception as exc:  # noqa: BLE001
            last_error = exc
            if attempt < 5:
                time.sleep(min(120, 2**attempt))
    raise RuntimeError(f"Page {page} could not be retrieved: {last_error}") from last_error


def build_rows() -> list[dict[str, Any]]:
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
                "Dropdown Selection": title,
                "ROM Hack Title": title,
                "Created By": joined(hack.get("authors")) or "Unknown",
                "Exits": int(length_match.group()) if length_match else 0,
                "Difficulty": DIFFICULTY.get(
                    diff_raw,
                    diff_raw.removeprefix("diff_").replace("_", " ").title() if diff_raw else "Unranked",
                ),
                "Type": type_text(raw.get("type", hack.get("type"))),
                "Added Date": added_date,
                "SMWC ID": hack_id,
                "SMWCentral Page URL": DETAIL_URL.format(id=hack_id),
                "Direct Download URL": normalize_url(hack.get("download_url")),
                "SMWCentral Rating": extract_rating(hack),
            }

        print(f"Fetched page {page}/{last_page}: {len(rows_by_id)}/{reported_total or '?'}")
        page += 1
        if page <= last_page:
            time.sleep(2.5)

    rows = sorted(
        rows_by_id.values(),
        key=lambda row: (str(row["ROM Hack Title"]).casefold(), int(row["SMWC ID"])),
    )
    title_counts = Counter(str(row["ROM Hack Title"]).casefold() for row in rows)
    for row in rows:
        title = str(row["ROM Hack Title"])
        row["Dropdown Selection"] = (
            f'{title} [SMWC #{row["SMWC ID"]}]'
            if title_counts[title.casefold()] > 1
            else title
        )
    return rows


def canonical_row(row: dict[str, Any]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for header in HEADERS:
        value = row.get(header, "")
        if header == "Exits":
            try:
                value = int(value)
            except (TypeError, ValueError):
                value = 0
        else:
            value = str(value or "").strip()
        result[header] = value
    return result


def load_old_rows(data_dir: Path) -> list[dict[str, Any]]:
    json_path = data_dir / "SMWCentral_All_Moderated_Hacks.json"
    if json_path.exists():
        try:
            payload = json.loads(json_path.read_text(encoding="utf-8-sig"))
            if isinstance(payload, dict) and "value" in payload:
                payload = payload["value"]
            if isinstance(payload, list):
                return [canonical_row(row) for row in payload if isinstance(row, dict)]
        except Exception:  # noqa: BLE001
            pass

    csv_path = data_dir / "SMWCentral_All_Moderated_Hacks.csv"
    if csv_path.exists():
        with csv_path.open("r", newline="", encoding="utf-8-sig") as handle:
            return [canonical_row(row) for row in csv.DictReader(handle)]
    return []


def load_version(data_dir: Path) -> dict[str, Any]:
    path = data_dir / "version.json"
    if not path.exists():
        return {"sequence": 0}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        return payload if isinstance(payload, dict) else {"sequence": 0}
    except Exception:  # noqa: BLE001
        return {"sequence": 0}


def is_official_smwc_id(value: Any) -> bool:
    """Official SMWCentral section IDs are numeric.

    The repository may also contain curated/manual catalog entries whose ID is
    something like ``N/A`` or ``MANUAL-...``. Those entries must not be passed
    to ``int()`` or deleted just because they are not returned by the moderated
    SMWCentral API.
    """
    return str(value or "").strip().isdigit()


def smwc_id_sort_key(value: Any) -> tuple[int, Any]:
    text = str(value or "").strip()
    if text.isdigit():
        return (0, int(text))
    return (1, text.casefold())


def index_by_id(rows: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for row in rows:
        smwc_id = str(row.get("SMWC ID", "")).strip()
        if is_official_smwc_id(smwc_id):
            result[smwc_id] = canonical_row(row)
    return result


def curated_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Preserve nonnumeric catalog entries maintained in the repository."""
    return [
        canonical_row(row)
        for row in rows
        if not is_official_smwc_id(row.get("SMWC ID", ""))
        and str(row.get("ROM Hack Title", "")).strip()
    ]


def calculate_changes(
    old_rows: list[dict[str, Any]],
    new_rows: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    old = index_by_id(old_rows)
    new = index_by_id(new_rows)

    added = [
        new[key]
        for key in sorted(new.keys() - old.keys(), key=smwc_id_sort_key)
    ]
    removed = [
        old[key]
        for key in sorted(old.keys() - new.keys(), key=smwc_id_sort_key)
    ]
    updated = [
        new[key]
        for key in sorted(new.keys() & old.keys(), key=smwc_id_sort_key)
        if new[key] != old[key]
    ]
    return added, updated, removed


def write_full_catalog(data_dir: Path, rows: list[dict[str, Any]]) -> None:
    with (data_dir / "SMWCentral_All_Moderated_Hacks.csv").open(
        "w", newline="", encoding="utf-8-sig"
    ) as handle:
        writer = csv.DictWriter(handle, fieldnames=HEADERS)
        writer.writeheader()
        writer.writerows(rows)

    (data_dir / "SMWCentral_All_Moderated_Hacks.json").write_text(
        json.dumps(rows, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def write_delta(
    data_dir: Path,
    sequence: int,
    added: list[dict[str, Any]],
    updated: list[dict[str, Any]],
    removed: list[dict[str, Any]],
) -> str:
    deltas_dir = data_dir / "deltas"
    deltas_dir.mkdir(exist_ok=True)
    filename = f"{sequence:08d}.csv"
    path = deltas_dir / filename

    operations: list[dict[str, Any]] = []
    for operation, rows in (("UPSERT", added), ("UPSERT", updated), ("DELETE", removed)):
        for row in rows:
            operations.append({"Operation": operation, **canonical_row(row)})

    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=DELTA_HEADERS)
        writer.writeheader()
        writer.writerows(operations)
    return f"data/deltas/{filename}"


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    data_dir = root / "data"
    data_dir.mkdir(exist_ok=True)

    old_rows = load_old_rows(data_dir)
    old_version = load_version(data_dir)
    old_sequence = int(old_version.get("sequence", 0) or 0)

    api_rows = [canonical_row(row) for row in build_rows()]
    preserved_rows = curated_rows(old_rows)
    added, updated, removed = calculate_changes(old_rows, api_rows)

    # The full downloadable catalog contains current official API rows plus any
    # curated nonnumeric entries already maintained in the repository.
    new_rows = sorted(
        [*api_rows, *preserved_rows],
        key=lambda row: (
            str(row.get("ROM Hack Title", "")).casefold(),
            smwc_id_sort_key(row.get("SMWC ID", "")),
        ),
    )

    if not added and not updated and not removed:
        print("No catalog changes. No delta file was created.")
        # Ensure an old-style version file is migrated to incremental format once.
        if "sequence" not in old_version or old_version.get("refresh_mode") != "incremental":
            old_version.update(
                {
                    "sequence": old_sequence,
                    "latest_delta": old_version.get("latest_delta", ""),
                    "added_count": 0,
                    "updated_count": 0,
                    "removed_count": 0,
                    "refresh_mode": "incremental",
                    "schema": "v3-incremental-rating-random-finder",
                }
            )
            (data_dir / "version.json").write_text(
                json.dumps(old_version, indent=2), encoding="utf-8"
            )
        return 0

    sequence = old_sequence + 1
    delta_file = write_delta(data_dir, sequence, added, updated, removed)
    write_full_catalog(data_dir, new_rows)

    now = datetime.now(timezone.utc).replace(microsecond=0)
    version = {
        "catalog_version": now.strftime("%Y.%m.%d"),
        "generated_at_utc": now.isoformat(),
        "hack_count": len(new_rows),
        "sequence": sequence,
        "latest_delta": delta_file,
        "added_count": len(added),
        "updated_count": len(updated),
        "removed_count": len(removed),
        "curated_count": len(preserved_rows),
        "refresh_mode": "incremental",
        "schema": "v3-incremental-rating-random-finder",
        "source": "SMWCentral moderated Super Mario World hacks catalog",
    }
    (data_dir / "version.json").write_text(
        json.dumps(version, indent=2), encoding="utf-8"
    )

    print(
        f"Published delta {sequence:08d}: "
        f"{len(added)} added, {len(updated)} updated, {len(removed)} removed, "
        f"{len(preserved_rows)} curated entries preserved."
    )
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        raise SystemExit(130)
    except Exception as exc:  # noqa: BLE001
        print(f"ERROR: {exc}", file=sys.stderr)
        traceback.print_exc()
        raise SystemExit(1)
