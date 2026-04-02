from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import urlopen

import pandas as pd


def _fetch_json(url: str) -> dict:
    with urlopen(url) as response:
        return json.loads(response.read().decode("utf-8"))


def _extract_codes_from_url(url: str) -> tuple[str, int]:
    m = re.search(r"/events/([^/]+)/race/(\d+)", url)
    if not m:
        raise ValueError(
            "Could not parse event code and race id from URL. "
            "Expected format .../events/<eventCode>/race/<raceId>."
        )
    return m.group(1), int(m.group(2))


def _resolve_event_and_race(event_code: str, race_id: int) -> tuple[int, int, str]:
    event_url = f"https://results.raceroster.com/v2/api/events/{event_code}"
    event_payload = _fetch_json(event_url)
    event_data = event_payload.get("data", {}).get("event", {})
    result_event_id = int(event_data.get("resultEventId"))

    race_url = f"https://results.raceroster.com/v2/api/events/{event_code}/sub-events/{race_id}"
    race_payload = _fetch_json(race_url)
    race_data = race_payload.get("data", {})
    result_sub_event_id = int(race_data.get("resultSubEventId") or race_data.get("id"))
    race_name = str(race_data.get("name") or f"Race {race_id}")

    return result_event_id, result_sub_event_id, race_name


def _fetch_all_rows(result_event_id: int, sub_event_id: int) -> list[dict]:
    start = 0
    limit = 50
    all_rows: list[dict] = []

    while True:
        query = urlencode({"start": start, "limit": limit})
        url = (
            f"https://results.raceroster.com/v2/api/result-events/"
            f"{result_event_id}/sub-events/{sub_event_id}/results?{query}"
        )
        payload = _fetch_json(url)
        rows = payload.get("data", [])
        if not rows:
            break

        all_rows.extend(rows)
        if len(rows) < limit:
            break

        start += len(rows)

    return all_rows


def _to_wrrl_rows(rows: list[dict]) -> list[dict]:
    out: list[dict] = []
    for row in rows:
        chip = str(row.get("chipTime") or "").strip()
        gun = str(row.get("gunTime") or "").strip()
        gender = str(row.get("genderSexId") or "").strip().lower()

        out.append(
            {
                "Pos": row.get("overallPlace", ""),
                "Bib": row.get("bib", ""),
                "Name": row.get("name", ""),
                "Club": row.get("custom23101578", ""),
                "Gender": "F" if gender.startswith("w") else "M",
                "Category": row.get("custom23096209", ""),
                "Time": chip or gun,
                "Chip Time": chip,
                "Gun Time": gun,
                "Delta": row.get("custom23096212", ""),
                "Category Pos": row.get("custom23096210", ""),
                "Source ID": row.get("id", ""),
            }
        )
    return out


def export_from_race_url(race_url: str, output_dir: Path) -> tuple[Path, Path, int]:
    event_code, race_id = _extract_codes_from_url(race_url)
    result_event_id, result_sub_event_id, race_name = _resolve_event_and_race(event_code, race_id)
    rows = _fetch_all_rows(result_event_id, result_sub_event_id)
    wrll_rows = _to_wrrl_rows(rows)

    output_dir.mkdir(parents=True, exist_ok=True)
    safe_race_name = re.sub(r"[^A-Za-z0-9 _#-]", "", race_name).strip() or f"Race {race_id}"
    xlsx_path = output_dir / f"Race #{race_id} - {safe_race_name} (extracted).xlsx"
    csv_path = output_dir / f"Race #{race_id} - {safe_race_name} (extracted).csv"

    df = pd.DataFrame(wrll_rows)
    df.to_excel(xlsx_path, index=False)
    df.to_csv(csv_path, index=False)

    return xlsx_path, csv_path, len(wrll_rows)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extract Race Roster leaderboard rows into WRRL-friendly files."
    )
    parser.add_argument("race_url", help="Race Roster race URL (contains /events/<code>/race/<id>)")
    parser.add_argument(
        "--out",
        default="output/division-format-check",
        help="Output folder for extracted files",
    )

    args = parser.parse_args()
    xlsx_path, csv_path, count = export_from_race_url(args.race_url, Path(args.out))
    print(f"Extracted rows: {count}")
    print(f"XLSX: {xlsx_path}")
    print(f"CSV:  {csv_path}")


if __name__ == "__main__":
    main()
