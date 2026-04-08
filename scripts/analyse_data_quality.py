"""Analyse source-to-audited data quality for a WRRL season.

Focuses on upstream quality indicators that materially affect downstream scoring:
- blank names
- blank categories
- blank genders
- invalid/missing times
- schema warnings

Outputs JSON and Markdown reports under outputs/quality/data-quality/year-<season>/.
"""

from __future__ import annotations

import argparse
import json
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path
import sys
from typing import Any

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from league_scorer.common_files import race_discovery_exclusions
from league_scorer.input_layout import build_input_paths
from league_scorer.normalisation import parse_time_to_seconds
from league_scorer.race_processor import extract_race_number
from league_scorer.race_validation import validate_race_schema
from league_scorer.source_loader import discover_race_files, load_race_dataframe


def _resolve_data_root(explicit: Path | None) -> Path | None:
    if explicit is not None:
        return explicit

    prefs = Path.home() / ".wrrl_prefs.json"
    if prefs.exists():
        try:
            data = json.loads(prefs.read_text(encoding="utf-8"))
            value = data.get("data_root")
            if value:
                return Path(value)
        except Exception:
            return None
    return None


def _find_raw_dir(input_dir: Path) -> Path | None:
    candidate = build_input_paths(input_dir).raw_data_dir
    return candidate if candidate.exists() and candidate.is_dir() else None


def _count_blank(value: Any) -> bool:
    if value is None:
        return True
    text = str(value).strip()
    return text == "" or text.lower() == "nan"


def _profile_dataframe(df: pd.DataFrame, path: Path) -> dict[str, Any]:
    validation = validate_race_schema(df, path)
    col = validation.column_map
    time_col = validation.time_column

    row_count = len(df)

    blank_name = 0
    blank_category = 0
    blank_gender = 0
    invalid_time = 0

    for _, row in df.iterrows():
        if _count_blank(row.get(col["Name"])):
            blank_name += 1
        if _count_blank(row.get(col["Category"])):
            blank_category += 1
        if _count_blank(row.get(col["Gender"])):
            blank_gender += 1

        time_value = row.get(time_col)
        if parse_time_to_seconds(time_value) is None:
            invalid_time += 1

    def pct(value: int) -> float:
        return round((value / row_count * 100.0), 2) if row_count else 0.0

    return {
        "file": str(path),
        "rows": row_count,
        "schema_warning_codes": [issue.code for issue in validation.issues if issue.code],
        "schema_warning_count": len(validation.issues),
        "blank_name": blank_name,
        "blank_name_pct": pct(blank_name),
        "blank_category": blank_category,
        "blank_category_pct": pct(blank_category),
        "blank_gender": blank_gender,
        "blank_gender_pct": pct(blank_gender),
        "invalid_time": invalid_time,
        "invalid_time_pct": pct(invalid_time),
    }


def _aggregate(profiles: list[dict[str, Any]]) -> dict[str, Any]:
    total_rows = sum(item["rows"] for item in profiles)

    def total(name: str) -> int:
        return sum(int(item[name]) for item in profiles)

    def pct(v: int) -> float:
        return round((v / total_rows * 100.0), 2) if total_rows else 0.0

    blanks_cat = total("blank_category")
    blanks_name = total("blank_name")
    blanks_gender = total("blank_gender")
    invalid_time = total("invalid_time")
    schema_warn = sum(int(item["schema_warning_count"]) for item in profiles)

    return {
        "files": len(profiles),
        "rows": total_rows,
        "schema_warnings": schema_warn,
        "blank_name": blanks_name,
        "blank_name_pct": pct(blanks_name),
        "blank_category": blanks_cat,
        "blank_category_pct": pct(blanks_cat),
        "blank_gender": blanks_gender,
        "blank_gender_pct": pct(blanks_gender),
        "invalid_time": invalid_time,
        "invalid_time_pct": pct(invalid_time),
    }


def _top_hotspots(profiles: list[dict[str, Any]], key: str, n: int = 10) -> list[dict[str, Any]]:
    key_pct = f"{key}_pct"
    ranked = sorted(
        [
            {
                "file": item["file"],
                "rows": item["rows"],
                key: item[key],
                key_pct: item[key_pct],
            }
            for item in profiles
            if item["rows"] > 0 and item[key] > 0
        ],
        key=lambda i: (-float(i[key_pct]), -int(i[key]), i["file"].lower()),
    )
    return ranked[:n]


def _build_markdown(payload: dict[str, Any]) -> str:
    raw = payload["raw_summary"]
    audited = payload["audited_summary"]

    lines = [
        "# WRRL Data Quality Analysis",
        "",
        f"Generated: {payload['generated_at']}",
        f"Season: {payload['year']}",
        f"Data Root: {payload['data_root']}",
        "",
        "## Summary",
        "",
        "| Stage | Files | Rows | Blank Category % | Invalid Time % | Blank Name % | Blank Gender % |",
        "| --- | ---: | ---: | ---: | ---: | ---: | ---: |",
        f"| Raw | {raw['files']} | {raw['rows']} | {raw['blank_category_pct']} | {raw['invalid_time_pct']} | {raw['blank_name_pct']} | {raw['blank_gender_pct']} |",
        f"| Audited Inputs | {audited['files']} | {audited['rows']} | {audited['blank_category_pct']} | {audited['invalid_time_pct']} | {audited['blank_name_pct']} | {audited['blank_gender_pct']} |",
        "",
        "## Top Blank Category Hotspots (Audited Inputs)",
        "",
    ]

    hotspots = payload["hotspots"]["audited_blank_category"]
    if not hotspots:
        lines.append("No blank-category hotspots detected.")
    else:
        lines.append("| File | Rows | Blank Category | Blank Category % |")
        lines.append("| --- | ---: | ---: | ---: |")
        for row in hotspots:
            lines.append(
                f"| {Path(row['file']).name} | {row['rows']} | {row['blank_category']} | {row['blank_category_pct']} |"
            )

    recommendations: list[str] = []
    if float(audited.get("blank_category_pct", 0.0)) > 0:
        recommendations.append(
            "Prioritize category cleanup in the top hotspot files first; this gives the fastest reduction in downstream audit noise."
        )
        recommendations.append(
            "Add normalization rules for common category aliases and placeholders (for example blank strings, dash values, and case variants)."
        )

    if float(audited.get("invalid_time_pct", 0.0)) > 0:
        recommendations.append(
            "Pre-clean time values before parse (trim whitespace, standardize separators, and handle hh:mm:ss / mm:ss variants)."
        )

    if hotspots:
        recommendations.append(
            "Create a race-level cleanup checklist for the top 3 hotspots and re-run staged checks after each fix batch to confirm impact."
        )

    lines.extend(
        [
            "",
            "## Interpretation",
            "",
            "- Prioritize files with highest blank-category percentage at source/audited-input stage.",
            "- Re-run the staged checks after source cleanups to validate downstream impact.",
            "",
            "## Suggested Fixes",
            "",
        ]
    )

    if recommendations:
        for recommendation in recommendations:
            lines.append(f"- {recommendation}")
    else:
        lines.append("- No major quality issues detected. Continue routine staged validation.")
    lines.append("")

    return "\n".join(lines)


def _empty_summary() -> dict[str, Any]:
    return {
        "files": 0,
        "rows": 0,
        "schema_warnings": 0,
        "blank_name": 0,
        "blank_name_pct": 0.0,
        "blank_category": 0,
        "blank_category_pct": 0.0,
        "blank_gender": 0,
        "blank_gender_pct": 0.0,
        "invalid_time": 0,
        "invalid_time_pct": 0.0,
    }


def analyse_season(year: int, data_root: Path, output_dir: Path) -> tuple[dict[str, Any], Path, Path]:
    input_dir = data_root / str(year) / "inputs"
    if not input_dir.exists():
        raise FileNotFoundError(f"Input directory not found: {input_dir}")

    raw_dir = _find_raw_dir(input_dir)
    raw_profiles: list[dict[str, Any]] = []

    if raw_dir and raw_dir.exists():
        for raw_file in sorted(raw_dir.rglob("*.xlsx")):
            try:
                raw_df = load_race_dataframe(raw_file)
                raw_profiles.append(_profile_dataframe(raw_df, raw_file))
            except Exception:
                continue

    audited_profiles: list[dict[str, Any]] = []
    audited_dir = build_input_paths(input_dir).audited_dir
    audited_files = discover_race_files(audited_dir, excluded_names=race_discovery_exclusions())
    for race_num, file in sorted(audited_files.items()):
        try:
            df = load_race_dataframe(file)
            profile = _profile_dataframe(df, file)
            profile["race"] = race_num
            audited_profiles.append(profile)
        except Exception:
            continue

    raw_summary = _aggregate(raw_profiles) if raw_profiles else _empty_summary()
    audited_summary = _aggregate(audited_profiles)

    by_race_blank_category: dict[int, int] = defaultdict(int)
    for item in audited_profiles:
        by_race_blank_category[int(item.get("race", 0))] += int(item["blank_category"])

    payload = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "year": year,
        "data_root": str(data_root),
        "input_dir": str(input_dir),
        "raw_dir": str(raw_dir) if raw_dir else "",
        "raw_summary": raw_summary,
        "audited_summary": audited_summary,
        "audited_blank_category_by_race": {str(k): v for k, v in sorted(by_race_blank_category.items())},
        "hotspots": {
            "raw_blank_category": _top_hotspots(raw_profiles, "blank_category"),
            "audited_blank_category": _top_hotspots(audited_profiles, "blank_category"),
            "audited_invalid_time": _top_hotspots(audited_profiles, "invalid_time"),
        },
    }

    report_root = output_dir / f"year-{year}"
    report_root.mkdir(parents=True, exist_ok=True)

    json_path = report_root / "data_quality_report.json"
    md_path = report_root / "data_quality_report.md"

    json_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    md_path.write_text(_build_markdown(payload), encoding="utf-8")

    return payload, json_path, md_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Analyse source-to-audited data quality for one season.")
    parser.add_argument("--year", type=int, default=1999)
    parser.add_argument("--data-root", type=Path, default=None)
    parser.add_argument("--output-dir", type=Path, default=Path("output") / "quality" / "data-quality")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    data_root = _resolve_data_root(args.data_root)
    if data_root is None:
        print("No data root configured. Set --data-root or configure WRRL data root in app settings.")
        return 1

    try:
        payload, json_path, md_path = analyse_season(args.year, data_root, args.output_dir)
    except FileNotFoundError as exc:
        print(str(exc))
        return 1

    print(f"Wrote: {json_path}")
    print(f"Wrote: {md_path}")
    audited_summary = payload["audited_summary"]
    print(f"Audited blank category %: {audited_summary['blank_category_pct']}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
