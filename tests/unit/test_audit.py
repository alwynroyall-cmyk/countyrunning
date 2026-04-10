from pathlib import Path

import pandas as pd

from league_scorer.audit import (
    LeagueAuditor,
    _classify_ea_review_category,
    _classify_race_scheme,
    _derived_audit_category,
    _extract_veteran_age,
    _strip_category_sex_prefix,
)


def test_extract_veteran_age_parses_only_veteran_ages():
    assert _extract_veteran_age("V35") == 35
    assert _extract_veteran_age("Vet40") == 40
    assert _extract_veteran_age("50+") == 50
    assert _extract_veteran_age("34") is None
    assert _extract_veteran_age("") is None


def test_classify_race_scheme_returns_ea_for_consecutive_five_year_ages():
    runners = [
        type("Runner", (), {"raw_category": "V45", "gender": "F"}),
        type("Runner", (), {"raw_category": "V50", "gender": "F"}),
    ]
    scheme, evidence = _classify_race_scheme(runners)
    assert scheme == "EA 5-Year"
    assert evidence == "V45, V50"


def test_classify_race_scheme_returns_league_bands_for_ten_year_ages():
    runners = [
        type("Runner", (), {"raw_category": "V40", "gender": "F"}),
        type("Runner", (), {"raw_category": "V60", "gender": "F"}),
    ]
    scheme, evidence = _classify_race_scheme(runners)
    assert scheme == "League Bands"
    assert evidence == "V40, V60"


def test_derived_audit_category_maps_ea_age_ranges_correctly():
    runner = type(
        "Runner",
        (),
        {
            "raw_category": "V55",
            "normalised_category": "V55",
        },
    )
    assert _derived_audit_category(runner, "EA 5-Year") == "V50"
    assert _derived_audit_category(runner, "League Bands") == "V55"


def test_strip_category_sex_prefix_removes_leading_gender_prefix():
    assert _strip_category_sex_prefix("F V45") == "V45"
    assert _strip_category_sex_prefix("Msen") == "sen"
    assert _strip_category_sex_prefix("L40") == "40"
    assert _strip_category_sex_prefix("") == ""


def test_classify_ea_review_category_flags_ambiguous_female_categories():
    candidate = _classify_ea_review_category("F 40-44")
    assert candidate is not None
    assert candidate["normalised_category"] == "40-44"
    assert "ea review" in candidate["reason"].lower()


def test_classify_ea_review_category_ignores_precise_ea_veteran_bands():
    assert _classify_ea_review_category("F V50") is None
    assert _classify_ea_review_category("F 50+") is None


def test_build_actionable_issues_df_includes_row_runner_and_club_issues(tmp_path):
    auditor = LeagueAuditor(tmp_path / "inputs", tmp_path / "outputs", 2025)
    row_df = pd.DataFrame(
        [
            {
                "Severity": "warning",
                "Race": 1,
                "Source Row": 2,
                "Issue Code": "AUD-ROW-001",
                "Runner Name": "Alice",
                "Club": "Club A",
                "Message": "Missing gender",
                "Next Step": "Fix source",
            }
        ]
    )
    runner_df = pd.DataFrame(
        [
            {
                "Severity": "warning",
                "Issue Code": "AUD-RUNNER-007",
                "Runner Key": "alice|collision",
                "Display Name": "Alice",
                "Clubs Seen": "Club A, Club B",
                "Message": "Ambiguous identity",
                "Next Step": "Confirm runner",
            }
        ]
    )
    club_df = pd.DataFrame(
        [
            {
                "Severity": "warning",
                "Issue Code": "AUD-CLUB-002",
                "Raw Club": "Club A",
                "Preferred Club": "Club A",
                "Message": "Inconsistent divisions",
                "Next Step": "Review club lookup",
            }
        ]
    )

    issues_df = auditor._build_actionable_issues_df(row_df, runner_df, club_df)
    assert len(issues_df) == 3
    assert set(issues_df["Type"]) == {"Row", "Runner", "Club"}
    assert "AUD-ROW-001" in issues_df["Issue Code"].values
    assert "AUD-RUNNER-007" in issues_df["Issue Code"].values
    assert "AUD-CLUB-002" in issues_df["Issue Code"].values
