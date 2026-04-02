import pandas as pd

from league_scorer.audit_cleanser import _build_output_frames, _derive_clean_category, _is_race_over_5k


def test_derive_clean_category_maps_added_age_ranges():
    assert _derive_clean_category("Ages 35 - 44", "M")[0] == "V35"
    assert _derive_clean_category("Ages 45 - 54", "F")[0] == "V45"
    assert _derive_clean_category("Ages 55 +", "M")[0] == "V55"
    assert _derive_clean_category("Ages 60 +", "F")[0] == "V60"


def test_derive_clean_category_maps_unknown_to_sen():
    cleaned, note = _derive_clean_category("Unknown", "M")
    assert cleaned == "Sen"
    assert "derived" in note


def test_derive_clean_category_maps_devizes_shorthand_codes():
    assert _derive_clean_category("OS", "M")[0] == "Sen"
    assert _derive_clean_category("FS", "F")[0] == "Sen"
    assert _derive_clean_category("OJ", "M")[0] == "Jun"
    assert _derive_clean_category("FJ", "F")[0] == "Jun"


def test_derive_clean_category_maps_broad_town_shorthand_codes():
    assert _derive_clean_category("SM", "M")[0] == "Sen"
    assert _derive_clean_category("SL", "F")[0] == "Sen"
    assert _derive_clean_category("MS", "M")[0] == "Sen"


def test_derive_clean_category_maps_top3_and_pacer_to_fix():
    assert _derive_clean_category("Top 3", "M")[0] == "FIX"
    assert _derive_clean_category("Top 3 Male", "M")[0] == "FIX"
    assert _derive_clean_category("Top 3 Female", "F")[0] == "FIX"
    assert _derive_clean_category("Pacer", "M")[0] == "FIX"


def test_build_output_frames_drops_wheelchair_non_eligible_club_rows():
    source_df = pd.DataFrame(
        [
            {
                "Name": "Runner One",
                "Club": "Unknown Club",
                "Gender": "M",
                "Category": "Wheelchair Racer",
                "Chip Time": "00:45:00",
            },
            {
                "Name": "Runner Two",
                "Club": "Known Club",
                "Gender": "M",
                "Category": "Wheelchair Racer",
                "Chip Time": "00:44:00",
            },
        ]
    )
    race_df, _, _ = _build_output_frames(
        source_df,
        {
            "name": "Name",
            "club": "Club",
            "gender": "Gender",
            "category": "Category",
        },
        "Chip Time",
        {"known club": "Known Club"},
        ["Known Club"],
        {},
    )

    assert len(race_df) == 1
    assert race_df.iloc[0]["Name"] == "Runner Two"


def test_is_race_over_5k_detects_long_distances_from_name():
    assert _is_race_over_5k("Race 2 - Corsham 10k 2025") is True
    assert _is_race_over_5k("Race 1 - Highworth 5 Mile 2025") is True
    assert _is_race_over_5k("Race 7 - Chippenham Half Marathon 2025") is True
    assert _is_race_over_5k("Race 3 - Westbury 5k") is False


def test_build_output_frames_long_race_prefers_excel_time_values():
    source_df = pd.DataFrame(
        [
            {
                "Name": "Runner Time",
                "Club": "Known Club",
                "Gender": "M",
                "Category": "Sen",
                "Chip Time": "01:02:30",
            }
        ]
    )

    race_df, _, _ = _build_output_frames(
        source_df,
        {
            "name": "Name",
            "club": "Club",
            "gender": "Gender",
            "category": "Category",
        },
        "Chip Time",
        {"known club": "Known Club"},
        ["Known Club"],
        {},
        prefer_excel_time_format=True,
    )

    time_value = race_df.iloc[0]["Time"]
    assert isinstance(time_value, float)
    assert abs(time_value - (3750.0 / 86400.0)) < 1e-9
