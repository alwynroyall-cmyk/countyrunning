import datetime
from pathlib import Path

from league_scorer.normalisation import (
    find_time_column,
    normalise_category,
    normalise_gender,
    parse_time_to_seconds,
    time_display,
)


def test_normalise_gender_handles_known_values():
    assert normalise_gender("M") == "M"
    assert normalise_gender("female") == "F"
    assert normalise_gender("open") == "M"
    assert normalise_gender(None) is None
    assert normalise_gender("unknown") is None


def test_normalise_category_handles_veterans_and_juniors():
    assert normalise_category("V40") == ("V40", "")
    assert normalise_category("fv45") == ("V45", "")
    assert normalise_category("M50+") == ("V50", "")
    assert normalise_category("U20") == ("Jun", "")
    assert normalise_category("Ages 35 - 44") == ("V35", "")
    assert normalise_category("Senior") == ("Sen", "")


def test_normalise_category_preserves_unrecognised():
    preserved, note = normalise_category("Mystery Cat")
    assert preserved == "Mystery Cat"
    assert "Unrecognised" in note


def test_normalise_category_defaults_blank_to_sen():
    assert normalise_category("") == ("Sen", "Missing — defaulted to Sen")
    assert normalise_category(None) == ("Sen", "Missing — defaulted to Sen")


def test_parse_time_to_seconds_accepts_multiple_formats():
    assert parse_time_to_seconds(datetime.time(1, 2, 3)) == 3723.0
    assert parse_time_to_seconds(datetime.timedelta(minutes=3, seconds=30)) == 210.0
    assert parse_time_to_seconds(0.5) == 43200.0
    assert parse_time_to_seconds("1:02:03.400") == 3723.4
    assert parse_time_to_seconds("12:34") == 754.0
    assert parse_time_to_seconds("DNF") is None
    assert parse_time_to_seconds("") is None


def test_time_display_formats_values_consistently():
    assert time_display(datetime.time(2, 3, 4)) == "02:03:04"
    assert time_display(datetime.time(2, 3, 4, 56000)) == "02:03:04.056"
    assert time_display(datetime.timedelta(hours=1, minutes=1, seconds=1)) == "01:01:01"
    assert time_display(0.25) == "06:00:00"
    assert time_display("raw text") == "raw text"


def test_find_time_column_prefers_chip_over_time_net_and_gun():
    columns = ["Name", "Gun Time", "Chip Time", "Finish Time"]
    assert find_time_column(columns) == "Chip Time"

    columns = ["Name", "Gun Time", "Time"]
    assert find_time_column(columns) == "Time"

    columns = ["Name", "Gun Time", "Net Time"]
    assert find_time_column(columns) == "Net Time"

    columns = ["Name", "Gun Time", "Finish Time"]
    assert find_time_column(columns) == "Gun Time"

    columns = ["Name", "Race Time", "Club"]
    assert find_time_column(columns) == "Race Time"

    columns = ["Name", "Club"]
    assert find_time_column(columns) is None
