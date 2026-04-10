from pathlib import Path

import pandas as pd
import pytest

from league_scorer.raceroster_import import (
    SporthiveRaceNotDirectlyImportableError,
    _parse_sporthive_table_page,
    import_sporthive_manual_pages,
    import_raceroster_results,
)


def test_parse_sporthive_table_page_reads_pipe_table_rows():
    page_text = """
    | 1 | Alice | 12 | Club A | F35 | 00:42:00 | 00:44:00 |
    | 2 | Bob | 13 | Club B | M40 | 00:43:10 | 00:45:00 |
    """

    rows = _parse_sporthive_table_page(page_text)

    assert len(rows) == 2
    assert rows[0]["Name"] == "Alice"
    assert rows[0]["Club"] == "Club A"
    assert rows[0]["Time"] == "00:44:00"


def test_import_sporthive_manual_pages_writes_excel_and_history(tmp_path: Path):
    input_dir = tmp_path / "inputs"
    input_dir.mkdir()
    pages_text = [
        "| 1 | Alice | 12 | Club A | F35 | 00:42:00 | 00:44:00 |",
        "| 2 | Bob | 13 | Club B | M40 | 00:43:10 | 00:45:00 |",
    ]

    output_path, row_count, history_file = import_sporthive_manual_pages(
        "https://sporthive.com/events/s/123/race/456",
        pages_text,
        input_dir,
        league_race_number=1,
    )

    assert row_count == 2
    assert output_path.exists()
    assert history_file.exists()

    df = pd.read_excel(output_path, engine="openpyxl")
    assert list(df["Name"]) == ["Alice", "Bob"]


def test_import_raceroster_results_rejects_unsupported_url():
    with pytest.raises(ValueError, match="Unsupported URL"):
        import_raceroster_results(
            "https://example.com/results/123",
            Path("/tmp"),
            league_race_number=1,
        )
