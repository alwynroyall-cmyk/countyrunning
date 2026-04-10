from league_scorer.common_files import race_discovery_exclusions


def test_race_discovery_exclusions_includes_default_names():
    exclusions = race_discovery_exclusions()
    assert "clubs.xlsx" in exclusions
    assert "name_corrections.xlsx" in exclusions
    assert "wrrl_events.xlsx" in exclusions


def test_race_discovery_exclusions_adds_extra_names():
    exclusions = race_discovery_exclusions(["Custom.xlsx", " Clubs.xlsx "])
    assert "custom.xlsx" in exclusions
    assert "clubs.xlsx" in exclusions
