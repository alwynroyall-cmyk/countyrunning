from league_scorer.issue_tracking import build_issue_identity


def test_build_issue_identity_for_row_issue():
    issue = {
        "Type": "Row",
        "Issue Code": "AUD-ROW-001",
        "Race": "1",
        "Source Row": "4",
        "Name": "Alice",
    }
    assert build_issue_identity(issue) == "row|aud-row-001|1|4|alice"


def test_build_issue_identity_for_runner_issue():
    issue = {
        "Type": "Runner",
        "Issue Code": "AUD-RUNNER-007",
        "Key": "alice|collision",
        "Name": "Alice",
    }
    assert build_issue_identity(issue) == "runner|aud-runner-007|alice|collision|alice"


def test_build_issue_identity_for_club_issue():
    issue = {
        "Type": "Club",
        "Issue Code": "AUD-CLUB-002",
        "Key": "Unknown Club",
        "Club": "Club A",
    }
    assert build_issue_identity(issue) == "club|aud-club-002|unknown club|club a"


def test_build_issue_identity_handles_missing_values():
    issue = {}
    assert build_issue_identity(issue) == "||||||"
