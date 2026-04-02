"""Helpers for identifying and tracking actionable audit issues."""

from __future__ import annotations

from typing import Mapping


def build_issue_identity(issue: Mapping[str, object]) -> str:
    """Build a stable identity for an actionable issue row."""
    issue_type = _norm(issue.get("Type"))
    issue_code = _norm(issue.get("Issue Code"))
    race = _norm(issue.get("Race"))
    source_row = _norm(issue.get("Source Row"))
    key = _norm(issue.get("Key"))
    name = _norm(issue.get("Name"))
    club = _norm(issue.get("Club"))

    if issue_type == "row":
        return "|".join([issue_type, issue_code, race, source_row, name])
    if issue_type == "runner":
        return "|".join([issue_type, issue_code, key, name])
    if issue_type == "club":
        return "|".join([issue_type, issue_code, key, club])
    return "|".join([issue_type, issue_code, race, source_row, key, name, club])


def _norm(value: object) -> str:
    return str(value or "").strip().lower()
