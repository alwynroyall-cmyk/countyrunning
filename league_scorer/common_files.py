"""Shared filename conventions for race discovery and lookup workbooks."""

from __future__ import annotations

from typing import Iterable

RACE_NON_RESULT_WORKBOOKS: tuple[str, ...] = (
    "clubs.xlsx",
    "name_corrections.xlsx",
    "wrrl_events.xlsx",
)


def race_discovery_exclusions(extra_names: Iterable[str] = ()) -> tuple[str, ...]:
    """Return normalized workbook names that should be excluded from race discovery."""
    names = {name.strip().lower() for name in RACE_NON_RESULT_WORKBOOKS}
    names.update(name.strip().lower() for name in extra_names if name)
    return tuple(sorted(names))
