from __future__ import annotations


def is_sporthive_event_summary_url(race_url: str) -> bool:
    lower_url = race_url.lower()
    return "sporthive.com/events/s/" in lower_url and "/race/" not in lower_url
