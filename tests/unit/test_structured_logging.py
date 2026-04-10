import importlib
from pathlib import Path

structured_logging = importlib.import_module("league_scorer.structured_logging")


def test_log_event_writes_jsonl_and_reads_back(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setattr(structured_logging, "_LOG_DIR", tmp_path)
    monkeypatch.setattr(structured_logging, "_LOG_FILE", tmp_path / "wrrl_events.jsonl")
    event = "test_event"
    structured_logging.log_event(event, user="tester", count=1)

    path = tmp_path / "wrrl_events.jsonl"
    assert path.exists()

    events = structured_logging.read_structured_events()
    assert events[0]["event"] == event
    assert events[0]["user"] == "tester"
    assert events[0]["count"] == 1


def test_read_structured_events_ignores_invalid_json(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setattr(structured_logging, "_LOG_FILE", tmp_path / "wrrl_events.jsonl")
    path = tmp_path / "wrrl_events.jsonl"
    path.write_text("not json\n{\"event\": \"ok\"}\n", encoding="utf-8")

    events = structured_logging.read_structured_events()
    assert len(events) == 1
    assert events[0]["event"] == "ok"
