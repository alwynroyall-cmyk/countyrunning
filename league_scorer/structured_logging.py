"""Structured JSONL logging helpers for operational diagnostics."""

from __future__ import annotations

import json
import logging
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

_LOG_DIR = Path.home() / ".wrrl_logs"
_LOG_FILE = _LOG_DIR / "wrrl_events.jsonl"


def structured_log_path() -> Path:
    """Return the JSONL event log path, creating the directory if needed."""
    _LOG_DIR.mkdir(parents=True, exist_ok=True)
    return _LOG_FILE


def _serialise(value: Any) -> Any:
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, Exception):
        return str(value)
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    if isinstance(value, dict):
        return {str(k): _serialise(v) for k, v in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [_serialise(v) for v in value]
    return str(value)


def log_event(event: str, *, level: str = "INFO", logger: logging.Logger | None = None, **fields: Any) -> None:
    """Write one structured event to JSONL and mirror a concise line to std logging."""
    payload = {
        "ts": datetime.now(timezone.utc).isoformat(),
        "level": level.upper(),
        "event": event,
    }
    payload.update({k: _serialise(v) for k, v in fields.items()})

    try:
        path = structured_log_path()
        with path.open("a", encoding="utf-8") as fh:
            fh.write(json.dumps(payload, ensure_ascii=True) + "\n")
    except Exception:
        # Never allow structured logging failures to impact app workflows.
        pass

    log_target = logger or logging.getLogger(__name__)
    level_no = getattr(logging, payload["level"], logging.INFO)
    details = " ".join(f"{k}={payload[k]}" for k in sorted(payload) if k not in {"ts", "level", "event"})
    message = f"event={event}" if not details else f"event={event} {details}"
    log_target.log(level_no, message)


def read_structured_events(limit: int = 500) -> list[dict[str, Any]]:
    """Read recent structured events from JSONL, newest-first."""
    path = _LOG_FILE
    if not path.exists():
        return []

    rows: list[dict[str, Any]] = []
    try:
        with path.open("r", encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if not line:
                    continue
                try:
                    data = json.loads(line)
                except json.JSONDecodeError:
                    continue
                if isinstance(data, dict):
                    rows.append(data)
    except OSError:
        return []

    rows.reverse()
    return rows[:limit]
