"""Compatibility shim for league_scorer.process.rules."""
from league_scorer import rules as _source

for _name in dir(_source):
    if not _name.startswith("_"):
        globals()[_name] = getattr(_source, _name)

__all__ = getattr(_source, "__all__", [name for name in globals() if not name.startswith("_")])
