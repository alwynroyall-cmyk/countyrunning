"""Publish helpers package for league_scorer.

Expose the `publish_results` function for scripts and the GUI to call.
"""
from .publish import publish_results

__all__ = ["publish_results"]
