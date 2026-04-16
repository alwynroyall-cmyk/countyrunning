"""WRRL Admin Suite v8.2.3"""

# Central version constant for all modules
__version__ = "8.2.3"

# Backward-compatible root imports for legacy consumers.
from .config import session_config
from .config.settings import settings
from .input import club_loader, common_files, events_loader, input_layout, raceroster_import, source_loader
from .output import output_layout, output_writer, report_writer, structured_logging
from .process import individual_scoring, main, models, name_lookup, normalisation, race_processor, race_validation, rules, season_aggregation, team_scoring

__all__ = [
    "__version__",
    "session_config",
    "settings",
    "club_loader",
    "common_files",
    "events_loader",
    "input_layout",
    "raceroster_import",
    "source_loader",
    "output_layout",
    "output_writer",
    "report_writer",
    "structured_logging",
    "individual_scoring",
    "main",
    "models",
    "name_lookup",
    "normalisation",
    "race_processor",
    "race_validation",
    "rules",
    "season_aggregation",
    "team_scoring",
]
