class FatalError(Exception):
    """Fatal error — abort the entire run."""


class RaceProcessingError(Exception):
    """Fatal error at race level — skip this race file and continue."""
