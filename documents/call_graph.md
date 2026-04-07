# Application Call Graph: CountyRunning

## Overview

This document describes the application's call structure, entry points, and a detailed call graph (text + mermaid) showing how control and data flow through the codebase.

## Entry points

- [__main__.py](__main__.py): CLI entry that typically dispatches to `league_scorer/main.py` for batch runs.
- [run_gui.py](run_gui.py): GUI launcher that initializes the graphical interface in `league_scorer/graphical/gui.py`.
- `scripts/` (e.g., `scripts/run_full_autopilot.py`): specialized automation scripts that call core modules directly.

## High-level flow

1. Startup: invoked via CLI (`__main__.py`) or GUI (`run_gui.py`) or a script.
2. Configuration: load `league_scorer/settings.py` and `league_scorer/session_config.py`.
3. Ingest: data loaders (`league_scorer/club_loader.py`, `league_scorer/events_loader.py`, `league_scorer/source_loader.py`).
4. Process: core processing (`league_scorer/race_processor.py`, `league_scorer/individual_scoring.py`, `league_scorer/team_scoring.py`).
5. Audit/Validation: audit modules (`league_scorer/audit_cleanser.py`, `league_scorer/audit_data_service.py`, `league_scorer/issue_tracking.py`).
6. Output: writers and reports (`league_scorer/output_writer.py`, `league_scorer/report_writer.py`, `league_scorer/output_layout.py`).
7. GUI: GUI modules call backend services for data and actions (see `league_scorer/graphical/`).

## Mermaid call graph

```mermaid
flowchart LR
  subgraph CLI
    A[__main__.py]
    S[scripts/*]
  end
  subgraph GUI
    G[run_gui.py] --> GUI[league_scorer/graphical/gui.py]
    GUI --> Panels[graphical panels]
  end

  A --> M[league_scorer/main.py]
  S --> M
  Panels --> M

  M --> Cfg[league_scorer/settings.py\nsession_config.py]
  M --> Loaders[Data loaders]
  Loaders --> Club[league_scorer/club_loader.py]
  Loaders --> Events[league_scorer/events_loader.py]
  Loaders --> Source[league_scorer/source_loader.py]

  M --> Processor[league_scorer/race_processor.py]
  Processor --> Indiv[league_scorer/individual_scoring.py]
  Processor --> Team[league_scorer/team_scoring.py]

  Processor --> Audit[Audit & validation]
  Audit --> Cleanser[league_scorer/audit_cleanser.py]
  Audit --> DataSvc[league_scorer/audit_data_service.py]
  Audit --> Issues[league_scorer/issue_tracking.py]

  M --> Output[Output writers]
  Output --> OutWriter[league_scorer/output_writer.py]
  Output --> Report[league_scorer/report_writer.py]
  Output --> Layout[league_scorer/output_layout.py]

  style A fill:#f9f,stroke:#333,stroke-width:1px
  style G fill:#bbf,stroke:#333,stroke-width:1px
  style M fill:#bfb,stroke:#333,stroke-width:1px
```

## Detailed textual call graph

- `__main__.py` -> `league_scorer/main.py` ->
  - load settings (`settings.py`, `session_config.py`)
  - call data loaders (`club_loader`, `events_loader`, `source_loader`)
  - invoke `race_processor.py`:
    - `individual_scoring.py`
    - `team_scoring.py`
    - aggregation modules (`season_aggregation.py`, `series_consolidation.py`)
  - call audit modules (`audit_cleanser.py`, `audit_data_service.py`, `issue_tracking.py`)
  - generate outputs (`output_writer.py`, `report_writer.py`)

- `run_gui.py` -> `league_scorer/graphical/gui.py` -> various panels/dialogs:
  - `results_viewer.py`, `club_editor.py`, `audit_gui.py`, etc.
  - GUI panels call backend service functions in `league_scorer/` (loaders, processor, audit, writers) to perform actions and update views.

- `scripts/*.py` -> either call `league_scorer/main.py` or directly instantiate loader/processor/writer classes for specialized batch workflows.

## Notes & next steps

- For a function-level call graph, run a static analysis (AST) or use runtime tracing to extract exact call relations.
- If you want, I can generate a function-level graph for a selected module (e.g., `league_scorer/race_processor.py`).

---

File created: docs/call_graph.md
