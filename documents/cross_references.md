# Cross-References and Dependencies in the Project

## Modules with Dependencies
| Module                          | Dependencies                                                                 |
|---------------------------------|------------------------------------------------------------------------------|
| league_scorer                  | league_scorer.config, league_scorer.input, league_scorer.output, league_scorer.process |
| league_scorer.config           | league_scorer.config.session_config, league_scorer.config.settings           |
| league_scorer.input            | league_scorer.input.club_loader, league_scorer.input.common_files, league_scorer.input.events_loader, league_scorer.input.input_layout, league_scorer.input.raceroster_import, league_scorer.input.source_loader |
| league_scorer.output           | league_scorer.output.output_layout, league_scorer.output.output_writer, league_scorer.output.report_writer, league_scorer.output.structured_logging |
| league_scorer.process          | league_scorer.process.individual_scoring, league_scorer.process.main, league_scorer.process.models, league_scorer.process.name_lookup, league_scorer.process.normalisation, league_scorer.process.race_processor, league_scorer.process.race_validation, league_scorer.process.rules, league_scorer.process.season_aggregation, league_scorer.process.team_scoring |
| league_scorer.publish          | league_scorer.publish.publish_results                                        |
| league_scorer.graphical        | league_scorer.graphical.qt.launch_dashboard                                  |
| league_scorer.graphical.qt     | league_scorer.graphical.qt.dashboard.launch_dashboard                        |

## Modules with No Dependencies
| Module                          |
|---------------------------------|
| league_scorer.input.__init__    |
| league_scorer.output.__init__   |
| league_scorer.process.__init__  |
| league_scorer.publish.__init__  |
| league_scorer.graphical.__init__|
| league_scorer.graphical.qt.__init__|

This table summarizes the cross-references and highlights modules with no dependencies.