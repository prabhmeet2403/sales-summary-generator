"""
gui
===
Desktop GUI for the Sales Forecast Automation Engine.

This package is purely presentational / orchestration glue. It imports
and calls the existing backend modules (config, excel_reader,
aggregator, comment_mapper, historical_lookup, summary_writer,
validator) exactly as they are -- no calculation or business logic is
duplicated or modified here. See gui/runner.py for the thin wrapper that
adapts main.py's CLI orchestration into a progress-callback-driven,
exception-safe form suitable for a GUI event loop.
"""
