from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path
from typing import Sequence

from journal_tracker.sync import (
    DEFAULT_YEARS,
    SyncSummary,
    default_config_path,
    export_articles_to_csv,
    sync_workbook,
)


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Sync OpenAlex publications into the journal tracker workbook."
    )
    parser.add_argument(
        "--workbook",
        required=True,
        help="Path to the Excel workbook that will be updated in place.",
    )
    parser.add_argument(
        "--years",
        type=int,
        default=DEFAULT_YEARS,
        help=f"Rolling publication window in years (default: {DEFAULT_YEARS}).",
    )
    parser.add_argument(
        "--api-key",
        help="OpenAlex API key. Falls back to OPENALEX_API_KEY or a local .env file.",
    )
    parser.add_argument(
        "--config",
        default=str(default_config_path()),
        help="Path to the journal/source mapping JSON file.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Fetch and compare records without writing to the workbook.",
    )
    parser.add_argument(
        "--csv-output",
        help="Optional path to export the 'Articles' sheet as CSV.",
    )
    return parser.parse_args(argv)


def load_env_file(dotenv_path: Path) -> None:
    if not dotenv_path.exists():
        return

    for raw_line in dotenv_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip("'").strip('"')
        os.environ.setdefault(key, value)


def print_summary(summary: SyncSummary) -> None:
    mode = "dry-run" if summary.dry_run else "write mode"
    print(
        f"Syncing {len(summary.journal_results)} journals from "
        f"{summary.cutoff_date.isoformat()} onward ({mode})"
    )
    for result in summary.journal_results:
        print(
            f"- {result.journal_name}: fetched={result.fetched_count} "
            f"new={result.new_count} duplicates={result.duplicate_count}"
        )
    print(
        f"Summary: journals={len(summary.journal_results)} fetched={summary.total_fetched} "
        f"new_rows={summary.total_new_rows} duplicates={summary.total_duplicates}"
    )
    if summary.dry_run:
        print("Dry run complete. Workbook was not modified.")
    elif summary.total_new_rows == 0:
        print("No new rows found. Workbook was not modified.")
    else:
        print(f"Backup created at: {summary.backup_path}")
        print(f"Workbook updated: {summary.workbook_path}")


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    load_env_file(Path(".env"))
    api_key = args.api_key or os.getenv("OPENALEX_API_KEY")

    if not api_key:
        print("Missing OpenAlex API key. Use --api-key or set OPENALEX_API_KEY.", file=sys.stderr)
        return 1

    summary = sync_workbook(
        workbook_path=Path(args.workbook),
        config_path=Path(args.config),
        api_key=api_key,
        years=args.years,
        dry_run=args.dry_run,
    )
    print_summary(summary)
    if args.csv_output:
        csv_path = export_articles_to_csv(Path(args.workbook), Path(args.csv_output))
        print(f"CSV exported: {csv_path}")
    return 0
