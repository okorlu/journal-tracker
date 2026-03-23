from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Sequence

from journal_tracker.discover import DiscoverySummary, discover_journals


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Find workbook journals that are missing from config and "
            "suggest OpenAlex sources."
        )
    )
    parser.add_argument(
        "--workbook",
        required=True,
        help="Path to the Excel workbook whose journal directory should be scanned.",
    )
    parser.add_argument(
        "--config",
        help="Optional path to the journal/source mapping JSON file.",
    )
    parser.add_argument(
        "--directory-sheet",
        default="Journal Directory",
        help="Directory sheet to scan for journal names.",
    )
    return parser.parse_args(argv)


def print_summary(summary: DiscoverySummary) -> None:
    print(
        f"Checked {summary.journals_checked} journals against {summary.config_path}",
        flush=True,
    )
    print(
        f"Missing journals: {summary.missing_journals} | "
        f"Suggestion rows written: {summary.suggestion_rows}",
        flush=True,
    )
    print(
        f"Workbook updated: {summary.workbook_path} "
        f"(sheet: {summary.output_sheet})",
        flush=True,
    )


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        summary = discover_journals(
            workbook_path=Path(args.workbook),
            config_path=Path(args.config) if args.config else None,
            directory_sheet_name=args.directory_sheet,
            progress_callback=lambda message: print(message, flush=True),
        )
    except (FileNotFoundError, RuntimeError, ValueError) as exc:
        print(str(exc), file=sys.stderr)
        return 1

    print_summary(summary)
    return 0
