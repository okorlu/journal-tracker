from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path
from typing import Sequence

from journal_tracker.profiles import TrackingProfile, load_profile
from journal_tracker.sync import (
    DEFAULT_YEARS,
    SyncSummary,
    default_config_path,
    export_articles_to_csv,
    sync_workbook,
)


def resolve_cli_path(value: str) -> Path:
    return Path(value).expanduser().resolve()


def positive_int(value: str) -> int:
    try:
        parsed = int(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("must be a positive integer") from exc
    if parsed <= 0:
        raise argparse.ArgumentTypeError("must be a positive integer")
    return parsed


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Sync OpenAlex publications into the journal tracker workbook."
    )
    parser.add_argument(
        "--profile",
        help="Optional path to a tracking profile JSON file.",
    )
    parser.add_argument(
        "--workbook",
        help="Path to the Excel workbook that will be updated in place.",
    )
    parser.add_argument(
        "--years",
        type=positive_int,
        help=f"Rolling publication window in years (default: {DEFAULT_YEARS}).",
    )
    parser.add_argument(
        "--api-key",
        help="OpenAlex API key. Falls back to OPENALEX_API_KEY or a local .env file.",
    )
    parser.add_argument(
        "--config",
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


def resolve_run_options(
    args: argparse.Namespace,
) -> tuple[TrackingProfile | None, dict[str, object]]:
    profile = load_profile(Path(args.profile)) if args.profile else None

    workbook_path = (
        resolve_cli_path(args.workbook)
        if args.workbook
        else profile.workbook_path
        if profile
        else None
    )
    if workbook_path is None:
        raise ValueError("Missing workbook path. Use --workbook or provide it in --profile.")

    config_path = (
        resolve_cli_path(args.config)
        if args.config
        else profile.config_path
        if profile and profile.config_path
        else default_config_path()
    )
    csv_output_path = (
        resolve_cli_path(args.csv_output)
        if args.csv_output
        else profile.csv_output_path
        if profile
        else None
    )
    options: dict[str, object] = {
        "workbook_path": workbook_path,
        "config_path": config_path,
        "years": (
            args.years if args.years is not None else profile.years if profile else DEFAULT_YEARS
        ),
        "dry_run": args.dry_run,
        "articles_sheet": profile.articles_sheet if profile else "Articles",
        "directory_sheet": profile.directory_sheet if profile else "Journal Directory",
        "journal_names": profile.journal_names if profile and profile.journal_names else None,
        "csv_output_path": csv_output_path,
    }
    return profile, options


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


def env_file_candidates(
    current_dir: Path,
    profile: TrackingProfile | None,
    workbook_path: Path,
) -> tuple[Path, ...]:
    repo_root = Path(__file__).resolve().parents[2]
    candidates = [
        profile.profile_path.parent / ".env" if profile else None,
        workbook_path.parent / ".env",
        repo_root / ".env",
        current_dir / ".env",
    ]
    seen: set[Path] = set()
    ordered: list[Path] = []
    for candidate in candidates:
        if candidate is None:
            continue
        resolved = candidate.expanduser().resolve()
        if resolved not in seen:
            seen.add(resolved)
            ordered.append(resolved)
    return tuple(ordered)


def load_runtime_env(profile: TrackingProfile | None, workbook_path: Path) -> None:
    for dotenv_path in env_file_candidates(Path.cwd(), profile, workbook_path):
        load_env_file(dotenv_path)


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
    if (
        summary.identifier_columns_migrated
        or summary.added_at_column_added
        or summary.added_at_backfilled
    ):
        migration_bits: list[str] = []
        if summary.identifier_columns_migrated:
            migration_bits.append("DOI/Article URL columns migrated")
        if summary.added_at_column_added:
            migration_bits.append("Added At column added")
        if summary.added_at_backfilled:
            migration_bits.append(f"Added At backfilled for {summary.added_at_backfilled} rows")
        print("Workbook migration: " + "; ".join(migration_bits))
    if summary.dry_run:
        print("Dry run complete. Workbook was not modified.")
    elif summary.total_new_rows == 0 and not summary.workbook_changed:
        print("No new rows found. Workbook was not modified.")
    else:
        print(f"Backup created at: {summary.backup_path}")
        print(f"Workbook updated: {summary.workbook_path}")


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        profile, options = resolve_run_options(args)
    except (FileNotFoundError, ValueError) as exc:
        print(str(exc), file=sys.stderr)
        return 1

    load_runtime_env(profile, options["workbook_path"])
    api_key = args.api_key or os.getenv("OPENALEX_API_KEY")

    if not api_key:
        print("Missing OpenAlex API key. Use --api-key or set OPENALEX_API_KEY.", file=sys.stderr)
        return 1

    print("Starting journal sync...", flush=True)
    if args.profile:
        print(f"Using profile: {Path(args.profile).expanduser()}", flush=True)

    summary = sync_workbook(
        workbook_path=options["workbook_path"],
        config_path=options["config_path"],
        api_key=api_key,
        years=options["years"],
        dry_run=options["dry_run"],
        articles_sheet=options["articles_sheet"],
        directory_sheet=options["directory_sheet"],
        journal_names=options["journal_names"],
        progress_callback=lambda message: print(message, flush=True),
        crossref_mailto=os.getenv("CROSSREF_MAILTO"),
    )
    print_summary(summary)
    if options["csv_output_path"]:
        csv_path = export_articles_to_csv(
            options["workbook_path"],
            options["csv_output_path"],
            articles_sheet_name=options["articles_sheet"],
        )
        print(f"CSV exported: {csv_path}")
    return 0
