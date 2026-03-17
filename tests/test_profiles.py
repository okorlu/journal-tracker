from __future__ import annotations

import json
from pathlib import Path

from journal_tracker.profiles import load_profile
from journal_tracker.sync import ARTICLES_SHEET, DIRECTORY_SHEET


def test_load_profile_resolves_relative_paths_and_deduplicates_journals(tmp_path: Path) -> None:
    profile_dir = tmp_path / "profiles"
    profile_dir.mkdir()
    workbook_path = tmp_path / "data" / "tracker.xlsx"
    config_path = tmp_path / "config" / "sources.json"
    csv_path = tmp_path / "exports" / "tracker.csv"
    workbook_path.parent.mkdir()
    config_path.parent.mkdir()
    csv_path.parent.mkdir()

    profile_path = profile_dir / "sample.json"
    profile_path.write_text(
        json.dumps(
            {
                "name": "demo-profile",
                "workbook": "../data/tracker.xlsx",
                "config": "../config/sources.json",
                "csv_output": "../exports/tracker.csv",
                "years": 5,
                "articles_sheet": "Custom Articles",
                "directory_sheet": "Custom Directory",
                "journals": ["Journal A", "Journal B", "Journal A"],
            }
        ),
        encoding="utf-8",
    )

    profile = load_profile(profile_path)

    assert profile.name == "demo-profile"
    assert profile.workbook_path == workbook_path.resolve()
    assert profile.config_path == config_path.resolve()
    assert profile.csv_output_path == csv_path.resolve()
    assert profile.years == 5
    assert profile.articles_sheet == "Custom Articles"
    assert profile.directory_sheet == "Custom Directory"
    assert profile.journal_names == ("Journal A", "Journal B")


def test_load_profile_uses_default_sheets_and_name(tmp_path: Path) -> None:
    profile_path = tmp_path / "starter.json"
    profile_path.write_text("{}", encoding="utf-8")

    profile = load_profile(profile_path)

    assert profile.name == "starter"
    assert profile.articles_sheet == ARTICLES_SHEET
    assert profile.directory_sheet == DIRECTORY_SHEET
    assert profile.journal_names == ()
