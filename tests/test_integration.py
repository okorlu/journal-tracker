from __future__ import annotations

import csv
import shutil
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

from journal_tracker.sync import (
    META_SHEET,
    default_config_path,
    export_articles_to_csv,
    sync_workbook,
)


def test_sync_workbook_is_idempotent_with_sample_workbook(tmp_path: Path) -> None:
    source_workbook = Path("examples/turkish_politics_articles_database.sample.xlsx")
    workbook_path = tmp_path / "tracker.xlsx"
    shutil.copy2(source_workbook, workbook_path)

    payloads = {
        "https://openalex.org/S77485876": [
            {
                "id": "https://openalex.org/W-existing",
                "display_name": "Sample Existing Article",
                "publication_year": 2025,
                "publication_date": "2025-02-01",
                "doi": "https://doi.org/10.5555/sample-existing",
                "biblio": {"volume": "25", "issue": "1", "first_page": "1", "last_page": "12"},
                "authorships": [{"author": {"display_name": "Aylin Demo"}}],
                "keywords": [{"display_name": "Existing Topic"}],
            },
            {
                "id": "https://openalex.org/W-new-1",
                "display_name": "Fresh Turkish Studies Article",
                "publication_year": 2026,
                "publication_date": "2026-03-01",
                "doi": "https://doi.org/10.5555/fresh-ts",
                "biblio": {"volume": "26", "issue": "1", "first_page": "13", "last_page": "28"},
                "authorships": [{"author": {"display_name": "Mert Example"}}],
                "keywords": [{"display_name": "Political science"}],
            },
        ],
        "https://openalex.org/S6959732": [
            {
                "id": "https://openalex.org/W-new-2",
                "display_name": "Fresh Party Politics Article",
                "publication_year": 2026,
                "publication_date": "2026-03-02",
                "doi": None,
                "primary_location": {"landing_page_url": "https://example.org/party-politics"},
                "biblio": {"volume": "32", "issue": "2", "first_page": "44", "last_page": "58"},
                "authorships": [{"author": {"display_name": "Deniz Example"}}],
                "topics": [{"display_name": "Party systems"}],
            }
        ],
    }

    def fake_fetcher(source_id: str, cutoff_date: date, api_key: str):
        assert api_key == "test-key"
        assert cutoff_date == date(2023, 3, 15)
        return payloads.get(source_id, [])

    summary = sync_workbook(
        workbook_path=workbook_path,
        config_path=default_config_path(),
        api_key="test-key",
        years=3,
        dry_run=False,
        today=date(2026, 3, 15),
        fetcher=fake_fetcher,
    )

    assert summary.total_new_rows == 2
    assert summary.total_duplicates == 1
    assert summary.backup_path is not None
    assert summary.backup_path.exists()

    workbook = load_workbook(workbook_path)
    sheet = workbook["Articles"]
    assert sheet.max_row == 5
    assert sheet["A4"].value == "Fresh Turkish Studies Article"
    assert sheet["A5"].value == "Fresh Party Politics Article"
    assert workbook[META_SHEET].sheet_state == "hidden"
    workbook.close()

    second_summary = sync_workbook(
        workbook_path=workbook_path,
        config_path=default_config_path(),
        api_key="test-key",
        years=3,
        dry_run=False,
        today=date(2026, 3, 15),
        fetcher=fake_fetcher,
    )

    assert second_summary.total_new_rows == 0
    assert second_summary.backup_path is None


def test_export_articles_to_csv_writes_current_articles_sheet(tmp_path: Path) -> None:
    source_workbook = Path("examples/turkish_politics_articles_database.sample.xlsx")
    workbook_path = tmp_path / "tracker.xlsx"
    csv_path = tmp_path / "exports" / "tracker.csv"
    shutil.copy2(source_workbook, workbook_path)

    export_articles_to_csv(workbook_path, csv_path)

    assert csv_path.exists()
    with csv_path.open(encoding="utf-8", newline="") as handle:
        rows = list(csv.reader(handle))

    assert rows[0] == [
        "Article Title",
        "Author(s)",
        "Journal",
        "Volume/Issue",
        "Year",
        "Pages",
        "DOI/Link",
        "Cluster",
        "Key Topics",
    ]
    assert rows[1][0] == "Sample Existing Article"
