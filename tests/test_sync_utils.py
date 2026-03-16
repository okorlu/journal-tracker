from __future__ import annotations

from journal_tracker.sync import (
    JournalDirectoryEntry,
    build_rows,
    format_topics,
    format_volume_issue,
    normalize_doi,
    normalize_text,
)


def test_normalize_text_removes_accents_and_extra_spacing() -> None:
    assert normalize_text("  İstanbul   Convention  ") == "istanbul convention"


def test_normalize_doi_extracts_from_url() -> None:
    assert (
        normalize_doi("https://doi.org/10.1080/14683849.2025.2458488")
        == "10.1080/14683849.2025.2458488"
    )


def test_format_volume_issue_ignores_missing_parts() -> None:
    assert format_volume_issue({"volume": "24", "issue": "3"}) == "Vol. 24, No. 3"
    assert format_volume_issue({"volume": "24", "issue": None}) == "Vol. 24"


def test_format_topics_prefers_keywords_and_deduplicates() -> None:
    work = {
        "keywords": [
            {"display_name": "Democracy"},
            {"display_name": "Democracy"},
            {"display_name": "Turkey"},
        ],
        "topics": [{"display_name": "Should Not Be Used"}],
    }
    assert format_topics(work) == "Democracy, Turkey"


def test_build_rows_uses_doi_then_work_id_then_normalized_key_for_deduping() -> None:
    entry = JournalDirectoryEntry(
        journal_name="Turkish Studies",
        publisher="Publisher",
        circle="1st Circle",
        cluster="Turkey-dedicated",
        quartile="Q1",
        website="https://example.com",
        source_id="https://openalex.org/S77485876",
    )
    works = [
        {
            "id": "https://openalex.org/W1",
            "display_name": "Existing DOI",
            "publication_year": 2025,
            "publication_date": "2025-01-01",
            "doi": "https://doi.org/10.1000/existing",
            "biblio": {"volume": "1", "issue": "1", "first_page": "1", "last_page": "5"},
            "authorships": [{"author": {"display_name": "Author One"}}],
            "keywords": [{"display_name": "Topic A"}],
        },
        {
            "id": "https://openalex.org/W2",
            "display_name": "Existing Work Id",
            "publication_year": 2025,
            "publication_date": "2025-01-02",
            "doi": None,
            "biblio": {},
            "authorships": [],
            "topics": [{"display_name": "Topic B"}],
        },
        {
            "id": "https://openalex.org/W3",
            "display_name": "Existing   Normalized Title",
            "publication_year": 2025,
            "publication_date": "2025-01-03",
            "doi": None,
            "biblio": {},
            "authorships": [],
            "topics": [{"display_name": "Topic C"}],
        },
        {
            "id": "https://openalex.org/W4",
            "display_name": "Brand New Article",
            "publication_year": 2025,
            "publication_date": "2025-01-04",
            "doi": "https://doi.org/10.1000/new",
            "biblio": {"volume": "2", "issue": "4", "first_page": "12", "last_page": "19"},
            "authorships": [{"author": {"display_name": "Author Two"}}],
            "keywords": [{"display_name": "Topic D"}],
        },
    ]

    rows, meta_rows, duplicate_count = build_rows(
        entry,
        works,
        doi_index={"10.1000/existing"},
        work_id_index={"https://openalex.org/W2"},
        normalized_index={"existing normalized title|turkish studies|2025"},
    )

    assert duplicate_count == 3
    assert len(rows) == 1
    assert rows[0][0] == "Brand New Article"
    assert rows[0][6] == "https://doi.org/10.1000/new"
    assert meta_rows == [
        ("https://openalex.org/W4", "10.1000/new", "brand new article|turkish studies|2025")
    ]
