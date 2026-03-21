from __future__ import annotations

from journal_tracker.sync import (
    ArticleIdentifiers,
    JournalDirectoryEntry,
    build_rows,
    format_topics,
    format_volume_issue,
    is_probably_article_url,
    normalize_doi,
    normalize_text,
    resolve_article_identifiers,
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
    assert rows[0][7] == "https://openalex.org/W4"
    assert meta_rows == [
        ("https://openalex.org/W4", "10.1000/new", "brand new article|turkish studies|2025")
    ]


def test_is_probably_article_url_rejects_journal_level_pages() -> None:
    assert not is_probably_article_url("https://www.tandfonline.com/toc/ftur20/current")
    assert not is_probably_article_url(
        "https://www.cambridge.org/core/journals/new-perspectives-on-turkey"
    )
    assert is_probably_article_url("https://publisher.example.org/article/my-paper")


def test_resolve_article_identifiers_uses_crossref_when_openalex_link_is_not_article_level() -> (
    None
):
    entry = JournalDirectoryEntry(
        journal_name="Turkish Studies",
        publisher="Publisher",
        circle="1st Circle",
        cluster="Turkey-dedicated",
        quartile="Q1",
        website="https://example.com",
        source_id="https://openalex.org/S77485876",
    )
    work = {
        "id": "https://openalex.org/W-crossref",
        "display_name": "Recovered DOI Article",
        "publication_year": 2025,
        "doi": None,
        "primary_location": {"landing_page_url": "https://www.tandfonline.com/toc/ftur20/current"},
        "authorships": [{"author": {"display_name": "Author Two"}}],
    }

    identifiers = resolve_article_identifiers(
        work,
        entry,
        crossref_lookup=lambda _work, _entry: ArticleIdentifiers(
            doi_url="https://doi.org/10.1000/recovered",
            article_url="https://publisher.example.org/article/recovered",
        ),
    )

    assert identifiers.doi_url == "https://doi.org/10.1000/recovered"
    assert identifiers.article_url == "https://publisher.example.org/article/recovered"
