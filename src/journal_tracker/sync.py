from __future__ import annotations

import csv
import json
import re
import shutil
import time
import unicodedata
from copy import copy
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable, Iterable
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from openpyxl import load_workbook

ARTICLES_SHEET = "Articles"
DIRECTORY_SHEET = "Journal Directory"
META_SHEET = "_openalex_sync_meta"
DEFAULT_PER_PAGE = 200
DEFAULT_YEARS = 3
OPENALEX_API_URL = "https://api.openalex.org/works"


@dataclass(frozen=True)
class JournalConfig:
    journal_name: str
    source_id: str
    publisher: str | None
    cluster: str | None
    alias: str | None = None


@dataclass(frozen=True)
class JournalDirectoryEntry:
    journal_name: str
    publisher: str
    circle: str
    cluster: str
    quartile: str
    website: str
    source_id: str
    alias: str | None = None


@dataclass(frozen=True)
class JournalSyncResult:
    journal_name: str
    fetched_count: int
    new_count: int
    duplicate_count: int


@dataclass(frozen=True)
class SyncSummary:
    workbook_path: Path
    cutoff_date: date
    journal_results: list[JournalSyncResult]
    total_fetched: int
    total_new_rows: int
    total_duplicates: int
    dry_run: bool
    backup_path: Path | None = None


def default_config_path() -> Path:
    return Path(__file__).resolve().parents[2] / "config" / "openalex_sources.json"


def load_config(config_path: Path) -> dict[str, JournalConfig]:
    items = json.loads(config_path.read_text(encoding="utf-8"))
    config: dict[str, JournalConfig] = {}
    for item in items:
        config[item["journal_name"]] = JournalConfig(
            journal_name=item["journal_name"],
            source_id=item["source_id"],
            publisher=item.get("publisher"),
            cluster=item.get("cluster"),
            alias=item.get("alias"),
        )
    return config


def openalex_get(url: str) -> dict[str, Any]:
    request = Request(
        url,
        headers={
            "User-Agent": "journal-tracker-openalex-sync/1.0",
            "Accept": "application/json",
        },
    )
    with urlopen(request, timeout=60) as response:
        return json.load(response)


def rolling_cutoff(today: date, years: int) -> date:
    try:
        return today.replace(year=today.year - years)
    except ValueError:
        return today.replace(month=2, day=28, year=today.year - years)


def normalize_text(value: str) -> str:
    collapsed = " ".join((value or "").strip().split())
    ascii_value = unicodedata.normalize("NFKD", collapsed)
    ascii_value = "".join(ch for ch in ascii_value if not unicodedata.combining(ch))
    ascii_value = re.sub(r"[^0-9a-zA-Z]+", " ", ascii_value)
    return " ".join(ascii_value.lower().split())


def normalize_doi(value: str | None) -> str | None:
    if not value:
        return None
    match = re.search(r"10\.\S+", value, re.IGNORECASE)
    if not match:
        return None
    doi = match.group(0).strip().rstrip(" .;,)")
    return doi.lower()


def read_directory_sheet(
    workbook_path: Path,
    config: dict[str, JournalConfig],
) -> list[JournalDirectoryEntry]:
    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    if DIRECTORY_SHEET not in workbook.sheetnames:
        raise ValueError(f"Workbook is missing the '{DIRECTORY_SHEET}' sheet.")

    sheet = workbook[DIRECTORY_SHEET]
    entries: list[JournalDirectoryEntry] = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        journal_name = (row[0] or "").strip()
        if not journal_name:
            continue
        if journal_name not in config:
            raise ValueError(
                f"Journal '{journal_name}' is present in the workbook "
                "but missing from config."
            )
        source = config[journal_name]
        entries.append(
            JournalDirectoryEntry(
                journal_name=journal_name,
                publisher=row[1] or source.publisher or "",
                circle=row[2] or "",
                cluster=row[3] or source.cluster or "",
                quartile=row[4] or "",
                website=row[5] or "",
                source_id=source.source_id,
                alias=source.alias,
            )
        )
    workbook.close()
    return entries


def fetch_works(source_id: str, cutoff_date: date, api_key: str) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []
    cursor = "*"
    while True:
        params = {
            "filter": (
                "primary_location.source.id:"
                f"{source_id},from_publication_date:{cutoff_date.isoformat()}"
            ),
            "per-page": str(DEFAULT_PER_PAGE),
            "cursor": cursor,
            "api_key": api_key,
        }
        url = f"{OPENALEX_API_URL}?{urlencode(params)}"
        try:
            payload = openalex_get(url)
        except HTTPError as exc:
            body = exc.read().decode("utf-8", errors="replace")
            raise RuntimeError(
                f"OpenAlex request failed for {source_id}: HTTP {exc.code} {body}"
            ) from exc
        except URLError as exc:
            raise RuntimeError(f"OpenAlex request failed for {source_id}: {exc}") from exc

        results.extend(payload.get("results", []))
        cursor = payload.get("meta", {}).get("next_cursor")
        if not cursor:
            break
        time.sleep(0.1)
    return results


def format_authors(authorships: Iterable[dict[str, Any]]) -> str:
    names: list[str] = []
    for authorship in authorships:
        author = authorship.get("author") or {}
        name = (author.get("display_name") or authorship.get("raw_author_name") or "").strip()
        if name and name not in names:
            names.append(name)
    return "; ".join(names)


def format_volume_issue(biblio: dict[str, Any]) -> str:
    volume = biblio.get("volume")
    issue = biblio.get("issue")
    parts: list[str] = []
    if volume:
        parts.append(f"Vol. {volume}")
    if issue:
        parts.append(f"No. {issue}")
    return ", ".join(parts)


def format_pages(biblio: dict[str, Any]) -> str:
    first_page = biblio.get("first_page")
    last_page = biblio.get("last_page")
    if first_page and last_page:
        return f"{first_page}-{last_page}"
    return first_page or last_page or ""


def format_topics(work: dict[str, Any]) -> str:
    items: list[str] = []
    source_items = (work.get("keywords") or []) or (work.get("topics") or [])
    for item in source_items:
        label = (item.get("display_name") or "").strip()
        if label and label not in items:
            items.append(label)
        if len(items) == 5:
            break
    return ", ".join(items)


def best_link(work: dict[str, Any]) -> str:
    doi = work.get("doi")
    if doi:
        return doi
    primary_location = work.get("primary_location") or {}
    if primary_location.get("landing_page_url"):
        return primary_location["landing_page_url"]
    return work.get("id") or ""


def normalized_title_key(title: str, journal: str, year: Any) -> str:
    return "|".join(
        [
            normalize_text(title),
            normalize_text(journal),
            normalize_text(str(year or "")),
        ]
    )


def ensure_meta_sheet(workbook) -> Any:
    if META_SHEET in workbook.sheetnames:
        return workbook[META_SHEET]

    meta_sheet = workbook.create_sheet(META_SHEET)
    meta_sheet.sheet_state = "hidden"
    meta_sheet.append(["work_id", "doi", "normalized_key", "synced_at"])
    return meta_sheet


def read_existing_indexes(articles_sheet, meta_sheet) -> tuple[set[str], set[str], set[str]]:
    doi_index: set[str] = set()
    work_id_index: set[str] = set()
    normalized_index: set[str] = set()

    for row in articles_sheet.iter_rows(min_row=2, max_col=7):
        title = row[0].value or ""
        journal = row[2].value or ""
        year = row[4].value or ""
        link = row[6].value or ""
        doi = normalize_doi(link)
        if doi:
            doi_index.add(doi)
        normalized_index.add(normalized_title_key(str(title), str(journal), year))

    for row in meta_sheet.iter_rows(min_row=2, max_col=3, values_only=True):
        work_id = row[0] or ""
        doi = normalize_doi(row[1] or "")
        normalized_key = row[2] or ""
        if work_id:
            work_id_index.add(work_id)
        if doi:
            doi_index.add(doi)
        if normalized_key:
            normalized_index.add(normalized_key)

    return doi_index, work_id_index, normalized_index


def build_rows(
    directory_entry: JournalDirectoryEntry,
    works: list[dict[str, Any]],
    doi_index: set[str],
    work_id_index: set[str],
    normalized_index: set[str],
) -> tuple[list[list[Any]], list[tuple[str, str | None, str]], int]:
    rows: list[list[Any]] = []
    meta_rows: list[tuple[str, str | None, str]] = []
    duplicate_count = 0

    sorted_works = sorted(
        works,
        key=lambda work: (
            work.get("publication_date") or "",
            work.get("publication_year") or 0,
            (work.get("display_name") or "").lower(),
        ),
    )

    for work in sorted_works:
        work_id = work.get("id") or ""
        doi = normalize_doi(work.get("doi"))
        normalized_key = normalized_title_key(
            work.get("display_name") or "",
            directory_entry.journal_name,
            work.get("publication_year") or "",
        )

        if doi and doi in doi_index:
            duplicate_count += 1
            continue
        if work_id and work_id in work_id_index:
            duplicate_count += 1
            continue
        if normalized_key in normalized_index:
            duplicate_count += 1
            continue

        biblio = work.get("biblio") or {}
        rows.append(
            [
                work.get("display_name") or "",
                format_authors(work.get("authorships") or []),
                directory_entry.journal_name,
                format_volume_issue(biblio),
                work.get("publication_year") or "",
                format_pages(biblio),
                best_link(work),
                directory_entry.cluster,
                format_topics(work),
            ]
        )
        meta_rows.append((work_id, doi, normalized_key))
        if doi:
            doi_index.add(doi)
        if work_id:
            work_id_index.add(work_id)
        normalized_index.add(normalized_key)

    return rows, meta_rows, duplicate_count


def clone_row_style(sheet, template_row: int, target_row: int, total_columns: int) -> None:
    for column_index in range(1, total_columns + 1):
        source = sheet.cell(row=template_row, column=column_index)
        target = sheet.cell(row=target_row, column=column_index)
        if source.has_style:
            target._style = copy(source._style)
        if source.number_format:
            target.number_format = source.number_format
        if source.font:
            target.font = copy(source.font)
        if source.fill:
            target.fill = copy(source.fill)
        if source.border:
            target.border = copy(source.border)
        if source.alignment:
            target.alignment = copy(source.alignment)
        if source.protection:
            target.protection = copy(source.protection)
    if sheet.row_dimensions[template_row].height is not None:
        sheet.row_dimensions[target_row].height = sheet.row_dimensions[template_row].height


def append_rows(
    workbook_path: Path,
    rows: list[list[Any]],
    meta_rows: list[tuple[str, str | None, str]],
) -> Path:
    workbook = load_workbook(workbook_path)
    if ARTICLES_SHEET not in workbook.sheetnames:
        raise ValueError(f"Workbook is missing the '{ARTICLES_SHEET}' sheet.")

    articles_sheet = workbook[ARTICLES_SHEET]
    meta_sheet = ensure_meta_sheet(workbook)
    template_row = 2 if articles_sheet.max_row >= 2 else 1
    synced_at = datetime.now().isoformat(timespec="seconds")

    for values, meta in zip(rows, meta_rows):
        next_row = articles_sheet.max_row + 1
        clone_row_style(articles_sheet, template_row, next_row, len(values))
        for column_index, value in enumerate(values, start=1):
            cell = articles_sheet.cell(row=next_row, column=column_index)
            cell.value = value
            if column_index == 7 and value:
                cell.hyperlink = value
        meta_sheet.append([meta[0], meta[1], meta[2], synced_at])

    backup_path = workbook_path.with_name(
        f"{workbook_path.stem}.{datetime.now().strftime('%Y%m%d-%H%M%S')}.bak{workbook_path.suffix}"
    )
    shutil.copy2(workbook_path, backup_path)
    workbook.save(workbook_path)
    workbook.close()
    return backup_path


def export_articles_to_csv(workbook_path: Path, csv_path: Path) -> Path:
    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    if ARTICLES_SHEET not in workbook.sheetnames:
        raise ValueError(f"Workbook is missing the '{ARTICLES_SHEET}' sheet.")

    csv_path = csv_path.expanduser().resolve()
    csv_path.parent.mkdir(parents=True, exist_ok=True)

    articles_sheet = workbook[ARTICLES_SHEET]
    with csv_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        for row in articles_sheet.iter_rows(values_only=True):
            writer.writerow(["" if value is None else value for value in row])

    workbook.close()
    return csv_path


def sync_workbook(
    workbook_path: Path,
    config_path: Path,
    api_key: str,
    years: int = DEFAULT_YEARS,
    dry_run: bool = False,
    today: date | None = None,
    fetcher: Callable[[str, date, str], list[dict[str, Any]]] = fetch_works,
) -> SyncSummary:
    workbook_path = workbook_path.expanduser().resolve()
    config_path = config_path.expanduser().resolve()
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")
    if not config_path.exists():
        raise FileNotFoundError(f"Config not found: {config_path}")

    config = load_config(config_path)
    directory_entries = read_directory_sheet(workbook_path, config)
    cutoff_date = rolling_cutoff(today or date.today(), years)

    workbook = load_workbook(workbook_path)
    if ARTICLES_SHEET not in workbook.sheetnames:
        raise ValueError(f"Workbook is missing the '{ARTICLES_SHEET}' sheet.")
    articles_sheet = workbook[ARTICLES_SHEET]
    meta_sheet = ensure_meta_sheet(workbook)
    doi_index, work_id_index, normalized_index = read_existing_indexes(articles_sheet, meta_sheet)
    workbook.close()

    journal_results: list[JournalSyncResult] = []
    all_new_rows: list[list[Any]] = []
    all_meta_rows: list[tuple[str, str | None, str]] = []
    total_fetched = 0
    total_duplicates = 0

    for entry in directory_entries:
        works = fetcher(entry.source_id, cutoff_date, api_key)
        rows, meta_rows, duplicate_count = build_rows(
            entry,
            works,
            doi_index,
            work_id_index,
            normalized_index,
        )
        journal_results.append(
            JournalSyncResult(
                journal_name=entry.journal_name,
                fetched_count=len(works),
                new_count=len(rows),
                duplicate_count=duplicate_count,
            )
        )
        total_fetched += len(works)
        total_duplicates += duplicate_count
        all_new_rows.extend(rows)
        all_meta_rows.extend(meta_rows)

    backup_path = None
    if not dry_run and all_new_rows:
        backup_path = append_rows(workbook_path, all_new_rows, all_meta_rows)

    return SyncSummary(
        workbook_path=workbook_path,
        cutoff_date=cutoff_date,
        journal_results=journal_results,
        total_fetched=total_fetched,
        total_new_rows=len(all_new_rows),
        total_duplicates=total_duplicates,
        dry_run=dry_run,
        backup_path=backup_path,
    )
