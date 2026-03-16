# Journal Tracker OpenAlex Sync

Public-ready Python project for syncing OpenAlex journal records into an Excel/CSV
tracker workbook.

## What is tracked in git

- Source code, tests, CI, and the fixed journal-to-OpenAlex mapping.
- A clean sample workbook at `examples/turkish_politics_articles_database.sample.xlsx`.

## What stays local

- Your real workbook in `data/turkish_politics_articles_database.xlsx`.
- Your OpenAlex key in `.env` or shell environment variables.
- Any generated `.bak.xlsx` backups.

Rotate any API key that has ever been pasted into chats, screenshots, shell
history, or notes before publishing the repo.

## Quick start

macOS / Linux:

```bash
python3 -m venv .venv
.venv/bin/pip install -e ".[dev]"
cp .env.example .env
cp examples/turkish_politics_articles_database.sample.xlsx data/turkish_politics_articles_database.xlsx
```

Windows PowerShell:

```powershell
py -m venv .venv
.\.venv\Scripts\pip install -e ".[dev]"
Copy-Item .env.example .env
Copy-Item examples\turkish_politics_articles_database.sample.xlsx data\turkish_politics_articles_database.xlsx
```

Edit `.env` and set:

```bash
OPENALEX_API_KEY=your-key-here
```

## Usage

Dry run with the local workbook:

```bash
.venv/bin/journal-tracker-sync \
  --workbook data/turkish_politics_articles_database.xlsx \
  --dry-run
```

Windows PowerShell:

```powershell
.\.venv\Scripts\journal-tracker-sync `
  --workbook data\turkish_politics_articles_database.xlsx `
  --dry-run
```

Write new rows into the workbook:

```bash
.venv/bin/journal-tracker-sync \
  --workbook data/turkish_politics_articles_database.xlsx
```

Windows PowerShell:

```powershell
.\.venv\Scripts\journal-tracker-sync `
  --workbook data\turkish_politics_articles_database.xlsx
```

Export the articles sheet as CSV while syncing:

```bash
.venv/bin/journal-tracker-sync \
  --workbook data/turkish_politics_articles_database.xlsx \
  --csv-output data/turkish_politics_articles_database.csv
```

Windows PowerShell:

```powershell
.\.venv\Scripts\journal-tracker-sync `
  --workbook data\turkish_politics_articles_database.xlsx `
  --csv-output data\turkish_politics_articles_database.csv
```

The legacy wrapper still works:

```bash
.venv/bin/python scripts/sync_openalex.py \
  --workbook data/turkish_politics_articles_database.xlsx
```

You can also use the module form on any platform:

```bash
python -m journal_tracker.cli --workbook data/turkish_politics_articles_database.xlsx --dry-run
```

## How it works

- Reads journals from the `Journal Directory` sheet.
- Resolves each journal to a fixed OpenAlex `source_id` from
  `config/openalex_sources.json`.
- Fetches records from the rolling last 3 years using OpenAlex cursor paging.
- Writes `Article Title`, `Author(s)`, `Journal`, `Volume/Issue`, `Year`,
  `Pages`, `DOI/Link`, `Cluster`, and `Key Topics`.
- Can also export the `Turkish Politics Articles` sheet as a CSV file when
  `--csv-output` is provided.
- Preserves sheet styling by cloning the first visible data row style.
- Creates a timestamped `.bak.xlsx` backup before any write.
- Stores OpenAlex work IDs in a hidden metadata sheet so reruns stay idempotent.

## Why these journals

This starter list focuses on journals where Turkish politics articles are often
published, especially across Turkish studies, comparative politics,
democratization, area studies, and political sociology. The goal is to provide
useful coverage out of the box rather than claim that this is the only valid or
complete journal set.

You can absolutely expand the tracker with additional journals that fit your own
research question, region, method, or subfield.

## Add more journals

To add another journal:

1. Add the journal to the `Journal Directory` sheet in your workbook and fill in
   the descriptive columns such as publisher, circle, cluster, quartile, and
   website.
2. Find the journal's OpenAlex source record and note its `source_id`.
3. Add the same journal name and `source_id` to
   `config/openalex_sources.json`. If OpenAlex uses a slightly different journal
   title, also add an `alias`.
4. Run a dry run first to confirm the mapping works:

```bash
.venv/bin/journal-tracker-sync \
  --workbook data/turkish_politics_articles_database.xlsx \
  --dry-run
```

5. If the results look right, run the full sync command.

This makes it easy to adapt the repo for new journals, new subfields, or a
different publication strategy.

## Development

Run checks locally:

```bash
.venv/bin/ruff check .
.venv/bin/pytest
```

Install pre-commit hooks:

```bash
.venv/bin/pre-commit install
```

## Notes

- `Key Topics` prefers OpenAlex `keywords`, then falls back to `topics`.
- Dedupe order is DOI, then OpenAlex work ID, then normalized `title + journal + year`.
- OpenAlex premium-only filters such as `from_created_date` are intentionally not used.
- Re-running the same command next month will scan the rolling 3-year window again and append only unseen records.
