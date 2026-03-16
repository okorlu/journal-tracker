# Contributing

## Local setup

```bash
python3 -m venv .venv
.venv/bin/pip install -e ".[dev]"
cp .env.example .env
```

Add your own `OPENALEX_API_KEY` to `.env`. Do not commit `.env` or local workbooks.

## Run checks

```bash
.venv/bin/ruff check .
.venv/bin/pytest
```

## Pre-commit

```bash
.venv/bin/pre-commit install
```

## Data handling

- Keep real workbooks under `data/`.
- Use `examples/turkish_politics_articles_database.sample.xlsx` as a starter example for tests, demos, and docs.
- Rotate any API key that was ever pasted into chats, screenshots, or command history before publishing the repo.
