# Changelog

All notable changes to this project will be documented in this file.

The format is inspired by Keep a Changelog, with entries grouped by user-facing
impact.

## [Unreleased]

### Added

- Added configurable tracking profiles so sync runs can reuse saved defaults
  for workbook path, config path, year window, sheet names, journal subsets,
  and optional CSV output.
- Added live sync progress output in the CLI, including startup messages,
  journal-by-journal progress, page-level fetch progress for larger OpenAlex
  result sets, and write-stage updates.
- Added a visible `Added At` column to the `Articles` sheet so newly synced
  records are easier to identify.

### Changed

- Split the old combined `DOI/Link` field into separate `DOI` and
  `Article URL` columns.
- Tightened article link handling so journal home pages, table-of-contents
  pages, and similar non-article destinations are no longer treated as valid
  article URLs.
- Added a conservative Crossref fallback for DOI enrichment when OpenAlex does
  not provide a DOI directly.
- Improved workbook migration behavior so existing trackers can be upgraded to
  the new identifier columns and `Added At` column without losing historical
  metadata.
- Improved rebuild behavior so a workbook with a blank styled template row can
  be repopulated cleanly without leaving an empty row behind.

### Fixed

- Fixed a major data-quality issue where some rows stored journal-level landing
  pages instead of article-level links.
- Fixed long-running sync runs appearing idle in the terminal by surfacing
  incremental progress as the job proceeds.

### Validation

- Verified with `ruff check`.
- Verified with `pytest` (`15 passed`).
- Rebuilt the local tracker workbook from scratch using the updated schema to
  confirm the new DOI, article URL, and added-at behavior end to end.

### Commits

- `c28cd28` Add configurable tracking profiles
- `4f113ae` Add added-at tracking and live sync progress
- `c2d3340` Split DOI and article URL with Crossref fallback
- `a4b0972` Reuse blank template row during rebuilds

## [0.1.0] - 2026-03-15

### Added

- Initial public release of `journal-tracker`.
- OpenAlex-powered sync into Excel workbooks.
- CSV export support.
- Sample workbook, tests, CI, and public project documentation.
