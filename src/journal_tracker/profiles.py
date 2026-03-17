from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path

from journal_tracker.sync import ARTICLES_SHEET, DIRECTORY_SHEET, default_config_path


@dataclass(frozen=True)
class TrackingProfile:
    name: str
    profile_path: Path
    workbook_path: Path | None = None
    config_path: Path | None = None
    csv_output_path: Path | None = None
    years: int | None = None
    articles_sheet: str = ARTICLES_SHEET
    directory_sheet: str = DIRECTORY_SHEET
    journal_names: tuple[str, ...] = ()


def default_profiles_dir() -> Path:
    return Path(__file__).resolve().parents[2] / "config" / "profiles"


def _resolve_optional_path(base_dir: Path, value: str | None) -> Path | None:
    if not value:
        return None
    path = Path(value).expanduser()
    if not path.is_absolute():
        path = (base_dir / path).resolve()
    return path


def _dedupe_preserving_order(items: list[str]) -> tuple[str, ...]:
    seen: set[str] = set()
    values: list[str] = []
    for item in items:
        normalized = item.strip()
        if normalized and normalized not in seen:
            seen.add(normalized)
            values.append(normalized)
    return tuple(values)


def load_profile(profile_path: Path) -> TrackingProfile:
    resolved_path = profile_path.expanduser().resolve()
    payload = json.loads(resolved_path.read_text(encoding="utf-8"))
    base_dir = resolved_path.parent
    journal_names = _dedupe_preserving_order(payload.get("journals") or [])

    return TrackingProfile(
        name=(payload.get("name") or resolved_path.stem).strip(),
        profile_path=resolved_path,
        workbook_path=_resolve_optional_path(base_dir, payload.get("workbook")),
        config_path=_resolve_optional_path(base_dir, payload.get("config"))
        or default_config_path(),
        csv_output_path=_resolve_optional_path(base_dir, payload.get("csv_output")),
        years=payload.get("years"),
        articles_sheet=(payload.get("articles_sheet") or ARTICLES_SHEET).strip(),
        directory_sheet=(payload.get("directory_sheet") or DIRECTORY_SHEET).strip(),
        journal_names=journal_names,
    )
