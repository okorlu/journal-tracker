from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path

from journal_tracker.sync import default_config_path


@dataclass(frozen=True)
class TrackingProfile:
    name: str
    profile_path: Path
    workbook_path: Path | None = None
    config_path: Path | None = None
    csv_output_path: Path | None = None
    years: int | None = None
    articles_sheet: str | None = None
    directory_sheet: str | None = None
    journal_names: tuple[str, ...] = ()


def default_profiles_dir() -> Path:
    return Path(__file__).resolve().parents[2] / "config" / "profiles"


def _expect_object(payload: object, profile_path: Path) -> dict[str, object]:
    if not isinstance(payload, dict):
        raise ValueError(f"Profile '{profile_path}' must contain a JSON object.")
    return payload


def _optional_string(
    payload: dict[str, object],
    key: str,
    profile_path: Path,
    *,
    allow_blank: bool = False,
) -> str | None:
    value = payload.get(key)
    if value is None:
        return None
    if not isinstance(value, str):
        raise ValueError(f"Profile '{profile_path}' field '{key}' must be a string.")
    normalized = value.strip()
    if not normalized and not allow_blank:
        raise ValueError(f"Profile '{profile_path}' field '{key}' cannot be blank.")
    return normalized


def _optional_years(payload: dict[str, object], profile_path: Path) -> int | None:
    value = payload.get("years")
    if value is None:
        return None
    if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
        raise ValueError(f"Profile '{profile_path}' field 'years' must be a positive integer.")
    return value


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
    payload = _expect_object(json.loads(resolved_path.read_text(encoding="utf-8")), resolved_path)
    base_dir = resolved_path.parent
    raw_journals = payload.get("journals") or []
    if not isinstance(raw_journals, list) or any(
        not isinstance(item, str) for item in raw_journals
    ):
        raise ValueError(
            f"Profile '{resolved_path}' field 'journals' must be a JSON array of strings."
        )
    journal_names = _dedupe_preserving_order(raw_journals)
    name = _optional_string(payload, "name", resolved_path) or resolved_path.stem
    workbook_value = _optional_string(payload, "workbook", resolved_path)
    config_value = _optional_string(payload, "config", resolved_path)
    csv_output_value = _optional_string(payload, "csv_output", resolved_path)
    articles_sheet = _optional_string(payload, "articles_sheet", resolved_path)
    directory_sheet = _optional_string(payload, "directory_sheet", resolved_path)

    return TrackingProfile(
        name=name,
        profile_path=resolved_path,
        workbook_path=_resolve_optional_path(base_dir, workbook_value),
        config_path=_resolve_optional_path(base_dir, config_value) or default_config_path(),
        csv_output_path=_resolve_optional_path(base_dir, csv_output_value),
        years=_optional_years(payload, resolved_path),
        articles_sheet=articles_sheet,
        directory_sheet=directory_sheet,
        journal_names=journal_names,
    )
