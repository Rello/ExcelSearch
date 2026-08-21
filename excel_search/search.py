"""Pure, UI-independent search operations."""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass
from functools import reduce
from operator import and_, or_
from typing import Literal

import pandas as pd

from .presentation import formatted_dataframe

MatchMode = Literal["all", "any"]


@dataclass(frozen=True, slots=True)
class SearchCriterion:
    """One literal search criterion."""

    value: str
    column: str | None = None
    exact: bool = False

    def __post_init__(self) -> None:
        if not self.value:
            raise ValueError("Der Suchbegriff darf nicht leer sein.")


def normalize_columns(dataframe: pd.DataFrame) -> pd.DataFrame:
    """Return a copy with non-empty, unique string column names."""

    result = dataframe.copy()
    seen: dict[str, int] = {}
    normalized: list[str] = []
    for index, raw_name in enumerate(result.columns, start=1):
        base = str(raw_name).strip() or f"Spalte {index}"
        seen[base] = seen.get(base, 0) + 1
        normalized.append(base if seen[base] == 1 else f"{base} ({seen[base]})")
    result.columns = normalized
    return result


def search_dataframe(
    dataframe: pd.DataFrame,
    criteria: Sequence[SearchCriterion],
    *,
    match_mode: MatchMode = "all",
    case_sensitive: bool = False,
) -> pd.DataFrame:
    """Search a dataframe using literal, consistently case-aware matching."""

    if not criteria:
        raise ValueError("Mindestens ein Suchkriterium ist erforderlich.")
    if match_mode not in {"all", "any"}:
        raise ValueError(f"Unbekannte Verknüpfung: {match_mode}")

    searchable = formatted_dataframe(dataframe)
    masks = [
        _criterion_mask(searchable, criterion, case_sensitive=case_sensitive)
        for criterion in criteria
    ]
    operator = and_ if match_mode == "all" else or_
    combined = reduce(operator, masks)
    return dataframe.loc[combined].copy()


def _criterion_mask(
    dataframe: pd.DataFrame,
    criterion: SearchCriterion,
    *,
    case_sensitive: bool,
) -> pd.Series:
    if criterion.column is not None:
        if criterion.column not in dataframe.columns:
            raise ValueError(f"Unbekannte Spalte: {criterion.column}")
        return _series_mask(
            dataframe[criterion.column],
            criterion.value,
            exact=criterion.exact,
            case_sensitive=case_sensitive,
        )

    column_masks = [
        _series_mask(
            dataframe[column],
            criterion.value,
            exact=criterion.exact,
            case_sensitive=case_sensitive,
        )
        for column in dataframe.columns
    ]
    if not column_masks:
        return pd.Series(False, index=dataframe.index, dtype=bool)
    return reduce(or_, column_masks)


def _series_mask(
    series: pd.Series,
    value: str,
    *,
    exact: bool,
    case_sensitive: bool,
) -> pd.Series:
    normalized = series.fillna("").astype(str)
    needle = value
    if not case_sensitive:
        normalized = normalized.str.casefold()
        needle = needle.casefold()

    if exact:
        return normalized.eq(needle)
    return normalized.str.contains(needle, regex=False, na=False)
