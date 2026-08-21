"""Excel workbook loading and result export."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from .search import normalize_columns

SUPPORTED_SUFFIXES = {".xlsx", ".xls", ".xlsm"}


def load_workbook(
    path: str | Path,
    sheet_name: str | None = None,
) -> tuple[list[str], str, pd.DataFrame]:
    """Return sheet names, the selected sheet, and its normalized dataframe."""

    workbook_path = Path(path)
    if workbook_path.suffix.lower() not in SUPPORTED_SUFFIXES:
        raise ValueError("Unterstützt werden .xlsx, .xls und .xlsm.")

    with pd.ExcelFile(workbook_path, engine=None) as workbook:
        sheets = list(workbook.sheet_names)
        if not sheets:
            raise ValueError("Die Arbeitsmappe enthält keine Tabellenblätter.")
        selected = sheet_name or sheets[0]
        if selected not in sheets:
            raise ValueError(f"Unbekanntes Tabellenblatt: {selected}")
        dataframe = pd.read_excel(workbook, sheet_name=selected, dtype=str)

    return sheets, selected, normalize_columns(dataframe)


def export_dataframe(
    dataframe: pd.DataFrame,
    path: str | Path,
    *,
    csv_separator: str = ";",
) -> Path:
    """Export all results as XLSX or Excel-friendly UTF-8 CSV."""

    destination = Path(path)
    suffix = destination.suffix.lower()
    if suffix == ".xlsx":
        dataframe.to_excel(destination, index=False, engine="openpyxl")
    elif suffix == ".csv":
        dataframe.to_csv(destination, index=False, sep=csv_separator, encoding="utf-8-sig")
    else:
        raise ValueError("Der Export muss auf .xlsx oder .csv enden.")
    return destination
