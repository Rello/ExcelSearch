"""Formatting shared by search results, exports, and print previews."""

from __future__ import annotations

import html
from datetime import date, datetime
from pathlib import Path

import pandas as pd


def format_cell_value(value: object) -> str:
    """Format one value as it should appear to users."""

    if value is None or pd.isna(value):
        return ""
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.strftime("%d.%m.%Y")
    return str(value)


def formatted_dataframe(dataframe: pd.DataFrame) -> pd.DataFrame:
    """Return a string dataframe using the common presentation format."""

    return dataframe.map(format_cell_value)


def create_print_preview(dataframe: pd.DataFrame, destination: str | Path) -> Path:
    """Create a landscape HTML preview with a printable table."""

    path = Path(destination)
    headers = [str(column) for column in dataframe.columns]
    description_indexes = {
        index for index, heading in enumerate(headers) if "beschreibung" in heading.casefold()
    }

    header_html = "".join(
        f'<th class="{"description" if index in description_indexes else ""}">'
        f"{html.escape(heading)}</th>"
        for index, heading in enumerate(headers)
    )
    rows: list[str] = []
    for values in dataframe.itertuples(index=False, name=None):
        cells = "".join(
            f'<td class="{"description" if index in description_indexes else ""}">'
            f"{html.escape(format_cell_value(value))}</td>"
            for index, value in enumerate(values)
        )
        rows.append(f'<tr class="data-row">{cells}</tr>')
        rows.append(f'<tr class="spacer"><td colspan="{max(1, len(headers))}"></td></tr>')

    document = f"""<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<title>ExcelSearcher – Druckvorschau</title>
<style>
@page {{ size: landscape; margin: 12mm; }}
body {{ color: #111; font-family: Arial, sans-serif; font-size: 10pt; margin: 16px; }}
.toolbar {{ margin-bottom: 14px; }}
button {{ cursor: pointer; font-size: 11pt; padding: 7px 16px; }}
table {{ border-collapse: collapse; table-layout: auto; width: 100%; }}
thead {{ display: table-header-group; }}
th {{ background: #e9ecef; font-weight: 700; text-align: left; }}
th, td {{ border: 1px solid #777; padding: 5px 7px; vertical-align: top; }}
th.description, td.description {{ width: 48%; white-space: pre-wrap; overflow-wrap: anywhere; }}
tr.data-row {{ break-inside: avoid; page-break-inside: avoid; }}
tr.spacer {{ height: 9px; }}
tr.spacer td {{ border: 0; padding: 0; }}
@media print {{
  body {{ margin: 0; }}
  .toolbar {{ display: none; }}
}}
</style>
</head>
<body>
<div class="toolbar">
  <button type="button" onclick="window.print()">Drucken…</button>
</div>
<table>
  <thead><tr>{header_html}</tr></thead>
  <tbody>{"".join(rows)}</tbody>
</table>
</body>
</html>
"""
    path.write_text(document, encoding="utf-8")
    return path
