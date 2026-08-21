from datetime import date
from pathlib import Path

import pandas as pd

from excel_search.presentation import create_print_preview, format_cell_value


def test_date_is_formatted_for_display() -> None:
    assert format_cell_value(date(2026, 8, 21)) == "21.08.2026"


def test_print_preview_is_landscape_table_with_spacers(tmp_path: Path) -> None:
    destination = tmp_path / "preview.html"
    dataframe = pd.DataFrame(
        {
            "Datum": [pd.Timestamp("2026-08-21")],
            "Beschreibung": ["Langer <Text> & Details"],
        }
    )

    create_print_preview(dataframe, destination)

    document = destination.read_text(encoding="utf-8")
    assert "@page { size: landscape" in document
    assert 'class="spacer"' in document
    assert "window.print()" in document
    assert "21.08.2026" in document
    assert "Langer &lt;Text&gt; &amp; Details" in document
    assert 'td class="description"' in document
