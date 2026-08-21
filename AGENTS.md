# AGENTS.md

Diese Datei enthält die wichtigsten Hinweise für KI-Agenten, die an diesem Repository arbeiten.

## Projektüberblick

ExcelSearcher ist eine lokale Python-/Tkinter-Desktopanwendung zum Durchsuchen von
`.xlsx`-, `.xls`- und `.xlsm`-Arbeitsmappen. Excel-Dateien werden ausschließlich lokal
verarbeitet. Die Benutzeroberfläche und alle Meldungen sind deutschsprachig.

## Struktur

- `search_excel.py`: Tkinter-Oberfläche, Hintergrundaufgaben, Seitennavigation, Druck und
  Programmeinstieg.
- `excel_search/search.py`: UI-unabhängige Suchlogik und Spaltennormalisierung.
- `excel_search/workbook.py`: Import von Arbeitsmappen sowie XLSX-/CSV-Export.
- `excel_search/presentation.py`: gemeinsame Darstellung von Zellwerten für Suche, Export
  und Druckvorschau.
- `tests/`: Pytest-Tests für die fachliche Logik.
- `.github/workflows/`: CI, PyInstaller-Builds und Releases.

## Entwicklung

Voraussetzung ist Python 3.11 oder neuer. CI verwendet Python 3.13.15 und die fest
definierten Versionen aus `requirements-dev.txt`.

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements-dev.txt
python search_excel.py
```

Vor Abschluss einer Änderung sind mindestens diese Prüfungen auszuführen:

```bash
ruff check .
python -m pytest
python search_excel.py --smoke-test
```

Bei reinen Logikänderungen zuerst die betroffene Testdatei gezielt ausführen. Änderungen
an der GUI sollen zusätzlich die Initialisierung unter einer verfügbaren Anzeige prüfen;
die CI verwendet dafür `xvfb-run`.

## Wichtige Regeln und Invarianten

- Fachlogik nach Möglichkeit in `excel_search/` halten und unabhängig von Tkinter testen.
- Lang laufende Datei-, Such-, Export- und Druckoperationen dürfen die Tkinter-Ereignisschleife
  nicht blockieren. Hintergrundarbeit über den vorhandenen Executor ausführen und UI-Updates
  ausschließlich im Hauptthread vornehmen.
- Suchbegriffe sind literal. Zeichen wie `[`, `*` und `.` dürfen niemals unbeabsichtigt als
  reguläre Ausdrücke interpretiert werden.
- Exakte/partielle Suche, Groß-/Kleinschreibung sowie UND-/ODER-Verknüpfung müssen
  konsistent bleiben.
- Datentypen im geladenen `DataFrame` erhalten. Benutzerorientierte Formatierung zentral
  über `excel_search/presentation.py` anwenden; Datumswerte werden als `DD.MM.YYYY` angezeigt.
- CSV-Exporte bleiben Excel-kompatibel: UTF-8 mit BOM und standardmäßig Semikolon als
  Trennzeichen. XLSX-Exporte behalten native Werte und passende Datumsformate.
- Alle Treffer werden exportiert oder gedruckt; die Seitengröße von 500 Zeilen betrifft nur
  die Anzeige.
- Unterstützte Dateiendungen und Fehlermeldungen nicht stillschweigend ändern. Neues
  Verhalten mit fokussierten Pytest-Tests absichern.
- Abhängigkeiten sind absichtlich exakt gepinnt. Aktualisierungen in `pyproject.toml`,
  `requirements.txt` und gegebenenfalls `requirements-dev.txt` synchron halten.
- Bei einer Versionsänderung mindestens `pyproject.toml` und `excel_search/__init__.py`
  gemeinsam aktualisieren.
- `logo.jpg` wird in die PyInstaller-Artefakte eingebettet; Ressourcen müssen sowohl aus dem
  Quellbaum als auch aus einem gebündelten Programm erreichbar bleiben.
- Vor Änderungen `git status --short` prüfen und vorhandene, nicht zugehörige Änderungen
  nicht überschreiben oder zurücksetzen.

## Commits

Commits sollen klein und thematisch geschlossen sein. Jeder Commit muss diese beiden Trailer
enthalten:

```text
Signed-off-by: Rello <github@scherello.de>
Assisted-by: Codex:GPT-5
```

