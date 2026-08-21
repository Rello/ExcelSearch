# ExcelSearcher

ExcelSearcher ist eine lokale Desktop-Anwendung zum Durchsuchen von Excel-Arbeitsmappen.
Die Dateien werden ausschließlich auf dem eigenen Rechner verarbeitet.

## Funktionen

- `.xlsx`, `.xls` und `.xlsm` öffnen
- Tabellenblatt auswählen
- beliebig viele Suchkriterien über eine Spaltenauswahl zusammenstellen
- Kriterien mit UND oder ODER verknüpfen
- literale Teiltreffer oder exakte Übereinstimmungen suchen
- Groß-/Kleinschreibung optional beachten
- Datumswerte einheitlich als `TT.MM.JJJJ` anzeigen und durchsuchen
- Spaltenbreiten am Inhalt ausrichten und den freien Platz für „Beschreibung“ verwenden
- vollständige Beschreibungen beim Überfahren mit der Maus anzeigen
- Trefferlisten ab 501 Einträgen in Seiten mit jeweils 500 Zeilen anzeigen
- sämtliche Treffer nach `.xlsx` oder als Excel-kompatibles UTF-8-CSV exportieren
- Browser-Druckvorschau als Tabelle im Querformat mit einer Leerzeile je Datensatz öffnen

Kommas und Zeichen wie `[`, `*` oder `.` sind normale Bestandteile eines Suchbegriffs und
werden nicht als reguläre Ausdrücke interpretiert.

## Lokale Entwicklung

Vorausgesetzt wird Python 3.11 oder neuer. Für reproduzierbare Ergebnisse verwendet CI
Python 3.13.15 und die exakt festgelegten Versionen aus `requirements-dev.txt`.

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements-dev.txt
python search_excel.py
```

Unter Windows wird die virtuelle Umgebung mit `.venv\Scripts\activate` aktiviert.

## Prüfungen

```bash
ruff check .
python -m pytest
python search_excel.py --smoke-test
```

Die Tests decken literale Suche, exakte Suche, Groß-/Kleinschreibung, UND/ODER,
Datumsformatierung, Spaltenvalidierung, Tabellenblätter, Druckvorschau sowie XLSX- und
CSV-Export ab. CI initialisiert die Oberfläche außerdem in einer virtuellen Anzeige.

## Standalone-Builds

GitHub Actions baut bei Änderungen an `main` automatisch:

- `ExcelSearcher.exe` für Windows
- `ExcelSearcher.app` als ZIP für macOS

Jedes gebaute Programm wird anschließend mit `--smoke-test` gestartet. Die Builds sind
absichtlich weder signiert noch notarisiert und für die lokale Verwendung vorgesehen.

Ein Tag wie `v2.1.0` erzeugt nach erfolgreichen Builds automatisch ein GitHub-Release:

```bash
git tag v2.1.0
git push origin v2.1.0
```

## Projektstruktur

- `search_excel.py`: Tkinter-Oberfläche und Hintergrundaufgaben
- `excel_search/search.py`: testbare Suchlogik
- `excel_search/workbook.py`: Excel-Import und Export
- `excel_search/presentation.py`: Datumsformatierung und Druckvorschau
- `tests/`: automatisierte Tests

## Grenzen

- Formeln werden als die von der Excel-Bibliothek gelesenen Zellwerte verarbeitet.
- Die Druckvorschau öffnet den Standardbrowser; dessen Druckdialog übernimmt die Ausgabe.
- Eine Projektlizenz wurde bewusst nicht festgelegt; sie muss vom Repository-Eigentümer
  separat gewählt werden.
