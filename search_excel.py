#!/usr/bin/env python3
"""ExcelSearcher desktop application."""

from __future__ import annotations

import platform
import sys
import tempfile
import tkinter as tk
import tkinter.font as tkfont
import webbrowser
from collections.abc import Callable
from concurrent.futures import Future, ThreadPoolExecutor
from contextlib import suppress
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any

import pandas as pd
from PIL import Image, ImageTk

from excel_search import __version__
from excel_search.presentation import create_print_preview, format_cell_value
from excel_search.search import SearchCriterion, search_dataframe
from excel_search.workbook import export_dataframe, load_workbook

ALL_COLUMNS = "Alle Spalten"
PAGE_SIZE = 500


class ExcelSearcher(tk.Tk):
    """Responsive Tkinter UI for literal searches in Excel workbooks."""

    def __init__(self) -> None:
        super().__init__()
        self.title(f"ExcelSearcher {__version__}")
        self.geometry("1120x760")
        self.minsize(820, 560)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self.executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="excel-search")
        self.active_future: Future[Any] | None = None
        self.file_path: Path | None = None
        self.dataframe: pd.DataFrame | None = None
        self.result: pd.DataFrame | None = None
        self.criteria: list[SearchCriterion] = []
        self.page = 0
        self.logo: ImageTk.PhotoImage | None = None
        self.display_columns: list[str] = []
        self.resize_after_id: str | None = None
        self.tooltip_after_id: str | None = None
        self.tooltip_target: tuple[str, int, str] | None = None
        self.tooltip_window: tk.Toplevel | None = None
        self.preview_paths: set[Path] = set()

        self._build_ui()
        self._set_workbook_controls(False)
        self._update_result_controls()

    def _build_ui(self) -> None:
        style = ttk.Style(self)
        available_themes = style.theme_names()
        preferred = "aqua" if platform.system() == "Darwin" else "vista"
        if preferred in available_themes:
            style.theme_use(preferred)
        style.configure("Treeview", rowheight=24)

        root = ttk.Frame(self, padding=10)
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(4, weight=1)

        self._build_header(root)
        self._build_file_row(root)
        self._build_search_panel(root)
        self._build_status_row(root)
        self._build_result_table(root)
        self._build_pagination(root)

    def _build_header(self, parent: ttk.Frame) -> None:
        frame = ttk.Frame(parent)
        frame.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        logo_path = self.resource_path("logo.jpg")
        if logo_path.exists():
            try:
                with Image.open(logo_path) as image:
                    image.thumbnail((750, 90), Image.Resampling.LANCZOS)
                    self.logo = ImageTk.PhotoImage(image.copy())
                ttk.Label(frame, image=self.logo).pack()
            except (OSError, ValueError):
                pass

    def _build_file_row(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="Arbeitsmappe", padding=8)
        frame.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        frame.columnconfigure(1, weight=1)

        self.open_button = ttk.Button(frame, text="Excel auswählen…", command=self.choose_file)
        self.open_button.grid(row=0, column=0, padx=(0, 8))

        self.file_label = ttk.Label(frame, text="Keine Datei ausgewählt", anchor="w")
        self.file_label.grid(row=0, column=1, sticky="ew", padx=(0, 12))

        ttk.Label(frame, text="Tabellenblatt:").grid(row=0, column=2, padx=(0, 5))
        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(
            frame, textvariable=self.sheet_var, state="readonly", width=24
        )
        self.sheet_combo.grid(row=0, column=3)
        self.sheet_combo.bind("<<ComboboxSelected>>", self._sheet_changed)

    def _build_search_panel(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="Suchkriterien", padding=8)
        frame.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Spalte:").grid(row=0, column=0, sticky="w")
        self.column_var = tk.StringVar(value=ALL_COLUMNS)
        self.column_combo = ttk.Combobox(
            frame, textvariable=self.column_var, state="readonly", width=25
        )
        self.column_combo.grid(row=1, column=0, sticky="ew", padx=(0, 8))

        ttk.Label(frame, text="Suchbegriff:").grid(row=0, column=1, sticky="w")
        self.term_entry = ttk.Entry(frame)
        self.term_entry.grid(row=1, column=1, sticky="ew", padx=(0, 8))
        self.term_entry.bind("<Return>", lambda _event: self.add_criterion())

        self.exact_var = tk.BooleanVar(value=False)
        self.exact_check = ttk.Checkbutton(
            frame, text="Exakte Übereinstimmung", variable=self.exact_var
        )
        self.exact_check.grid(row=1, column=2, padx=(0, 8))

        self.add_button = ttk.Button(frame, text="Kriterium hinzufügen", command=self.add_criterion)
        self.add_button.grid(row=1, column=3)

        self.criteria_tree = ttk.Treeview(
            frame,
            columns=("column", "term", "mode"),
            show="headings",
            height=3,
            selectmode="extended",
        )
        self.criteria_tree.heading("column", text="Spalte")
        self.criteria_tree.heading("term", text="Suchbegriff")
        self.criteria_tree.heading("mode", text="Vergleich")
        self.criteria_tree.column("column", width=180, stretch=False)
        self.criteria_tree.column("term", width=400)
        self.criteria_tree.column("mode", width=110, stretch=False)
        self.criteria_tree.grid(row=2, column=0, columnspan=4, sticky="ew", pady=(8, 6))

        options = ttk.Frame(frame)
        options.grid(row=3, column=0, columnspan=4, sticky="ew")
        options.columnconfigure(6, weight=1)

        self.remove_button = ttk.Button(
            options, text="Auswahl entfernen", command=self.remove_criteria
        )
        self.remove_button.grid(row=0, column=0, padx=(0, 5))
        self.clear_button = ttk.Button(options, text="Alle entfernen", command=self.clear_criteria)
        self.clear_button.grid(row=0, column=1, padx=(0, 14))

        ttk.Label(options, text="Verknüpfung:").grid(row=0, column=2, padx=(0, 5))
        self.match_mode_var = tk.StringVar(value="Alle (UND)")
        self.match_mode_combo = ttk.Combobox(
            options,
            textvariable=self.match_mode_var,
            values=("Alle (UND)", "Mindestens eines (ODER)"),
            state="readonly",
            width=22,
        )
        self.match_mode_combo.grid(row=0, column=3, padx=(0, 8))

        self.case_sensitive_var = tk.BooleanVar(value=False)
        self.case_sensitive_check = ttk.Checkbutton(
            options,
            text="Groß-/Kleinschreibung beachten",
            variable=self.case_sensitive_var,
        )
        self.case_sensitive_check.grid(row=0, column=4, padx=(0, 8))

        self.search_button = ttk.Button(options, text="Suchen", command=self.start_search)
        self.search_button.grid(row=0, column=7)

    def _build_status_row(self, parent: ttk.Frame) -> None:
        frame = ttk.Frame(parent)
        frame.grid(row=3, column=0, sticky="ew", pady=(0, 5))
        frame.columnconfigure(0, weight=1)

        self.status_var = tk.StringVar(value="Bereit")
        ttk.Label(frame, textvariable=self.status_var).grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(frame, mode="indeterminate", length=130)
        self.progress.grid(row=0, column=1, padx=(8, 0))

        ttk.Label(frame, text="CSV-Trennzeichen:").grid(row=0, column=2, padx=(14, 5))
        self.csv_separator_var = tk.StringVar(value="Semikolon (;)")
        self.csv_separator_combo = ttk.Combobox(
            frame,
            textvariable=self.csv_separator_var,
            values=("Semikolon (;)", "Komma (,)", "Tabulator"),
            state="readonly",
            width=16,
        )
        self.csv_separator_combo.grid(row=0, column=3)

        self.export_button = ttk.Button(frame, text="Exportieren…", command=self.export_results)
        self.export_button.grid(row=0, column=4, padx=(8, 5))
        self.print_button = ttk.Button(frame, text="Druckvorschau", command=self.print_results)
        self.print_button.grid(row=0, column=5)

    def _build_result_table(self, parent: ttk.Frame) -> None:
        frame = ttk.Frame(parent)
        frame.grid(row=4, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(frame, columns=(), show="headings")
        vertical = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        horizontal = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<Configure>", lambda _event: self._schedule_column_widths())
        self.tree.bind("<Motion>", self._tree_motion)
        self.tree.bind("<Leave>", lambda _event: self._hide_tooltip())
        self.tree.bind("<ButtonPress>", lambda _event: self._hide_tooltip())
        self.tree.bind("<MouseWheel>", lambda _event: self._hide_tooltip())

    def _build_pagination(self, parent: ttk.Frame) -> None:
        self.pagination_frame = ttk.Frame(parent)
        self.pagination_frame.grid(row=5, column=0, sticky="ew", pady=(6, 0))
        self.pagination_frame.columnconfigure(1, weight=1)

        self.previous_button = ttk.Button(
            self.pagination_frame, text="← Zurück", command=self.previous_page
        )
        self.previous_button.grid(row=0, column=0)
        self.page_var = tk.StringVar(value="Keine Ergebnisse")
        ttk.Label(self.pagination_frame, textvariable=self.page_var, anchor="center").grid(
            row=0, column=1
        )
        self.next_button = ttk.Button(
            self.pagination_frame, text="Weiter →", command=self.next_page
        )
        self.next_button.grid(row=0, column=2)

    def resource_path(self, relative: str) -> Path:
        if getattr(sys, "frozen", False):
            base = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
        else:
            base = Path(__file__).resolve().parent
        direct = base / relative
        fallback = base / "Resources" / relative
        return fallback if not direct.exists() and fallback.exists() else direct

    def choose_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Excel-Datei öffnen",
            filetypes=[
                ("Excel-Dateien", ("*.xlsx", "*.xls", "*.xlsm")),
                ("Alle Dateien", "*.*"),
            ],
        )
        if not path:
            return
        self.file_path = Path(path)
        self._reset_loaded_data()
        self.file_label.configure(text=self.file_path.name)
        self._submit_task(
            "Arbeitsmappe wird geladen…",
            load_workbook,
            self.file_path,
            on_success=self._workbook_loaded,
        )

    def _sheet_changed(self, _event: tk.Event[Any] | None = None) -> None:
        if self.file_path is None or not self.sheet_var.get():
            return
        self._reset_search_state()
        self._submit_task(
            f"Tabellenblatt „{self.sheet_var.get()}“ wird geladen…",
            load_workbook,
            self.file_path,
            self.sheet_var.get(),
            on_success=self._workbook_loaded,
        )

    def _workbook_loaded(self, payload: tuple[list[str], str, pd.DataFrame]) -> None:
        sheets, selected, dataframe = payload
        self.dataframe = dataframe
        self.result = None
        self.sheet_combo.configure(values=sheets)
        self.sheet_var.set(selected)
        columns = [ALL_COLUMNS, *map(str, dataframe.columns)]
        self.column_combo.configure(values=columns)
        self.column_var.set(ALL_COLUMNS)
        self._configure_result_columns(dataframe)
        self.status_var.set(
            f"{self.file_path.name if self.file_path else 'Datei'} · {selected} · "
            f"{len(dataframe):,} Zeilen · {len(dataframe.columns)} Spalten"
        )
        self._set_workbook_controls(True)
        self._update_result_controls()
        self.term_entry.focus_set()

    def add_criterion(self) -> bool:
        value = self.term_entry.get().strip()
        if not value:
            messagebox.showwarning("Leerer Suchbegriff", "Bitte einen Suchbegriff eingeben.")
            return False
        selected_index = self.column_combo.current()
        selected_column = None
        if selected_index > 0 and self.dataframe is not None:
            selected_column = str(self.dataframe.columns[selected_index - 1])
        criterion = SearchCriterion(
            value=value,
            column=selected_column,
            exact=self.exact_var.get(),
        )
        self.criteria.append(criterion)
        self.criteria_tree.insert(
            "",
            "end",
            values=(
                criterion.column or ALL_COLUMNS,
                criterion.value,
                "Exakt" if criterion.exact else "Enthält",
            ),
        )
        self.term_entry.delete(0, "end")
        self.term_entry.focus_set()
        return True

    def remove_criteria(self) -> None:
        selected = set(self.criteria_tree.selection())
        if not selected:
            return
        retained: list[SearchCriterion] = []
        for index, item in enumerate(self.criteria_tree.get_children()):
            if item in selected:
                self.criteria_tree.delete(item)
            else:
                retained.append(self.criteria[index])
        self.criteria = retained

    def clear_criteria(self) -> None:
        self.criteria.clear()
        self.criteria_tree.delete(*self.criteria_tree.get_children())

    def start_search(self) -> None:
        if self.dataframe is None:
            messagebox.showwarning("Keine Datei", "Bitte zuerst eine Excel-Datei auswählen.")
            return
        if self.term_entry.get().strip() and not self.add_criterion():
            return
        if not self.criteria:
            messagebox.showwarning(
                "Keine Suchkriterien", "Bitte mindestens ein Kriterium hinzufügen."
            )
            return

        match_mode = "all" if self.match_mode_var.get() == "Alle (UND)" else "any"
        self._submit_task(
            "Tabelle wird durchsucht…",
            search_dataframe,
            self.dataframe,
            tuple(self.criteria),
            match_mode=match_mode,
            case_sensitive=self.case_sensitive_var.get(),
            on_success=self._search_completed,
        )

    def _search_completed(self, result: pd.DataFrame) -> None:
        self.result = result
        self.page = 0
        self._render_page()
        self.status_var.set(f"{len(result):,} Treffer gefunden")

    def _configure_result_columns(self, dataframe: pd.DataFrame) -> None:
        self.display_columns = list(map(str, dataframe.columns))
        column_ids = [f"column_{index}" for index in range(len(self.display_columns))]
        self.tree.configure(columns=column_ids)
        for column_id, heading in zip(column_ids, self.display_columns, strict=True):
            self.tree.heading(column_id, text=heading)
            self.tree.column(column_id, width=110, minwidth=60, anchor="w", stretch=False)
        self.tree.delete(*self.tree.get_children())
        self._schedule_column_widths()

    @staticmethod
    def _is_description_heading(heading: str) -> bool:
        return "beschreibung" in heading.casefold()

    def _schedule_column_widths(self) -> None:
        if self.resize_after_id is not None:
            with suppress(tk.TclError):
                self.after_cancel(self.resize_after_id)
        self.resize_after_id = self.after(100, self._fit_column_widths)

    def _fit_column_widths(self) -> None:
        self.resize_after_id = None
        if not self.display_columns:
            return

        source = (
            self.result if self.result is not None and not self.result.empty else self.dataframe
        )
        font = tkfont.nametofont("TkDefaultFont")
        widths: list[int] = []
        description_indexes: list[int] = []
        for index, heading in enumerate(self.display_columns):
            if self._is_description_heading(heading):
                description_indexes.append(index)
                widths.append(0)
                continue
            measured = font.measure(heading) + 28
            if source is not None and index < len(source.columns):
                for value in source.iloc[:250, index]:
                    measured = max(measured, font.measure(format_cell_value(value)) + 28)
            widths.append(min(max(measured, 70), 190))

        available = max(self.tree.winfo_width() - 24, 600)
        fixed_width = sum(widths)
        if description_indexes:
            description_width = max(
                320,
                (available - fixed_width) // len(description_indexes),
            )
            for index in description_indexes:
                widths[index] = description_width

        for index, width in enumerate(widths):
            self.tree.column(
                f"column_{index}",
                width=width,
                minwidth=60,
                stretch=index in description_indexes,
            )

    def _tree_motion(self, event: tk.Event[Any]) -> None:
        row = self.tree.identify_row(event.y)
        column_token = self.tree.identify_column(event.x)
        if not row or not column_token.startswith("#"):
            self._hide_tooltip()
            return
        try:
            column_index = int(column_token[1:]) - 1
        except ValueError:
            self._hide_tooltip()
            return
        if not 0 <= column_index < len(self.display_columns):
            self._hide_tooltip()
            return
        if not self._is_description_heading(self.display_columns[column_index]):
            self._hide_tooltip()
            return

        value = self.tree.set(row, f"column_{column_index}")
        target = (row, column_index, value)
        if not value or target == self.tooltip_target:
            return
        self._hide_tooltip()
        self.tooltip_target = target
        self.tooltip_after_id = self.after(
            350,
            self._show_tooltip,
            target,
            event.x_root + 14,
            event.y_root + 18,
        )

    def _show_tooltip(
        self,
        target: tuple[str, int, str],
        x_position: int,
        y_position: int,
    ) -> None:
        self.tooltip_after_id = None
        if target != self.tooltip_target:
            return
        if self.tooltip_window is not None:
            self.tooltip_window.destroy()
        tooltip = tk.Toplevel(self)
        tooltip.wm_overrideredirect(True)
        tooltip.wm_attributes("-topmost", True)
        label = ttk.Label(
            tooltip,
            text=target[2],
            justify="left",
            padding=9,
            relief="solid",
            borderwidth=1,
            wraplength=650,
        )
        label.pack()
        tooltip.update_idletasks()
        x_position = min(
            x_position,
            self.winfo_screenwidth() - tooltip.winfo_reqwidth() - 12,
        )
        if y_position + tooltip.winfo_reqheight() > self.winfo_screenheight() - 12:
            y_position -= tooltip.winfo_reqheight() + 34
        tooltip.geometry(f"+{max(0, x_position)}+{max(0, y_position)}")
        self.tooltip_window = tooltip

    def _hide_tooltip(self) -> None:
        if self.tooltip_after_id is not None:
            with suppress(tk.TclError):
                self.after_cancel(self.tooltip_after_id)
            self.tooltip_after_id = None
        if self.tooltip_window is not None:
            self.tooltip_window.destroy()
            self.tooltip_window = None
        self.tooltip_target = None

    def _render_page(self) -> None:
        self.tree.delete(*self.tree.get_children())
        if self.result is None:
            self._update_result_controls()
            return
        start = self.page * PAGE_SIZE
        stop = min(start + PAGE_SIZE, len(self.result))
        for row in self.result.iloc[start:stop].itertuples(index=False, name=None):
            self.tree.insert("", "end", values=[format_cell_value(value) for value in row])
        self._update_result_controls()
        self._schedule_column_widths()

    def previous_page(self) -> None:
        if self.page > 0:
            self.page -= 1
            self._render_page()

    def next_page(self) -> None:
        if self.result is not None and (self.page + 1) * PAGE_SIZE < len(self.result):
            self.page += 1
            self._render_page()

    def _update_result_controls(self) -> None:
        count = 0 if self.result is None else len(self.result)
        pages = max(1, (count + PAGE_SIZE - 1) // PAGE_SIZE)
        if self.result is None:
            self.page_var.set("Keine Ergebnisse")
        elif count == 0:
            self.page_var.set("0 Treffer")
        elif count <= PAGE_SIZE:
            self.page_var.set(f"{count:,} Treffer")
        else:
            start = self.page * PAGE_SIZE + 1
            stop = min((self.page + 1) * PAGE_SIZE, count)
            self.page_var.set(f"{start:,}–{stop:,} von {count:,} · Seite {self.page + 1}/{pages}")
        if count > PAGE_SIZE:
            self.previous_button.grid()
            self.next_button.grid()
            self.previous_button.configure(state="normal" if self.page > 0 else "disabled")
            self.next_button.configure(
                state="normal" if (self.page + 1) * PAGE_SIZE < count else "disabled"
            )
        else:
            self.previous_button.grid_remove()
            self.next_button.grid_remove()
        result_state = "normal" if self.result is not None else "disabled"
        self.export_button.configure(state=result_state)
        self.print_button.configure(state=result_state)

    def export_results(self) -> None:
        if self.result is None:
            messagebox.showwarning("Keine Ergebnisse", "Bitte zuerst eine Suche durchführen.")
            return
        path = filedialog.asksaveasfilename(
            title="Suchergebnisse exportieren",
            defaultextension=".xlsx",
            filetypes=(("Excel-Arbeitsmappe", "*.xlsx"), ("CSV-Datei", "*.csv")),
        )
        if not path:
            return
        separator = {
            "Semikolon (;)": ";",
            "Komma (,)": ",",
            "Tabulator": "\t",
        }.get(self.csv_separator_var.get(), ";")
        self._submit_task(
            "Ergebnisse werden exportiert…",
            export_dataframe,
            self.result.copy(),
            path,
            csv_separator=separator,
            on_success=self._export_completed,
        )

    def _export_completed(self, destination: Path) -> None:
        self.status_var.set("Export abgeschlossen")
        messagebox.showinfo("Export abgeschlossen", f"Gespeichert unter:\n{destination}")

    def print_results(self) -> None:
        if self.result is None:
            messagebox.showwarning("Keine Ergebnisse", "Bitte zuerst eine Suche durchführen.")
            return
        with tempfile.NamedTemporaryFile(
            prefix="excelsearch-preview-", suffix=".html", delete=False
        ) as temporary:
            preview_path = Path(temporary.name)
        self.preview_paths.add(preview_path)
        self._submit_task(
            "Druckvorschau wird erstellt…",
            create_print_preview,
            self.result.copy(),
            preview_path,
            on_success=self._print_preview_completed,
        )

    def _print_preview_completed(self, preview_path: Path) -> None:
        try:
            opened = webbrowser.open_new_tab(preview_path.as_uri())
        except webbrowser.Error as error:
            messagebox.showerror("Druckvorschau", str(error))
            return
        if not opened:
            messagebox.showerror(
                "Druckvorschau",
                "Die Druckvorschau konnte nicht im Browser geöffnet werden.",
            )
            return
        self.status_var.set("Druckvorschau geöffnet")

    def _submit_task(
        self,
        status: str,
        function: Callable[..., Any],
        *args: Any,
        on_success: Callable[[Any], None],
        **kwargs: Any,
    ) -> None:
        if self.active_future is not None:
            return
        self._set_busy(True, status)
        self.active_future = self.executor.submit(function, *args, **kwargs)
        self.after(50, self._poll_task, on_success)

    def _poll_task(self, on_success: Callable[[Any], None]) -> None:
        future = self.active_future
        if future is None:
            return
        if not future.done():
            self.after(50, self._poll_task, on_success)
            return
        self.active_future = None
        self._set_busy(False)
        try:
            payload = future.result()
        except Exception as error:
            self.status_var.set("Vorgang fehlgeschlagen")
            messagebox.showerror("Fehler", str(error) or error.__class__.__name__)
            return
        on_success(payload)

    def _set_busy(self, busy: bool, status: str = "Bereit") -> None:
        state = "disabled" if busy else "normal"
        for widget in (self.open_button, self.add_button, self.search_button):
            widget.configure(state=state)
        if busy:
            self.sheet_combo.configure(state="disabled")
            self.export_button.configure(state="disabled")
            self.print_button.configure(state="disabled")
            self.previous_button.configure(state="disabled")
            self.next_button.configure(state="disabled")
            self.progress.start(10)
            self.status_var.set(status)
        else:
            self.progress.stop()
            self._set_workbook_controls(self.dataframe is not None)
            self._update_result_controls()

    def _set_workbook_controls(self, enabled: bool) -> None:
        readonly_state = "readonly" if enabled else "disabled"
        normal_state = "normal" if enabled else "disabled"
        self.sheet_combo.configure(state=readonly_state)
        self.column_combo.configure(state=readonly_state)
        self.term_entry.configure(state=normal_state)
        self.exact_check.configure(state=normal_state)
        self.add_button.configure(state=normal_state)
        self.search_button.configure(state=normal_state)

    def _reset_search_state(self) -> None:
        self.dataframe = None
        self.result = None
        self.page = 0
        self.clear_criteria()
        self.tree.delete(*self.tree.get_children())
        self._update_result_controls()

    def _reset_loaded_data(self) -> None:
        self._reset_search_state()
        self.sheet_var.set("")
        self.sheet_combo.configure(values=())
        self.column_var.set(ALL_COLUMNS)
        self.column_combo.configure(values=())
        self._set_workbook_controls(False)

    def _on_close(self) -> None:
        self._hide_tooltip()
        for path in self.preview_paths:
            path.unlink(missing_ok=True)
        self.executor.shutdown(wait=False, cancel_futures=True)
        self.destroy()


def main() -> int:
    if "--version" in sys.argv:
        print(__version__)
        return 0
    if "--smoke-test" in sys.argv:
        print(f"ExcelSearcher {__version__}: OK")
        return 0
    app = ExcelSearcher()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
