#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import platform
import subprocess
import tempfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from PIL import Image, ImageTk


class ExcelSearcher(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel-Searcher")
        self.geometry("1000x650")
        self.configure(bg="white")

        # --- Style konfigurieren ---
        style = ttk.Style(self)
        sys_plat = platform.system()
        if sys_plat == "Darwin":
            style.theme_use("aqua")
        elif sys_plat == "Windows":
            style.theme_use("vista")
        else:
            style.theme_use("clam")

        style.configure('Toolbar.TFrame', background='#e0e0e0')
        style.configure('Toolbar.TLabel', background='#e0e0e0', font=("Helvetica", 10))
        style.configure('Toolbar.TEntry', fieldbackground='#ffffff', background='#ffffff')
        style.configure('Toolbar.TCheckbutton', background='#e0e0e0')
        style.configure('Toolbar.TButton', padding=4, relief='flat')
        style.map('Toolbar.TButton', background=[('active', '#d0d0d0')])

        # --- Logo oben ---
        logo_frame = ttk.Frame(self, style='Toolbar.TFrame')
        logo_frame.pack(fill="x", pady=(10, 5))
        logo = self._load_logo("logo.jpg", max_w=750, max_h=95)
        if logo:
            ttk.Label(logo_frame, image=logo, style='Toolbar.TLabel').pack(anchor="center")
            self.logo = logo

        # --- Toolbar / Suchleiste ---
        toolbar = ttk.Frame(self, style='Toolbar.TFrame')
        toolbar.pack(fill="x", padx=10, pady=(10, 5))

        btn_open = ttk.Button(
            toolbar,
            text="Excel auswählen…",
            command=self.load_file,
            style='Toolbar.TButton'
        )
        btn_open.pack(side="left", padx=2)

        ttk.Label(
            toolbar,
            text="Suchbegriffe (Komma, Spalte=Begriff optional):",
            style='Toolbar.TLabel'
        ).pack(side="left", padx=(10, 0))

        self.term_entry = ttk.Entry(toolbar, style='Toolbar.TEntry', width=40)
        self.term_entry.pack(side="left", padx=(5, 0))
        self.term_entry.bind('<Return>', lambda e: self.search())

        self.exact_var = tk.BooleanVar(False)
        chk = ttk.Checkbutton(
            toolbar,
            text="Exact match",
            variable=self.exact_var,
            style='Toolbar.TCheckbutton'
        )
        chk.pack(side="left", padx=5)

        self.search_button = ttk.Button(
            toolbar, text="Search",
            command=self.search, style='Toolbar.TButton'
        )
        self.search_button.pack(side="left", padx=5)
        ttk.Button(
            toolbar, text="Export CSV",
            command=self.export_csv, style='Toolbar.TButton'
        ).pack(side="left", padx=5)
        ttk.Button(
            toolbar, text="Print",
            command=self.print_results, style='Toolbar.TButton'
        ).pack(side="left", padx=5)

        for w in (btn_open, chk, self.search_button):
            w.configure(cursor="hand2")

        # --- Ergebnis-Tabelle ---
        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.tree = ttk.Treeview(tree_frame, show="headings")
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.df = None
        self.result = None

    def _load_logo(self, filename, max_w, max_h):
        """Lädt das Logo aus Resources und skaliert es proportional."""
        try:
            RES = Image.Resampling.LANCZOS
        except AttributeError:
            RES = Image.LANCZOS

        path = self.resource_path(filename)
        if not os.path.exists(path):
            return None
        try:
            img = Image.open(path)
            ratio = min(max_w / img.width, max_h / img.height, 1)
            new_size = (int(img.width * ratio), int(img.height * ratio))
            img = img.resize(new_size, RES)
            return ImageTk.PhotoImage(img)
        except Exception:
            return None

    def resource_path(self, rel):
        """Absoluter Pfad zu einer Resource, auch im .app- oder PyInstaller-Bundle."""
        if getattr(sys, "frozen", False):
            if hasattr(sys, "_MEIPASS"):
                base = sys._MEIPASS
            else:
                exe_dir = os.path.dirname(sys.executable)
                base = os.path.join(os.path.dirname(exe_dir), "Resources")
        else:
            base = os.path.dirname(os.path.abspath(__file__))

        p = os.path.join(base, rel)
        alt = os.path.join(base, "Resources", rel)
        return alt if not os.path.exists(p) and os.path.exists(alt) else p

    def load_file(self):
        path = filedialog.askopenfilename(
            title="Excel-Datei öffnen",
            filetypes=[("Excel-Dateien", ("*.xlsx", "*.xls")), ("Alle Dateien", "*.*")]
        )
        if not path:
            return
        try:
            self.df = pd.read_excel(path, dtype=str, engine="openpyxl")
        except Exception as e:
            messagebox.showerror("Fehler", f"Konnte Datei nicht lesen:\n{e}")
            return

        cols = list(self.df.columns)
        self.tree.config(columns=cols)
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor="w")

    def search(self):
        if self.df is None:
            messagebox.showwarning("Keine Datei", "Bitte zuerst eine Excel-Datei auswählen.")
            return

        raw_terms = [t.strip() for t in self.term_entry.get().split(",") if t.strip()]
        if not raw_terms:
            messagebox.showwarning("Keine Suchbegriffe", "Bitte mindestens einen Suchbegriff eingeben.")
            return

        self.search_button.config(text="⌛ Searching...", state="disabled")
        self.update_idletasks()

        df = self.df.fillna("").astype(str)
        mask = pd.Series(True, index=df.index)

        for raw in raw_terms:
            # Spaltensuche?
            if "=" in raw:
                col_name, raw_val = [p.strip() for p in raw.split("=", 1)]
            else:
                col_name, raw_val = None, raw

            # Exakte Suche, wenn raw_val in Hochkommas steht
            if raw_val.startswith("'") and raw_val.endswith("'") and len(raw_val) >= 2:
                val = raw_val[1:-1]
                exact = True
            else:
                val = raw_val
                exact = self.exact_var.get()

            if col_name and col_name in df.columns:
                if exact:
                    m = df[col_name] == val
                else:
                    m = df[col_name].str.contains(val, case=False, na=False)
            else:
                if exact:
                    m = df.apply(lambda row: row.str.fullmatch(val, case=False).any(), axis=1)
                else:
                    m = df.apply(lambda row: row.str.contains(val, case=False, na=False).any(), axis=1)

            mask &= m

        self.result = df[mask]
        self.tree.delete(*self.tree.get_children())
        for _, row in self.result.iterrows():
            self.tree.insert("", "end", values=list(row))

        self.search_button.config(text="Search", state="normal")

    def export_csv(self):
        if self.result is None:
            messagebox.showwarning("Keine Daten", "Bitte zuerst eine Suche durchführen.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension='.csv',
            filetypes=[('CSV-Datei', '*.csv')]
        )
        if path:
            self.result.to_csv(path, index=False)

    def print_results(self):
        if self.result is None:
            messagebox.showwarning("Keine Daten", "Bitte zuerst eine Suche durchführen.")
            return
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.csv')
        self.result.to_csv(tmp.name, index=False)
        tmp.close()
        if platform.system() == 'Windows':
            os.startfile(tmp.name, 'print')
        else:
            subprocess.run(['lp', tmp.name])


if __name__ == "__main__":
    app = ExcelSearcher()
    app.mainloop()
