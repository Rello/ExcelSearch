#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import subprocess
import tempfile
import platform
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
from PIL import Image, ImageTk
import traceback

class ExcelSearcher(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel-Searcher")
        self.geometry("1000x650")
        self.configure(bg="white")

        # Einheitlicher Font
        default_font = ("Helvetica", 10)

        # TTK-Style für native macOS-Buttons
        style = ttk.Style(self)
        try:
            style.theme_use('aqua')
        except Exception:
            pass
        style.configure('TButton', padding=6, font=default_font)
        style.configure('TCheckbutton', padding=6, font=default_font)

        # --- Logo-Frame ---
        logo_frame = tk.Frame(self, bg="white")
        logo_frame.pack(fill="x", pady=(10, 5))

        logo_path = self.resource_path("logo.jpg")
        print(f"[DEBUG] logo_path resolved to: {logo_path!r}")
        print(f"[DEBUG] exists: {os.path.exists(logo_path)}")

        if os.path.exists(logo_path):
            try:
                img = Image.open(logo_path)
                print(f"[DEBUG] Original image size: {img.width}x{img.height}")
                # Maximal 750×95 px, nur verkleinern
                max_w, max_h = 750, 95
                ratio = min(max_w / img.width, max_h / img.height, 1)
                new_size = (int(img.width * ratio), int(img.height * ratio))
                print(f"[DEBUG] Resizing image with ratio {ratio:.3f} to {new_size}")
                img = img.resize(new_size, Image.ANTIALIAS)
                self.logo = ImageTk.PhotoImage(img)
                tk.Label(logo_frame, image=self.logo, bg="white").pack(anchor="center")
                print("[DEBUG] Logo image displayed successfully.")
            except Exception as e:
                print(f"[ERROR] Fehler beim Laden des Logos: {e}")
                traceback.print_exc()
        else:
            print("[WARNING] Logo-Datei nicht gefunden; überspringe Anzeige.")

        # --- Top-Frame: Dateiauswahl & Suchoptionen ---
        frm = tk.Frame(self, bg="white")
        frm.pack(fill="x", padx=10, pady=(10, 5))

        btn_open = ttk.Button(frm, text="Excel auswählen…", command=self.load_file)
        btn_open.pack(side="left", padx=2)

        tk.Label(frm,
                 text="Suchbegriffe (Komma getrennt, Spalte=Begriff optional):",
                 bg="white", font=default_font).pack(side="left", padx=(10, 0))

        self.term_entry = tk.Entry(frm, font=default_font)
        self.term_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))
        self.term_entry.bind('<Return>', lambda e: self.search())

        self.exact_var = tk.BooleanVar(value=False)
        chk = ttk.Checkbutton(frm, text="Exact match", variable=self.exact_var)
        chk.pack(side="left", padx=5)

        self.search_button = ttk.Button(frm, text="Search", command=self.search)
        self.search_button.pack(side="left", padx=5)

        btn_export = ttk.Button(frm, text="Export CSV", command=self.export_csv)
        btn_export.pack(side="left", padx=5)

        btn_print = ttk.Button(frm, text="Print", command=self.print_results)
        btn_print.pack(side="left", padx=5)

        # Mauszeiger-Hand für alle Buttons
        for widget in (btn_open, chk, self.search_button, btn_export, btn_print):
            widget.configure(cursor="hand2")

        # --- Treeview für Ergebnisse ---
        tree_frame = tk.Frame(self, bg="white")
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.tree = ttk.Treeview(tree_frame, columns=[], show="headings")
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

    def resource_path(self, rel):
        """
        Absoluter Pfad zu einer Resource, egal ob Skript, One-File oder .app-Bundle.
        """
        print(f"[DEBUG] resource_path called with rel={rel!r}")
        if getattr(sys, "frozen", False):
            print("[DEBUG] Running in frozen mode.")
            if hasattr(sys, "_MEIPASS"):
                base = sys._MEIPASS
                print(f"[DEBUG] sys._MEIPASS is {base!r}")
            else:
                exe_dir = os.path.dirname(sys.executable)
                contents = os.path.dirname(exe_dir)
                base = os.path.join(contents, "Resources")
                print(f"[DEBUG] sys.executable is {sys.executable!r}")
                print(f"[DEBUG] Derived Resources base is {base!r}")
        else:
            base = os.path.dirname(os.path.abspath(__file__))
            print(f"[DEBUG] Not frozen, using script dir as base: {base!r}")

        # Fallback bei doppelter "Resources"-Ebene
        path = os.path.join(base, rel)
        print(f"[DEBUG] Checking primary path: {path!r}")
        if not os.path.exists(path):
            alt = os.path.join(base, "Resources", rel)
            print(f"[DEBUG] Primary not found, checking alternative: {alt!r}")
            if os.path.exists(alt):
                print(f"[DEBUG] Alternative path exists: {alt!r}")
                return alt
        return path

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

        terms = [t.strip() for t in self.term_entry.get().split(",") if t.strip()]
        if not terms:
            messagebox.showwarning("Keine Suchbegriffe", "Bitte mindestens einen Suchbegriff eingeben.")
            return

        self.search_button.config(text="⌛ Searching...", state="disabled")
        self.update_idletasks()

        df = self.df.fillna("").astype(str)
        mask = pd.Series(True, index=df.index)

        for term in terms:
            col = None
            if '=' in term:
                col, val = [p.strip() for p in term.split('=', 1)]
            else:
                val = term

            if col and col in df.columns:
                m = (df[col] == val) if self.exact_var.get() else df[col].str.contains(val, case=False, na=False)
            else:
                if self.exact_var.get():
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
    print("[DEBUG] Starting ExcelSearcher application")
    print(f"[DEBUG] __file__ is {__file__!r}")
    app = ExcelSearcher()
    app.mainloop()
