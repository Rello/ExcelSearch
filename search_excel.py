#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import platform
import tempfile
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
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

        # Wähle ein Theme, das sich gut anpassen lässt
        # Auf macOS / Windows native Themes, auf anderen Systemen 'clam'
        sys_plat = platform.system()
        if sys_plat == "Darwin":
            style.theme_use("aqua")      # macOS
        elif sys_plat == "Windows":
            style.theme_use("vista")     # Windows
        else:
            style.theme_use("clam")      # Linux / sonst

        # Unser eigenes Toolbar-Design
        style.configure('Toolbar.TFrame',     background='#e0e0e0')
        style.configure('Toolbar.TLabel',     background='#e0e0e0', font=("Helvetica", 10))
        style.configure('Toolbar.TEntry',     fieldbackground='#ffffff', background='#ffffff')
        style.configure('Toolbar.TCheckbutton', background='#e0e0e0')
        style.configure('Toolbar.TButton',
                        padding=4,
                        relief='flat')
        # Hover-Effekt für Buttons
        style.map('Toolbar.TButton',
                  background=[('active','#d0d0d0')])

        # --- Logo oben ---
        logo_frame = ttk.Frame(self, style='Toolbar.TFrame')
        logo_frame.pack(fill="x", pady=(10,5))
        logo = self._load_logo("logo.jpg", max_w=750, max_h=95)
        if logo:
            ttk.Label(logo_frame, image=logo, style='Toolbar.TLabel').pack(anchor="center")
            self.logo = logo

        # --- Toolbar / Suchleiste ---
        toolbar = ttk.Frame(self, style='Toolbar.TFrame')
        toolbar.pack(fill="x", padx=10, pady=(10,5))

        # Excel öffnen
        btn_open = ttk.Button(toolbar, text="Excel auswählen…",
                              command=self.load_file, style='Toolbar.TButton')
        btn_open.pack(side="left", padx=2)

        # Suchfeld
        ttk.Label(toolbar,
                  text="Suchbegriffe (Komma, Spalte=Begriff):",
                  style='Toolbar.TLabel').pack(side="left", padx=(10,0))

        self.term_entry = ttk.Entry(toolbar, style='Toolbar.TEntry', width=40)
        self.term_entry.pack(side="left", padx=(5,0))
        self.term_entry.bind('<Return>', lambda e: self.search())

        # Exact-Match Checkbox
        self.exact_var = tk.BooleanVar(False)
        chk = ttk.Checkbutton(toolbar, text="Exact match",
                              variable=self.exact_var, style='Toolbar.TCheckbutton')
        chk.pack(side="left", padx=5)

        # Search / Export / Print
        self.search_button = ttk.Button(toolbar, text="Search",
                                        command=self.search, style='Toolbar.TButton')
        self.search_button.pack(side="left", padx=5)
        ttk.Button(toolbar, text="Export CSV",
                   command=self.export_csv, style='Toolbar.TButton').pack(side="left", padx=5)
        ttk.Button(toolbar, text="Print",
                   command=self.print_results, style='Toolbar.TButton').pack(side="left", padx=5)

        # Mauszeiger-Hand
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
        """ lädt logo.jpg aus Resources und skaliert """
        # Resampling-Filter (Pillow 10+ kompatibel)
        try:
            RES = Image.Resampling.LANCZOS
        except AttributeError:
            RES = Image.LANCZOS

        path = self.resource_path(filename)
        if not os.path.exists(path):
            return None
        try:
            img = Image.open(path)
            ratio = min(max_w/img.width, max_h/img.height, 1)
            img = img.resize((int(img.width*ratio), int(img.height*ratio)), RES)
            return ImageTk.PhotoImage(img)
        except Exception:
            return None

    def resource_path(self, rel):
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

    # --- Restliche Methoden load_file, search, export_csv, print_results wie gehabt ---
    #    (kürze ich hier der Übersicht halber)

if __name__ == "__main__":
    app = ExcelSearcher()
    app.mainloop()
