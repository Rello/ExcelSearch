import os
import sys
import subprocess
import tempfile
import platform
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from PIL import Image, ImageTk

class ExcelSearcher(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel-Searcher")
        self.geometry("900x600")

        # Logo laden (unterstützt JPG, PNG, etc.)
        logo_path = self.resource_path("logo.jpg")  # oder "logo.png"
        if os.path.exists(logo_path):
            try:
                img = Image.open(logo_path)
                # auf max Höhe 90 px oder Breite 200 px skalieren
                max_w, max_h = 200, 90
                ratio = min(max_w / img.width, max_h / img.height, 1)
                new_size = (int(img.width * ratio), int(img.height * ratio))
                img = img.resize(new_size, Image.ANTIALIAS)
                self.logo = ImageTk.PhotoImage(img)
                tk.Label(self, image=self.logo).pack(side="top", pady=5)
            except Exception:
                pass

        # --- Top-Frame: Dateiauswahl & Suchoptionen ---
        frm = tk.Frame(self)
        frm.pack(fill="x", padx=10, pady=5)

        tk.Button(frm, text="Excel auswählen…", command=self.load_file).pack(side="left")
        tk.Label(frm, text="Suchbegriffe (Komma getrennt, Spalte=Begriff optional):").pack(side="left", padx=(10,0))
        self.term_entry = tk.Entry(frm)
        self.term_entry.pack(side="left", fill="x", expand=True, padx=(5,0))
        self.term_entry.bind('<Return>', lambda e: self.search())

        self.exact_var = tk.BooleanVar(value=False)
        tk.Checkbutton(frm, text="Exact match", variable=self.exact_var).pack(side="left", padx=5)

        self.search_button = tk.Button(frm, text="Search", command=self.search)
        self.search_button.pack(side="left", padx=5)

        tk.Button(frm, text="Export CSV", command=self.export_csv).pack(side="left", padx=5)
        tk.Button(frm, text="Print", command=self.print_results).pack(side="left", padx=5)

        # --- Treeview für Ergebnisse ---
        self.tree = ttk.Treeview(self, columns=[], show="headings")
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.pack(fill="both", expand=True, side="left")
        vsb.pack(fill="y", side="right")
        hsb.pack(fill="x", side="bottom")

        self.df = None
        self.result = None

    def resource_path(self, rel):
        if hasattr(sys, '_MEIPASS'):
            return os.path.join(sys._MEIPASS, rel)
        return os.path.join(os.path.abspath('.'), rel)

    def load_file(self):
        path = filedialog.askopenfilename(
            title="Excel-Datei öffnen",
            filetypes=[("Excel-Dateien", ("*.xlsx", "*.xls")), ("Alle Dateien", "*")]
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
            self.tree.column(col, width=100, anchor="w")

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
                col, val = [p.strip() for p in term.split('=',1)]
            else:
                val = term

            if col and col in df.columns:
                if self.exact_var.get():
                    m = df[col] == val
                else:
                    m = df[col].str.contains(val, case=False, na=False)
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
            defaultextension='.csv', filetypes=[('CSV-Datei','*.csv')]
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
