import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

class ExcelSearcher(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel-Searcher")
        self.geometry("800x600")
        
        # --- Top‐Frame: Dateiauswahl & Suchbegriffe ---
        frm = tk.Frame(self)
        frm.pack(fill="x", padx=10, pady=10)
        
        tk.Button(frm, text="Excel auswählen…", command=self.load_file).pack(side="left")
        tk.Label(frm, text="Suchbegriffe (Komma getrennt):").pack(side="left", padx=(10,0))
        self.term_entry = tk.Entry(frm)
        self.term_entry.pack(side="left", fill="x", expand=True, padx=(5,0))
        tk.Button(frm, text="Search", command=self.search).pack(side="left", padx=5)
        
        # --- Treeview für Ergebnisse ---
        cols = []  # werden gesetzt, sobald Datei geladen
        self.tree = ttk.Treeview(self, columns=cols, show="headings")
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.pack(fill="both", expand=True, side="left")
        vsb.pack(fill="y", side="right")
        hsb.pack(fill="x", side="bottom")
        
        self.df = None
        
    def load_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel-Datei","*.xlsx;*.xls")],
            title="Excel-Datei öffnen"
        )
        if not path:
            return
        try:
            # Alle Spalten als String einlesen
            self.df = pd.read_excel(path, dtype=str, engine="openpyxl")
        except Exception as e:
            messagebox.showerror("Fehler", f"Konnte Datei nicht lesen:\n{e}")
            return
        
        # Treeview anpassen
        self.tree["columns"] = list(self.df.columns)
        for col in self.df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor="w")
        messagebox.showinfo("Fertig", f"Datei geladen: {path}\n{len(self.df)} Zeilen.")
        
    def search(self):
        if self.df is None:
            messagebox.showwarning("Keine Datei", "Bitte zuerst eine Excel-Datei auswählen.")
            return
        terms = [t.strip() for t in self.term_entry.get().split(",") if t.strip()]
        if not terms:
            messagebox.showwarning("Keine Suchbegriffe", "Bitte mindestens einen Suchbegriff eingeben.")
            return
        
        df = self.df.fillna("").astype(str)
        mask = pd.Series(True, index=df.index)
        for term in terms:
            # Teilwort, case-insensitive, AND-Verknüpfung
            m = df.apply(lambda row: row.str.contains(term, case=False, na=False).any(), axis=1)
            mask &= m
        
        result = df[mask]
        # Alte Einträge löschen
        self.tree.delete(*self.tree.get_children())
        # Neue einfügen
        for _, row in result.iterrows():
            self.tree.insert("", "end", values=list(row))
        
        messagebox.showinfo("Ergebnis", f"{len(result)} Zeilen gefunden.")
        
if __name__ == "__main__":
    app = ExcelSearcher()
    app.mainloop()
