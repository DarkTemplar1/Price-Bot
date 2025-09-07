# -*- coding: utf-8 -*-
"""
automat_gui.py — minimalistyczny GUI do uruchamiania automatu
- potwierdza poprawność bazy (kolumny + liczba wierszy),
- pozwala wskazać RAPORT oraz BAZĘ,
- jednym kliknięciem przelicza *cały* raport i zapisuje wyniki.
"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

import pandas as pd

import automat as auto

# ====== GUI ======

class AutomatApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Wycena – tryb AUTOMAT")
        self.geometry("760x420")

        self.var_db = tk.StringVar(value=str(auto.DEFAULT_DB_XLSX))
        self.var_db_sheet = tk.StringVar(value=auto.DEFAULT_DB_SHEET)
        self.var_report = tk.StringVar(value="")
        self.var_level = tk.StringVar(value=auto.DEFAULT_LEVEL)
        self.var_tol = tk.StringVar(value=str(int(auto.DEFAULT_TOL)))

        self._build()

    def _build(self):
        root = ttk.Frame(self, padding=12)
        root.pack(fill="both", expand=True)

        # Baza
        frm_db = ttk.LabelFrame(root, text="Baza danych (scalona) – Excel", padding=10)
        frm_db.pack(fill="x")
        ttk.Label(frm_db, text="Plik:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm_db, textvariable=self.var_db, width=60).grid(row=0, column=1, sticky="we", padx=6)
        ttk.Button(frm_db, text="Wybierz…", command=self._pick_db).grid(row=0, column=2)
        ttk.Label(frm_db, text="Arkusz:").grid(row=0, column=3, padx=(12, 0))
        ttk.Entry(frm_db, textvariable=self.var_db_sheet, width=16).grid(row=0, column=4, padx=6)
        ttk.Button(frm_db, text="Sprawdź bazę", command=self._check_db).grid(row=0, column=5, padx=(10, 0))
        frm_db.columnconfigure(1, weight=1)

        # Raport
        frm_rp = ttk.LabelFrame(root, text="Plik RAPORTU (Excel)", padding=10)
        frm_rp.pack(fill="x", pady=(10,0))
        ttk.Label(frm_rp, text="Plik:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm_rp, textvariable=self.var_report, width=60).grid(row=0, column=1, sticky="we", padx=6)
        ttk.Button(frm_rp, text="Wybierz…", command=self._pick_report).grid(row=0, column=2)
        frm_rp.columnconfigure(1, weight=1)

        # Parametry
        frm_p = ttk.LabelFrame(root, text="Parametry", padding=10)
        frm_p.pack(fill="x", pady=(10,0))
        ttk.Label(frm_p, text="Poziom adresu:").grid(row=0, column=0, sticky="w")
        self.cmb_level = ttk.Combobox(
            frm_p, width=20, state="readonly",
            values=[name for (name, _, _) in auto.ADDRESS_LEVELS if name != "Województwo"],
            textvariable=self.var_level
        )
        self.cmb_level.grid(row=0, column=1, padx=(6, 12), sticky="w")
        ttk.Label(frm_p, text="Tolerancja (± m²):").grid(row=0, column=2, sticky="w")
        ttk.Entry(frm_p, textvariable=self.var_tol, width=10).grid(row=0, column=3, padx=(6, 12))
        ttk.Button(frm_p, text="Uruchom AUTOMAT", command=self._run).grid(row=0, column=4)

        # Log
        self.txt = tk.Text(root, height=10, wrap="word")
        self.txt.pack(fill="both", expand=True, pady=(10,0))

    # Handlery
    def _pick_db(self):
        p = filedialog.askopenfilename(
            title="Wybierz bazę danych (Excel)",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("Wszystkie pliki", "*.*")]
        )
        if p:
            self.var_db.set(p)

    def _pick_report(self):
        p = filedialog.askopenfilename(
            title="Wybierz RAPORT (Excel)",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("Wszystkie pliki", "*.*")]
        )
        if p:
            self.var_report.set(p)

    def _check_db(self):
        try:
            df = auto.load_db_excel(Path(self.var_db.get()), self.var_db_sheet.get().strip() or auto.DEFAULT_DB_SHEET)
            # potwierdzenie „dobra baza”
            ok_cols = set(amc for amc in df.columns if amc in set(auto.RESULT_COLS + [c for (_,_,c) in auto.ADDRESS_LEVELS]))
            msg = f"OK – baza poprawna.\nWierszy: {len(df)}.\nPrzykładowe kolumny adresowe wykryte: {', '.join(sorted(ok_cols)[:8])}"
            messagebox.showinfo("Baza OK", msg)
        except Exception as e:
            messagebox.showerror("Błąd bazy", str(e))

    def _run(self):
        try:
            tol = float(str(self.var_tol.get()).replace(",", ".").replace(" ", ""))
        except Exception:
            tol = auto.DEFAULT_TOL

        try:
            n, sheet = auto.process_report(
                report_xlsx=Path(self.var_report.get()).expanduser(),
                db_xlsx=Path(self.var_db.get()).expanduser(),
                db_sheet=self.var_db_sheet.get().strip() or auto.DEFAULT_DB_SHEET,
                level_human=self.var_level.get(),
                tol=tol,
            )
            self.txt.insert("end", f"Zakończono. Przeliczono {n} wierszy w arkuszu '{sheet}'.\n")
            messagebox.showinfo("Gotowe", f"Przeliczono {n} wierszy w arkuszu '{sheet}'.")
        except Exception as e:
            messagebox.showerror("Błąd", str(e))

if __name__ == "__main__":
    app = AutomatApp()
    app.mainloop()
