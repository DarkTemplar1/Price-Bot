#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from pathlib import Path
from typing import Optional, List, Dict, Tuple

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import pandas as pd
import wyniki_matma as wm  # pomocnicze formatowanie/statystyki

# ====== Konfiguracja / stałe ======

def _default_db_path() -> Path:
    # scalanie.py zapisuje tu: ~/Desktop/baza danych/Baza danych.xlsx
    return Path.home() / "Desktop" / "baza danych" / "Baza danych.xlsx"

DEFAULT_DB_XLSX = _default_db_path()
DEFAULT_DB_SHEET = "Polska"  # jeśli brak – wybierzemy pierwszy arkusz

COL_MEAN_M2      = "Średnia cena za m² (z bazy)"                    # surowa średnia
COL_MEAN_M2_ADJ  = "Średnia skorygowana cena za m² (z bazy)"        # po IQR
COL_PROP_VALUE   = "Statystyczna wartość nieruchomości"             # skorygowana × metry
RESULT_COLS = [COL_MEAN_M2, COL_MEAN_M2_ADJ, COL_PROP_VALUE]

REQUIRED_REPORT_COLUMNS = [
    "Nr KW","Typ Księgi","Stan Księgi","Województwo","Powiat","Gmina","Miejscowość","Dzielnica",
    "Położenie","Nr działek po średniku","Obręb po średniku","Ulica","Sposób korzystania","Obszar",
    "Ulica(dla budynku)","przeznaczenie (dla budynku)","Ulica(dla lokalu)","Nr budynku( dla lokalu)",
    "Przeznaczenie (dla lokalu)","Cały adres (dla lokalu)","Czy udziały?",
]

# mapowanie poziomów adresu: (etykieta_GUI, kolumna_w_bazie, kolumna_w_raporcie)
ADDRESS_LEVELS = [
    ("Województwo", "wojewodztwo", "Województwo"),
    ("Powiat",       "powiat",      "Powiat"),
    ("Gmina",        "gmina",       "Gmina"),
    ("Miejscowość",  "miejscowosc", "Miejscowość"),
    ("Dzielnica",    "dzielnica",   "Dzielnica"),
    ("Ulica",        "ulica",       "Ulica"),
]

# ====== I/O – wczytywanie bazy i raportu ======

def _pick_sheet_safely(xlsx: Path, prefer: str | None = None) -> str:
    xl = pd.ExcelFile(xlsx, engine="openpyxl")
    if prefer and prefer in xl.sheet_names:
        return prefer
    # jeśli nie ma „Polska” – weź pierwszy arkusz
    return xl.sheet_names[0]

def load_db_excel(path: Path, sheet: str) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(
            f"Nie znaleziono bazy danych: {path}\n"
            "Upewnij się, że najpierw uruchomiłeś scalanie.py (Scalanie), "
            "które tworzy plik 'Baza danych.xlsx' na Pulpicie."
        )

    # dobierz arkusz bezpiecznie (preferuj „Polska”)
    try:
        chosen_sheet = _pick_sheet_safely(path, prefer=sheet or DEFAULT_DB_SHEET)
    except Exception as e:
        raise RuntimeError(f"Nie udało się odczytać arkuszy z: {path}\n{e}")

    df = pd.read_excel(path, sheet_name=chosen_sheet, engine="openpyxl")

    required = [
        "cena","cena_za_metr","metry","liczba_pokoi","pietro","rynek","rok_budowy","material",
        "wojewodztwo","powiat","gmina","miejscowosc","dzielnica","ulica","link",
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            "Brakuje kolumn w bazie: " + ", ".join(missing) +
            f"\nPlik: {path.name}, arkusz: {chosen_sheet}"
        )

    # typy/liczby
    for num_col in ["cena_za_metr","metry","rok_budowy","liczba_pokoi","pietro"]:
        df[num_col] = wm._coerce_numeric(df[num_col])  # type: ignore[attr-defined]
    for txt_col in ["wojewodztwo","powiat","gmina","miejscowosc","dzielnica","ulica"]:
        df[txt_col] = df[txt_col].astype(str)

    return df

def _normalize_header_map(cols: List[str]) -> Dict[str, str]:
    def norm(s: str) -> str:
        return str(s).strip().casefold().replace("  ", " ").replace("\u00a0", " ")
    return {norm(c): c for c in cols}

def _pick_report_sheet(xlsx: Path) -> Tuple[str, pd.DataFrame]:
    xl = pd.ExcelFile(xlsx, engine="openpyxl")
    best_name = None
    best_df = None
    for name in xl.sheet_names:
        df = xl.parse(name)
        norm = _normalize_header_map(list(df.columns))
        if "nr kw" in norm:
            if "obszar" in norm:
                return name, df
            if best_name is None:
                best_name, best_df = name, df
    if best_name is None:
        raise ValueError("Nie znaleziono w raporcie arkusza z kolumną 'Nr KW'.")
    return best_name, best_df  # type: ignore[return-value]

def ensure_report_columns_and_append_results(xlsx: Path, sheet: str) -> pd.DataFrame:
    df = pd.read_excel(xlsx, sheet_name=sheet, engine="openpyxl")
    for col in REQUIRED_REPORT_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    existing = list(df.columns)
    rest = [c for c in existing if c not in REQUIRED_REPORT_COLUMNS and c not in RESULT_COLS]
    new_order = REQUIRED_REPORT_COLUMNS + RESULT_COLS + rest
    df = df.reindex(columns=new_order)
    with pd.ExcelWriter(xlsx, engine="openpyxl", mode="a", if_sheet_exists="replace") as wr:
        df.to_excel(wr, sheet_name=sheet, index=False)
    return df

def _get_report_address_first_row(df: pd.DataFrame) -> Dict[str, str]:
    if df.empty:
        return {k: "" for _, _, k in ADDRESS_LEVELS} | {"Nr KW": "", "Obszar": ""}
    first = df.iloc[0]
    def get(col: str) -> str:
        return str(first.get(col, "")).strip()
    out = {"Nr KW": get("Nr KW"), "Obszar": get("Obszar")}
    for _, _, rp_key in ADDRESS_LEVELS:
        out[rp_key] = get(rp_key)
    return out

# ====== Filtrowanie i liczenie ======

def _filter_db_by_level_and_area(
    df_db: pd.DataFrame,
    level_key_db: str,
    level_value: str,
    area_center_str: str,
    tol_str: str,
) -> Tuple[pd.DataFrame, float, float, float]:
    try:
        center = float(str(area_center_str).replace(",", ".").replace(" ", ""))
    except Exception:
        center = float("nan")
    try:
        tol = float(str(tol_str).replace(",", ".").replace(" ", ""))
    except Exception:
        tol = 0.0

    lo = center - tol if pd.notna(center) else float("-inf")
    hi = center + tol if pd.notna(center) else float("inf")

    m = wm._coerce_numeric(df_db["metry"])  # type: ignore[attr-defined]
    mask = m.between(lo, hi)

    if str(level_value).strip():
        mask &= (df_db[level_key_db].astype(str).str.casefold() ==
                 str(level_value).strip().casefold())

    out = df_db[mask].copy()
    return out, center, lo, hi

def count_offers_hierarchical(
    df_db: pd.DataFrame,
    area_center_str: str,
    tol_str: str,
    values: Dict[str, str],
) -> Dict[str, int]:
    try:
        center = float(str(area_center_str).replace(",", ".").replace(" ", ""))
    except Exception:
        center = None
    try:
        tol = float(str(tol_str).replace(",", ".").replace(" ", ""))
    except Exception:
        tol = 0.0

    if center is None:
        return {name: 0 for (name, _, _) in ADDRESS_LEVELS}

    lo = center - tol
    hi = center + tol
    m = wm._coerce_numeric(df_db["metry"])  # type: ignore[attr-defined]
    base = df_db[m.between(lo, hi)].copy()

    counts: Dict[str, int] = {}
    current = base
    for human, db_key, rp_key in ADDRESS_LEVELS:
        val = (values.get(rp_key) or "").strip()
        if val:
            current = current[current[db_key].astype(str).str.casefold() == val.casefold()]
        counts[human] = int(len(current))
    return counts

# ====== GUI ======

class Aplikacja(tk.Tk):
    def __init__(self, report_arg: Optional[str] = None, db_arg: Optional[str] = None, db_sheet_arg: Optional[str] = None):
        super().__init__()
        self.title("Wycena mieszkania – wyniki (Excel)")
        self.geometry("980x760")

        # ścieżki
        self.report_path: Optional[Path] = Path(report_arg).expanduser() if report_arg else None
        # domyślnie: scalony plik adresowy (Baza danych.xlsx / Polska)
        self.db_path: Optional[Path] = Path(db_arg).expanduser() if db_arg else DEFAULT_DB_XLSX
        self.db_sheet: str = db_sheet_arg or DEFAULT_DB_SHEET

        # dane
        self.df_db: Optional[pd.DataFrame] = None
        self.report_sheet: Optional[str] = None
        self.df_report: Optional[pd.DataFrame] = None

        # wartości z raportu
        self.address_values: Dict[str, str] = {}
        self.obszar: str = ""

        # UI-stany
        self.var_db = tk.StringVar(value=str(self.db_path) if self.db_path else "")
        self.var_db_sheet = tk.StringVar(value=self.db_sheet)
        self.var_report = tk.StringVar(value=str(self.report_path) if self.report_path else "")
        self.var_tol = tk.StringVar(value="15")

        # wybór poziomu adresu do obliczeń średniej
        self.var_level = tk.StringVar(value="Miejscowość")

        # edytowalne pola „Adres do sprawdzenia”
        self.var_woj = tk.StringVar()
        self.var_pow = tk.StringVar()
        self.var_gm = tk.StringVar()
        self.var_msc = tk.StringVar()
        self.var_dz = tk.StringVar()
        self.var_ul = tk.StringVar()

        self._build_ui()

        # auto-load jeśli podano argv, w innym wypadku spróbuj domyślny scalony plik
        if self.db_path and self.db_path.exists():
            self._load_db()
        if self.report_path and self.report_path.exists():
            self._load_report()

    # --- UI budowa ---
    def _build_ui(self):
        kontener = ttk.Frame(self, padding=12)
        kontener.pack(fill="both", expand=True)

        # Baza danych (adresowa)
        frm_db = ttk.LabelFrame(kontener, text="Baza adresowa (scalona) – Excel", padding=10)
        frm_db.pack(fill="x")
        ttk.Label(frm_db, text="Plik:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm_db, textvariable=self.var_db, width=70).grid(row=0, column=1, sticky="we", padx=6)
        ttk.Button(frm_db, text="Wybierz…", command=self._pick_db).grid(row=0, column=2)
        ttk.Label(frm_db, text="Arkusz:").grid(row=0, column=3, padx=(12, 0))
        ttk.Entry(frm_db, textvariable=self.var_db_sheet, width=18).grid(row=0, column=4, padx=6)
        ttk.Button(frm_db, text="Wczytaj bazę", command=self._load_db).grid(row=0, column=5)
        frm_db.columnconfigure(1, weight=1)

        # Raport
        frm_rp = ttk.LabelFrame(kontener, text="Raport (Excel) – plik źródłowy", padding=10)
        frm_rp.pack(fill="x", pady=(10, 0))
        ttk.Label(frm_rp, text="Plik:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm_rp, textvariable=self.var_report, width=70).grid(row=0, column=1, sticky="we", padx=6)
        ttk.Button(frm_rp, text="Wybierz…", command=self._pick_report).grid(row=0, column=2)
        ttk.Button(frm_rp, text="Wczytaj raport", command=self._load_report).grid(row=0, column=3, padx=(10, 0))
        frm_rp.columnconfigure(1, weight=1)

        # Parametry obliczeń
        frm_calc = ttk.LabelFrame(kontener, text="Parametry obliczeń", padding=10)
        frm_calc.pack(fill="x", pady=(10, 0))
        ttk.Label(frm_calc, text="Poziom adresu:").grid(row=0, column=0, sticky="w")
        self.cmb_level = ttk.Combobox(
            frm_calc, width=20, state="readonly",
            values=[name for (name, _, _) in ADDRESS_LEVELS if name != "Województwo"],
            textvariable=self.var_level
        )
        self.cmb_level.grid(row=0, column=1, padx=(6, 12), sticky="w")
        ttk.Label(frm_calc, text="Tolerancja (± m²):").grid(row=0, column=2, sticky="w")
        ttk.Entry(frm_calc, textvariable=self.var_tol, width=10).grid(row=0, column=3, padx=(6, 12))
        ttk.Button(frm_calc, text="Policz i zapisz do RAPORTU", command=self._run_calc).grid(row=0, column=4, padx=(10, 0))

        # Info z raportu
        frm_info = ttk.LabelFrame(kontener, text="Dane z RAPORTU (1. wiersz danych)", padding=10)
        frm_info.pack(fill="x", pady=(10, 0))
        self.lbl_info = ttk.Label(frm_info, text="(wczytaj raport)")
        self.lbl_info.pack(anchor="w")

        # Adres do sprawdzenia + liczenie w zakresie metrażu
        frm_chk = ttk.LabelFrame(kontener, text="Adres do sprawdzenia (edytowalny) + liczenie ofert w zakresie metrażu", padding=10)
        frm_chk.pack(fill="x", pady=(10, 0))

        # lewa kolumna
        ttk.Label(frm_chk, text="Województwo:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm_chk, textvariable=self.var_woj, width=28).grid(row=0, column=1, padx=6, pady=2, sticky="w")

        ttk.Label(frm_chk, text="Gmina:").grid(row=1, column=0, sticky="w")
        ttk.Entry(frm_chk, textvariable=self.var_gm, width=28).grid(row=1, column=1, padx=6, pady=2, sticky="w")

        ttk.Label(frm_chk, text="Dzielnica:").grid(row=2, column=0, sticky="w")
        ttk.Entry(frm_chk, textvariable=self.var_dz, width=28).grid(row=2, column=1, padx=6, pady=2, sticky="w")

        # prawa kolumna
        ttk.Label(frm_chk, text="Powiat:").grid(row=0, column=2, sticky="w")
        ttk.Entry(frm_chk, textvariable=self.var_pow, width=28).grid(row=0, column=3, padx=6, pady=2, sticky="w")

        ttk.Label(frm_chk, text="Miejscowość:").grid(row=1, column=2, sticky="w")
        ttk.Entry(frm_chk, textvariable=self.var_msc, width=28).grid(row=1, column=3, padx=6, pady=2, sticky="w")

        ttk.Label(frm_chk, text="Ulica:").grid(row=2, column=2, sticky="w")
        ttk.Entry(frm_chk, textvariable=self.var_ul, width=28).grid(row=2, column=3, padx=6, pady=2, sticky="w")

        ttk.Button(frm_chk, text="Sprawdź", command=self._on_sprawdz).grid(row=3, column=0, columnspan=1, pady=(6, 0), sticky="w")

        self.var_counts = tk.StringVar(value="")
        ttk.Label(kontener, text="Wyniki (liczba ofert w zakresie metrażu)", padding=6).pack(anchor="w")
        self.lbl_counts = ttk.Label(kontener, textvariable=self.var_counts, justify="left")
        self.lbl_counts.pack(fill="x")

        # Wyniki kalkulacji i statystyki
        frm_wyn = ttk.LabelFrame(kontener, text="Wyniki i statystyki (obliczenia do zapisu w RAPORCIE)", padding=10)
        frm_wyn.pack(fill="both", expand=True, pady=(10, 0))
        self.txt = tk.Text(frm_wyn, height=16, wrap="word")
        self.txt.pack(fill="both", expand=True)

    # --- Handlery plików ---
    def _pick_db(self):
        p = filedialog.askopenfilename(
            title="Wybierz scalony plik bazy danych (Excel)",
            initialdir=str(DEFAULT_DB_XLSX.parent),
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("Wszystkie pliki", "*.*")]
        )
        if p:
            self.var_db.set(p)

    def _pick_report(self):
        p = filedialog.askopenfilename(
            title="Wybierz plik RAPORTU (Excel)",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("Wszystkie pliki", "*.*")]
        )
        if p:
            self.var_report.set(p)

    def _load_db(self):
        try:
            path = Path(self.var_db.get()).expanduser()
            sheet = self.var_db_sheet.get().strip() or DEFAULT_DB_SHEET
            self.df_db = load_db_excel(path, sheet)
            messagebox.showinfo("OK", f"Wczytano bazę: {path.name} / {sheet} (wierszy: {len(self.df_db)})")
        except Exception as e:
            messagebox.showerror("Błąd bazy", str(e))

    def _load_report(self):
        try:
            rp = Path(self.var_report.get()).expanduser()
            if not rp.exists():
                raise FileNotFoundError("Nie znaleziono pliku raportu.")
            sheet, _ = _pick_report_sheet(rp)
            df = ensure_report_columns_and_append_results(rp, sheet)

            self.report_path = rp
            self.report_sheet = sheet
            self.df_report = df

            vals = _get_report_address_first_row(df)
            self.address_values = vals
            self.obszar = vals.get("Obszar", "")

            # autouzupełnij formularz
            self.var_woj.set(vals.get("Województwo", ""))
            self.var_pow.set(vals.get("Powiat", ""))
            self.var_gm.set(vals.get("Gmina", ""))
            self.var_msc.set(vals.get("Miejscowość", ""))
            self.var_dz.set(vals.get("Dzielnica", ""))
            self.var_ul.set(vals.get("Ulica", ""))

            info = (f"Arkusz: {sheet} | Nr KW: {vals.get('Nr KW') or '—'} | "
                    f"{vals.get('Województwo', '—')}, {vals.get('Powiat', '—')}, {vals.get('Gmina', '—')}, "
                    f"{vals.get('Miejscowość', '—')}, {vals.get('Dzielnica', '—')}, {vals.get('Ulica', '—')} | "
                    f"Obszar: {self.obszar or '—'} m²")
            self.lbl_info.config(text=info)
            messagebox.showinfo("OK", f"Wczytano raport: {rp.name} / {sheet}")
        except Exception as e:
            messagebox.showerror("Błąd raportu", str(e))

    # --- Kalkulacje i zapis ---
    def _run_calc(self):
        if self.df_db is None:
            messagebox.showwarning("Brak bazy", "Wczytaj bazę danych (Excel) – scalony plik adresowy.")
            return
        if self.df_report is None or self.report_path is None or self.report_sheet is None:
            messagebox.showwarning("Brak raportu", "Wczytaj raport (Excel).")
            return

        mapping = {h: (db_key, rp_key) for (h, db_key, rp_key) in ADDRESS_LEVELS}
        human = self.var_level.get()
        if human not in mapping:
            messagebox.showerror("Błąd", "Nieprawidłowy poziom adresu.")
            return
        db_key, rp_key = mapping[human]

        current_values = {
            "Województwo": self.var_woj.get(),
            "Powiat": self.var_pow.get(),
            "Gmina": self.var_gm.get(),
            "Miejscowość": self.var_msc.get(),
            "Dzielnica": self.var_dz.get(),
            "Ulica": self.var_ul.get(),
        }
        level_value = current_values.get(rp_key, "")

        df_filt, center, lo, hi = _filter_db_by_level_and_area(
            self.df_db, db_key, level_value, self.obszar, self.var_tol.get()
        )

        self.txt.delete("1.0", "end")

        if df_filt.empty:
            self.txt.insert("end", f"Brak ofert w bazie dla: {human}='{level_value or '—'}' oraz metrażu w zakresie [{lo:.2f}, {hi:.2f}] m².\n")
            self._write_results_to_report(mean_raw_m2=None, mean_adj_m2=None, prop_value=None)
            return

        # 1) ŚREDNIA SUROWA (bez IQR)
        mean_raw_m2 = wm.mean_numeric(df_filt["cena_za_metr"])

        # 2) ŚREDNIA SKORYGOWANA (po IQR)
        df_clean = wm.remove_outliers_iqr(df_filt.copy(), "cena_za_metr")
        mean_adj_m2 = wm.mean_numeric(df_clean["cena_za_metr"])

        n_przed = len(df_filt)
        n_po = len(df_clean)

        # 3) WARTOŚĆ NIERUCHOMOŚCI = skorygowana średnia m2 × metry
        prop_value = (mean_adj_m2 * center) if (mean_adj_m2 is not None and pd.notna(center)) else None

        self._write_results_to_report(mean_raw_m2=mean_raw_m2, mean_adj_m2=mean_adj_m2, prop_value=prop_value)

        # prezentacja
        self.txt.insert("end", f"Poziom adresu do obliczeń: {human} = '{level_value or '—'}'\n")
        self.txt.insert("end", f"Zakres metrażu: {lo:.2f} — {hi:.2f} m²  (Obszar={center:.2f}, tol=±{float(self.var_tol.get() or 0):.2f})\n")
        self.txt.insert("end", f"Liczba ofert w zakresie: {n_przed} (przed IQR), {n_po} po czyszczeniu IQR.\n\n")
        self.txt.insert("end", f"Średnia cena za m² (surowa): {wm.format_price_per_m2(mean_raw_m2)}\n")
        self.txt.insert("end", f"Średnia skorygowana cena za m² (po IQR): {wm.format_price_per_m2(mean_adj_m2)}\n")
        self.txt.insert("end", f"Statystyczna wartość nieruchomości: {wm.format_currency(prop_value)}\n\n")
        self.txt.insert("end", "Wyniki zapisane w raporcie (w 1. wierszu danych) w kolumnach:\n"
                               f"  - {COL_MEAN_M2}\n  - {COL_MEAN_M2_ADJ}\n  - {COL_PROP_VALUE}\n")

    def _write_results_to_report(self, mean_raw_m2: Optional[float], mean_adj_m2: Optional[float], prop_value: Optional[float]):
        if self.report_path is None or self.report_sheet is None:
            return
        rp = self.report_path
        sh = self.report_sheet

        df = ensure_report_columns_and_append_results(rp, sh)

        s_mean_raw = wm.format_price_per_m2(mean_raw_m2) if mean_raw_m2 is not None else "—"
        s_mean_adj = wm.format_price_per_m2(mean_adj_m2) if mean_adj_m2 is not None else "—"
        s_prop_val = wm.format_currency(prop_value) if prop_value is not None else "—"

        if df.empty:
            df = pd.DataFrame(columns=df.columns)
            df.loc[0, :] = ""

        df.loc[df.index[0], COL_MEAN_M2]     = s_mean_raw
        df.loc[df.index[0], COL_MEAN_M2_ADJ] = s_mean_adj
        df.loc[df.index[0], COL_PROP_VALUE]  = s_prop_val

        with pd.ExcelWriter(rp, engine="openpyxl", mode="a", if_sheet_exists="replace") as wr:
            df.to_excel(wr, sheet_name=sh, index=False)

    # --- Sprawdzacz ilości ogłoszeń ---
    def _on_sprawdz(self):
        if self.df_db is None:
            messagebox.showwarning("Brak bazy", "Najpierw wczytaj scalony plik bazy danych (Excel).")
            return
        if not self.obszar:
            messagebox.showwarning("Brak metrażu", "Wczytaj raport, aby pobrać wartość „Obszar”.")
            return

        values = {
            "Województwo": self.var_woj.get(),
            "Powiat": self.var_pow.get(),
            "Gmina": self.var_gm.get(),
            "Miejscowość": self.var_msc.get(),
            "Dzielnica": self.var_dz.get(),
            "Ulica": self.var_ul.get(),
        }
        counts = count_offers_hierarchical(self.df_db, self.obszar, self.var_tol.get(), values)

        lines = [
            f"Zakres metrażu: ±{self.var_tol.get()} m² wokół {self.obszar} m²",
            f"• Województwo: {counts.get('Województwo', 0)}",
            f"• Powiat:      {counts.get('Powiat', 0)}",
            f"• Gmina:       {counts.get('Gmina', 0)}",
            f"• Miejscowość: {counts.get('Miejscowość', 0)}",
            f"• Dzielnica:   {counts.get('Dzielnica', 0)}",
            f"• Ulica:       {counts.get('Ulica', 0)}",
        ]
        self.var_counts.set("\n".join(lines))


def _argv_or_none(i: int) -> Optional[str]:
    return sys.argv[i] if len(sys.argv) > i and sys.argv[i].strip() else None


if __name__ == "__main__":
    report_arg = _argv_or_none(1)
    db_arg = _argv_or_none(2)
    db_sheet_arg = _argv_or_none(3)

    app = Aplikacja(report_arg=report_arg, db_arg=db_arg, db_sheet_arg=db_sheet_arg)
    app.mainloop()
