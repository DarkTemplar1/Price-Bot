# -*- coding: utf-8 -*-
"""
automat.py — silnik automatycznego przeliczania całego RAPORTU w Excelu
- sprawdza bazę danych,
- iteruje po wszystkich wierszach RAPORTU,
- zapisuje wyniki do wybranych kolumn w tym samym pliku.
Użycie (CLI):
    python automat.py <raport.xlsx> [<baza.xlsx> [<arkusz_bazy=Polska> [<poziom=Miejscowość> [<tolerancja=15>]]]]
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Optional, Dict, Tuple, List

import pandas as pd
import openpyxl  # noqa: F401 (wymagane przez pandas engine)

import automat_matma as am

# ====== Stałe / konfiguracja ======

def _default_db_path() -> Path:
    return Path.home() / "Desktop" / "baza danych" / "Baza danych.xlsx"

DEFAULT_DB_XLSX = _default_db_path()
DEFAULT_DB_SHEET = "Polska"
DEFAULT_LEVEL = "Miejscowość"
DEFAULT_TOL = 15.0

COL_MEAN_M2      = "Średnia cena za m² (z bazy)"
COL_MEAN_M2_ADJ  = "Średnia skorygowana cena za m² (z bazy)"
COL_PROP_VALUE   = "Statystyczna wartość nieruchomości"
RESULT_COLS = [COL_MEAN_M2, COL_MEAN_M2_ADJ, COL_PROP_VALUE]

MSG_NO_SIMILAR = "brak podobnych ogłoszeń"

REQUIRED_REPORT_COLUMNS = [
    "Nr KW","Typ Księgi","Stan Księgi","Województwo","Powiat","Gmina","Miejscowość","Dzielnica",
    "Położenie","Nr działek po średniku","Obręb po średniku","Ulica","Sposób korzystania","Obszar",
    "Ulica(dla budynku)","przeznaczenie (dla budynku)","Ulica(dla lokalu)","Nr budynku( dla lokalu)",
    "Przeznaczenie (dla lokalu)","Cały adres (dla lokalu)","Czy udziały?",
]

# (etykieta_GUI, kolumna_w_bazie, kolumna_w_raporcie)
ADDRESS_LEVELS = [
    ("Województwo", "wojewodztwo", "Województwo"),
    ("Powiat",       "powiat",      "Powiat"),
    ("Gmina",        "gmina",       "Gmina"),
    ("Miejscowość",  "miejscowosc", "Miejscowość"),
    ("Dzielnica",    "dzielnica",   "Dzielnica"),
    ("Ulica",        "ulica",       "Ulica"),
]

# ====== I/O: baza / raport ======

def _pick_sheet_safely(xlsx: Path, prefer: str | None = None) -> str:
    xl = pd.ExcelFile(xlsx, engine="openpyxl")
    if prefer and prefer in xl.sheet_names:
        return prefer
    return xl.sheet_names[0]

def load_db_excel(path: Path, sheet: str) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(
            f"Nie znaleziono bazy danych: {path}\n"
            "Upewnij się, że istnieje plik 'Baza danych.xlsx'."
        )
    chosen_sheet = _pick_sheet_safely(path, prefer=sheet or DEFAULT_DB_SHEET)
    df = pd.read_excel(path, sheet_name=chosen_sheet, engine="openpyxl")

    # walidacja kolumn:
    missing = [c for c in am.REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            "Brakuje kolumn w bazie: " + ", ".join(missing) +
            f"\nPlik: {path.name}, arkusz: {chosen_sheet}"
        )

    # typy/liczby/teksty
    for num_col in ["cena_za_metr","metry","rok_budowy","liczba_pokoi","pietro"]:
        df[num_col] = am._coerce_numeric(df[num_col])
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

def ensure_report_columns(xlsx: Path, sheet: str) -> pd.DataFrame:
    """Upewnij się, że w raporcie istnieją wymagane kolumny oraz kolumny wynikowe; zwróć DataFrame z arkusza."""
    df = pd.read_excel(xlsx, sheet_name=sheet, engine="openpyxl")
    for col in REQUIRED_REPORT_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    existing = list(df.columns)
    rest = [c for c in existing if c not in REQUIRED_REPORT_COLUMNS and c not in RESULT_COLS]
    new_order = REQUIRED_REPORT_COLUMNS + RESULT_COLS + rest
    df = df.reindex(columns=new_order)
    return df

# ====== Filtr + obliczenia dla jednego wiersza ======

def _filter_db(
    df_db: pd.DataFrame,
    level_key_db: str,
    level_value: str,
    area_center_str: str,
    tol: float,
) -> Tuple[pd.DataFrame, float, float, float]:
    try:
        center = float(str(area_center_str).replace(",", ".").replace(" ", ""))
    except Exception:
        center = float("nan")

    lo = center - tol if pd.notna(center) else float("-inf")
    hi = center + tol if pd.notna(center) else float("inf")

    m = am._coerce_numeric(df_db["metry"])
    mask = m.between(lo, hi)

    if str(level_value).strip():
        mask &= (df_db[level_key_db].astype(str).str.casefold() ==
                 str(level_value).strip().casefold())

    out = df_db[mask].copy()
    return out, center, lo, hi

def compute_row(
    row: pd.Series,
    df_db: pd.DataFrame,
    level_human: str,
    tol: float,
) -> Tuple[str, str, str]:
    """Zwróć (surowa_śr_m2, skorygowana_śr_m2, wartość_nieruchomości) jako sformatowane stringi."""
    mapping = {h: (db_key, rp_key) for (h, db_key, rp_key) in ADDRESS_LEVELS}
    if level_human not in mapping:
        raise ValueError(f"Nieprawidłowy poziom adresu: {level_human}")
    db_key, rp_key = mapping[level_human]

    level_value = str(row.get(rp_key, "") or "")
    obszar = str(row.get("Obszar", "") or "")

    df_filt, center, lo, hi = _filter_db(df_db, db_key, level_value, obszar, tol)

    # brak wyników albo zbyt mało podobnych
    if df_filt.empty or len(df_filt) < 5:
        return (MSG_NO_SIMILAR, MSG_NO_SIMILAR, MSG_NO_SIMILAR)

    # 1) średnia surowa
    mean_raw_m2 = am.mean_numeric(df_filt["cena_za_metr"])

    # 2) średnia po IQR
    df_clean = am.remove_outliers_iqr(df_filt.copy(), "cena_za_metr")
    mean_adj_m2 = am.mean_numeric(df_clean["cena_za_metr"])

    # 3) wartość nieruchomości
    prop_value = (mean_adj_m2 * center) if (mean_adj_m2 is not None and pd.notna(center)) else None

    return (
        am.format_price_per_m2(mean_raw_m2),
        am.format_price_per_m2(mean_adj_m2),
        am.format_currency(prop_value),
    )

# ====== Główny przebieg ======

def process_report(
    report_xlsx: Path,
    db_xlsx: Path | None = None,
    db_sheet: str = DEFAULT_DB_SHEET,
    level_human: str = DEFAULT_LEVEL,
    tol: float = DEFAULT_TOL,
) -> Tuple[int, str]:
    """
    Przetwórz cały raport: zwróć (liczba_przeliczonych_wierszy, arkusz_raportu).
    Zapis odbywa się w miejscu (ten sam plik raportu).
    """
    db_xlsx = db_xlsx or DEFAULT_DB_XLSX

    # 1) wczytaj i potwierdź bazę
    df_db = load_db_excel(Path(db_xlsx), db_sheet)

    # 2) raport + arkusz
    rp_sheet, _tmp = _pick_report_sheet(Path(report_xlsx))
    df_rp = ensure_report_columns(Path(report_xlsx), rp_sheet)

    # 3) iteracja po wierszach danych
    out = df_rp.copy()
    n = len(out.index)
    for idx in out.index:
        row = out.loc[idx]
        s_raw, s_adj, s_prop = compute_row(row, df_db, level_human=level_human, tol=tol)
        out.loc[idx, COL_MEAN_M2]     = s_raw
        out.loc[idx, COL_MEAN_M2_ADJ] = s_adj
        out.loc[idx, COL_PROP_VALUE]  = s_prop

    # 4) zapis
    with pd.ExcelWriter(Path(report_xlsx), engine="openpyxl", mode="a", if_sheet_exists="replace") as wr:
        out.to_excel(wr, sheet_name=rp_sheet, index=False)

    return n, rp_sheet

# ====== CLI ======

def _argv_or_none(i: int) -> Optional[str]:
    return sys.argv[i] if len(sys.argv) > i and sys.argv[i].strip() else None

if __name__ == "__main__":
    report_arg = _argv_or_none(1)
    if not report_arg:
        print("Użycie: python automat.py <raport.xlsx> [<baza.xlsx> [<arkusz_bazy=Polska> [<poziom=Miejscowość> [<tolerancja=15>]]]]")
        sys.exit(1)
    db_arg     = _argv_or_none(2) or str(DEFAULT_DB_XLSX)
    db_sheet   = _argv_or_none(3) or DEFAULT_DB_SHEET
    level      = _argv_or_none(4) or DEFAULT_LEVEL
    tol_str    = _argv_or_none(5) or str(DEFAULT_TOL)
    try:
        tol = float(str(tol_str).replace(",", ".").replace(" ", ""))
    except Exception:
        tol = DEFAULT_TOL

    n, sheet = process_report(Path(report_arg), Path(db_arg), db_sheet=db_sheet, level_human=level, tol=tol)
    print(f"Zrobione. Przeliczono {n} wierszy w arkuszu '{sheet}'.")
