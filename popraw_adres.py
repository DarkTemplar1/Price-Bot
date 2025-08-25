#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
popraw_adres.py — automatyczna korekta pól adresowych w Excelu (z podsumowaniem w GUI)

Działanie:
- Przechodzi po arkuszach: „raport” oraz „raport_odfiltrowane” (jeśli istnieją).
  Jeżeli „raport” nie istnieje — używa pierwszego arkusza.
- Upewnia się, że istnieją kolumny: Województwo, Powiat, Gmina, Miejscowość, Dzielnica.
- Dla każdego wiersza:
    • czyści prefiksy (woj./powiat/gmina/dzielnica) i białe znaki,
    • próbuje kanonizować/uzupełnić pola przez TERYT (moduł `adres_otodom`),
    • w razie braku TERYT — heurystyki + lista 16 województw.
- Zapisuje poprawione arkusze (replace).
- Na koniec pokazuje okienko z podsumowaniem zmian.

Użycie:
    python popraw_adres.py --in <plik.xlsx>
    # opcjonalnie:
    python popraw_adres.py --in <plik.xlsx> --sheets raport,raport_odfiltrowane
"""

from __future__ import annotations

import sys
import re
import argparse
import unicodedata
from pathlib import Path
from typing import Dict, Optional, Tuple, List

import pandas as pd

# ======= Opcjonalny TERYT (adres_otodom) =======
_HAS_TERYT = False
try:
    from adres_otodom import (
        TerytClient,
        uzupelnij_braki_z_teryt,
        dopelnij_powiat_gmina_jesli_brak,
        _clean_gmina as _clean_gm_name,
    )
    _HAS_TERYT = True
except Exception:
    def _clean_gm_name(x):  # type: ignore
        return x

VOIVODESHIPS = [
    "Dolnośląskie","Kujawsko-Pomorskie","Lubelskie","Lubuskie","Łódzkie","Małopolskie",
    "Mazowieckie","Opolskie","Podkarpackie","Podlaskie","Pomorskie","Śląskie",
    "Świętokrzyskie","Warmińsko-Mazurskie","Wielkopolskie","Zachodniopomorskie",
]
VOIV_K2CANON = {}  # wypełnimy po zdefiniowaniu normalizatora

ADDR_COLS = ["Województwo", "Powiat", "Gmina", "Miejscowość", "Dzielnica"]
SHEET_RAPORT = "raport"
SHEET_ODF = "raport_odfiltrowane"


# ======= Normalizacja tekstu / prefiksów =======
def _strip_accents(s: str) -> str:
    return "".join(ch for ch in unicodedata.normalize("NFD", s) if unicodedata.category(ch) != "Mn")

def _norm_key(s: str) -> str:
    return re.sub(r"\s+", " ", _strip_accents(str(s or "")).strip().lower())

VOIV_K2CANON = {_norm_key(v): v for v in VOIVODESHIPS}

def _title_pl(s: str) -> str:
    s = str(s or "").strip()
    if not s:
        return s
    return s[:1].upper() + s[1:].lower()

def _clean_prefixes(txt: str, level: str) -> str:
    t = str(txt or "")
    t = re.sub(r"\s+", " ", t).strip()
    if not t:
        return t
    if level == "woj":
        t = re.sub(r"^(woj\.?|wojewodztwo|województwo)\s*", "", t, flags=re.I)
    elif level == "pow":
        t = re.sub(r"^(powiat|pow\.)\s*", "", t, flags=re.I)
    elif level == "gm":
        t = re.sub(r"^(gmina|gm\.)\s*(miejska|wiejska|miejsko-wiejska)?\s*", "", t, flags=re.I)
    elif level == "dz":
        t = re.sub(r"^(dzielnica|dz\.)\s*", "", t, flags=re.I)
    return t.strip()

def _ensure_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in ADDR_COLS:
        if c not in df.columns:
            df[c] = ""
    return df


# ======= Fallback (bez TERYT) =======
def _canon_woj_fallback(txt: str) -> Tuple[str, bool]:
    raw = _clean_prefixes(txt, "woj")
    key = _norm_key(raw)
    if not key:
        return "", False
    if key in VOIV_K2CANON:
        return VOIV_K2CANON[key], True
    ALIASES = {
        "slaskie": "Śląskie",
        "lodzkie": "Łódzkie",
        "malopolskie": "Małopolskie",
        "mazowieckie": "Mazowieckie",
        "wielkopolskie": "Wielkopolskie",
        "zachodniopomorskie": "Zachodniopomorskie",
        "kujawsko pomorskie": "Kujawsko-Pomorskie",
        "warminsko mazurskie": "Warmińsko-Mazurskie",
        "swietokrzyskie": "Świętokrzyskie",
    }
    if key in ALIASES:
        return ALIASES[key], True
    return _title_pl(raw), False


# ======= Korekta jednego wiersza =======
def _fix_row_with_teryt(row: pd.Series, klient: Optional["TerytClient"]) -> Dict[str, str]:
    woj = _clean_prefixes(row.get("Województwo", ""), "woj")
    powiat = _clean_prefixes(row.get("Powiat", ""), "pow")
    gmina = _clean_prefixes(row.get("Gmina", ""), "gm")
    msc = _clean_prefixes(row.get("Miejscowość", ""), "")
    dz = _clean_prefixes(row.get("Dzielnica", ""), "dz")

    out = {
        "Województwo": woj,
        "Powiat": powiat,
        "Gmina": gmina,
        "Miejscowość": msc,
        "Dzielnica": dz,
    }

    if not _HAS_TERYT or klient is None:
        woj_c, _ = _canon_woj_fallback(woj)
        out["Województwo"] = woj_c
        out["Powiat"] = _title_pl(powiat)
        out["Gmina"] = _title_pl(_clean_gm_name(gmina))
        out["Miejscowość"] = _title_pl(msc)
        out["Dzielnica"] = _title_pl(dz)
        return out

    ad = {
        "ulica_nazwa": None,
        "nr": None,
        "miasto": msc or dz,
        "dzielnica": dz,
        "wojewodztwo": woj,
        "gmina": gmina,
        "powiat": powiat,
        "oryginal": "",
    }

    try:
        ad2 = uzupelnij_braki_z_teryt(ad, klient=klient)
        ad3 = dopelnij_powiat_gmina_jesli_brak(ad2, klient=klient)
    except Exception:
        woj_c, _ = _canon_woj_fallback(woj)
        out["Województwo"] = woj_c
        out["Powiat"] = _title_pl(powiat)
        out["Gmina"] = _title_pl(_clean_gm_name(gmina))
        out["Miejscowość"] = _title_pl(msc)
        out["Dzielnica"] = _title_pl(dz)
        return out

    woj_fin, _ = _canon_woj_fallback(ad3.get("wojewodztwo") or woj)
    out["Województwo"] = woj_fin
    out["Powiat"] = _title_pl(ad3.get("powiat") or powiat)
    out["Gmina"] = _title_pl(_clean_gm_name(ad3.get("gmina") or gmina))
    out["Miejscowość"] = _title_pl(ad3.get("miasto") or msc)
    out["Dzielnica"] = _title_pl(ad3.get("dzielnica") or dz)
    return out


def _normalize_df(df: pd.DataFrame, klient: Optional["TerytClient"]) -> pd.DataFrame:
    df = _ensure_columns(df)
    if df.empty:
        return df
    fixed = []
    for _, row in df.iterrows():
        fixed.append(_fix_row_with_teryt(row, klient))
    res = df.copy()
    for c in ADDR_COLS:
        res[c] = [r[c] for r in fixed]
    return res


def _count_addr_changes(df_before: pd.DataFrame, df_after: pd.DataFrame) -> int:
    """Policz, w ilu wierszach cokolwiek zmieniło się w adresowych kolumnach."""
    cols = [c for c in ADDR_COLS if c in df_before.columns and c in df_after.columns]
    if not cols or len(df_before) != len(df_after):
        return 0
    a = df_before[cols].fillna("").astype(str)
    b = df_after[cols].fillna("").astype(str)
    return int((a != b).any(axis=1).sum())


# ======= I/O arkuszy =======
def _pick_sheets(xlsx: Path, explicit: Optional[List[str]]) -> List[str]:
    try:
        xl = pd.ExcelFile(xlsx, engine="openpyxl")
    except Exception as e:
        print(f"[ERR] Nie można otworzyć Excela: {e}")
        sys.exit(2)

    names = xl.sheet_names
    if explicit:
        return [s for s in explicit if s in names]

    out = []
    if SHEET_RAPORT in names:
        out.append(SHEET_RAPORT)
    else:
        out.append(names[0])
    if SHEET_ODF in names:
        out.append(SHEET_ODF)
    # unikatowo
    seen, uniq = set(), []
    for s in out:
        if s not in seen:
            uniq.append(s); seen.add(s)
    return uniq


def _show_summary_gui(message: str):
    """Pokaż krótkie okno z podsumowaniem. Jeśli Tkinter niedostępny — pomiń."""
    try:
        import tkinter as tk
        from tkinter import messagebox
    except Exception:
        return
    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        messagebox.showinfo("Poprawa adresów — zakończono", message, parent=root)
        root.destroy()
    except Exception:
        pass


def main(argv=None):
    ap = argparse.ArgumentParser(description="Automatyczna korekta pól adresowych (TERYT jeśli dostępny).")
    ap.add_argument("--in", dest="path", required=True, help="Ścieżka do pliku Excel")
    ap.add_argument("--sheets", dest="sheets", default=None,
                    help="Lista arkuszy po przecinku (np. raport,raport_odfiltrowane)")
    args = ap.parse_args(argv)

    xlsx = Path(args.path).expanduser()
    if not xlsx.exists():
        print(f"[ERR] Brak pliku: {xlsx}")
        sys.exit(1)

    sheets = _pick_sheets(xlsx, [s.strip() for s in args.sheets.split(",")] if args.sheets else None)

    # TERYT klient (opcjonalnie)
    klient = None
    mode = "TERYT"
    if _HAS_TERYT:
        try:
            klient = TerytClient()
        except Exception as e:
            print(f"[WARN] TERYT niedostępny ({e}) – używam trybu fallback.")
            klient = None
            mode = "fallback"
    else:
        mode = "fallback"

    updated = 0
    changed_rows_total = 0
    per_sheet_info: List[str] = []

    for sh in sheets:
        try:
            df_before = pd.read_excel(xlsx, sheet_name=sh, engine="openpyxl")
        except Exception:
            print(f"[INFO] Pomijam arkusz „{sh}” – nie można wczytać.")
            continue

        df_after = _normalize_df(df_before, klient)
        rows_changed = _count_addr_changes(df_before, df_after)

        if not df_after.equals(df_before):
            with pd.ExcelWriter(xlsx, engine="openpyxl", mode="a", if_sheet_exists="replace") as wr:
                df_after.to_excel(wr, sheet_name=sh, index=False)
            updated += 1
            changed_rows_total += rows_changed
            per_sheet_info.append(f"• {sh}: zmienione wiersze (adres) = {rows_changed}")
            print(f"[OK] Zaktualizowano arkusz „{sh}” ({rows_changed} wierszy z modyfikacją adresu).")
        else:
            per_sheet_info.append(f"• {sh}: bez zmian")
            print(f"[OK] Arkusz „{sh}” bez zmian.")

    summary = (
        f"Plik: {xlsx.name}\n"
        f"Tryb: {mode}\n"
        f"Zmienione arkusze: {updated}/{len(sheets)}\n"
        f"Suma zmienionych wierszy (adres): {changed_rows_total}\n\n"
        + "\n".join(per_sheet_info)
    )
    print("\n" + summary)

    # Pokaż podsumowanie w GUI
    _show_summary_gui(summary)


if __name__ == "__main__":
    main()
