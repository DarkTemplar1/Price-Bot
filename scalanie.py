#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from pathlib import Path
from typing import List

import pandas as pd

# ===== Konfiguracja ścieżek =====
def _desktop() -> Path:
    # Uniwersalnie: Windows/Mac/Linux
    return Path.home() / "Desktop"

def _base_dir() -> Path:
    # wspiera obie wersje nazwy folderu
    d = _desktop()
    for name in ("baza danych", "Baza danych"):
        p = d / name
        if p.exists():
            return p
    # domyślnie utwórz małymi
    p = d / "baza danych"
    p.mkdir(parents=True, exist_ok=True)
    return p

SRC_DIR = _base_dir() / "województwa"
DST_FILE = _base_dir() / "Baza danych.xlsx"
DST_SHEET = "Polska"

# kanoniczny układ kolumn (jeśli dostępny w danych)
CANON_COLS: List[str] = [
    "cena",
    "cena_za_metr",
    "metry",
    "liczba_pokoi",
    "pietro",
    "rynek",
    "rok_budowy",
    "material",
    "wojewodztwo",
    "powiat",
    "gmina",
    "miejscowosc",
    "dzielnica",
    "ulica",
    "link",
]

# ===== UI powiadomienia (opcjonalnie) =====
def _info(msg: str) -> None:
    try:
        import tkinter as tk
        from tkinter import messagebox
        r = tk.Tk(); r.withdraw()
        messagebox.showinfo("Scalanie – OK", msg, parent=r)
        r.destroy()
    except Exception:
        print(msg)

def _error(msg: str) -> None:
    try:
        import tkinter as tk
        from tkinter import messagebox
        r = tk.Tk(); r.withdraw()
        messagebox.showerror("Scalanie – błąd", msg, parent=r)
        r.destroy()
    except Exception:
        print("BŁĄD:", msg, file=sys.stderr)

# ===== Główna logika =====
def _read_all_excels_from_folder(folder: Path) -> list[pd.DataFrame]:
    if not folder.exists():
        raise FileNotFoundError(f"Nie znaleziono folderu: {folder}")

    files = sorted(list(folder.glob("*.xlsx")) + list(folder.glob("*.xlsm")))
    if not files:
        raise FileNotFoundError(f"Brak plików .xlsx/.xlsm w {folder}")

    frames: list[pd.DataFrame] = []
    for f in files:
        try:
            xl = pd.ExcelFile(f, engine="openpyxl")
            for sh in xl.sheet_names:
                df = xl.parse(sh)
                # pomiń całkiem puste
                if df.empty or all(c is None for c in df.columns):
                    continue
                # usuń wiersze kompletnie puste
                df = df.dropna(how="all")
                if df.empty:
                    continue
                frames.append(df)
        except Exception as e:
            _error(f"Nie udało się wczytać pliku: {f.name}\n{e}")
    return frames

def _unify_columns(frames: list[pd.DataFrame]) -> pd.DataFrame:
    if not frames:
        return pd.DataFrame()

    # pełny zbiór kolumn
    all_cols: list[str] = []
    seen = set()
    for df in frames:
        for c in df.columns:
            cc = str(c)
            if cc not in seen:
                all_cols.append(cc)
                seen.add(cc)

    # preferuj kanoniczny układ, a resztę dorzuć na końcu
    ordered = [c for c in CANON_COLS if c in all_cols] + [c for c in all_cols if c not in CANON_COLS]

    # uzupełnij brakujące kolumny pustymi wartościami i ułóż w jedną ramkę
    normed = []
    for df in frames:
        for c in ordered:
            if c not in df.columns:
                df[c] = pd.NA
        normed.append(df[ordered])

    out = pd.concat(normed, ignore_index=True)

    # deduplikacja po 'link' jeśli kolumna istnieje
    if "link" in out.columns:
        out = out.drop_duplicates(subset=["link"], keep="first")

    # posprzątaj typowe whitespace w nagłówkach
    out.columns = [str(c).replace("\xa0", " ").strip() for c in out.columns]
    return out

def main():
    try:
        frames = _read_all_excels_from_folder(SRC_DIR)
        if not frames:
            _error("Nie znaleziono arkuszy z danymi w plikach źródłowych.")
            sys.exit(2)

        df = _unify_columns(frames)
        if df.empty:
            _error("Po scaleniu nie ma żadnych danych do zapisania.")
            sys.exit(3)

        # zapis
        DST_FILE.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(DST_FILE, engine="openpyxl", mode="w") as wr:
            df.to_excel(wr, sheet_name=DST_SHEET, index=False)

        _info(f"Scalenie zakończone.\n\nPlik: {DST_FILE}\nArkusz: {DST_SHEET}\nWierszy: {len(df)}")
    except Exception as e:
        _error(str(e))
        sys.exit(1)

if __name__ == "__main__":
    main()
