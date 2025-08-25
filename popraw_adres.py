#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
popraw_adres.py

• Przechodzi po arkuszach: „raport” (jeśli istnieje) oraz „raport_odfiltrowane” (jeśli istnieje).
• Upewnia się, że istnieją kolumny: Województwo, Powiat, Gmina, Miejscowość, Dzielnica
  (jeśli którejś brakuje — dodaje ją).
• Czyści wartości (trim, usunięcie prefiksów „woj./powiat/gmina…”, normalizacja wielkości liter).
• Województwo: weryfikacja względem listy 16 województw, tolerancyjne dopasowanie
  (bez polskich znaków, lower). Przykłady typu „płockie” zostaną odrzucone lub
  zmapowane, jeśli da się pewnie dopasować.
• Interaktywnie (okienka) zapyta o domyślne wartości do UZUPEŁNIENIA pustych pól
  (możesz anulować dowolne pole — wtedy nie uzupełnia).
• Zapisuje z powrotem arkusze (replace).

Użycie:
  python popraw_adres.py --in <plik.xlsx>

Opcje:
  --no-dialog    -> nie pyta o domyślne wartości (tylko czyszczenie/normalizacja)
"""

from __future__ import annotations

import sys
import re
import argparse
import unicodedata
from pathlib import Path
from difflib import get_close_matches
from typing import Dict, Optional, Tuple

import pandas as pd

VOIVODESHIPS = [
    "Dolnośląskie","Kujawsko-Pomorskie","Lubelskie","Lubuskie","Łódzkie","Małopolskie",
    "Mazowieckie","Opolskie","Podkarpackie","Podlaskie","Pomorskie","Śląskie",
    "Świętokrzyskie","Warmińsko-Mazurskie","Wielkopolskie","Zachodniopomorskie",
]

# -------- tekst/normalizacja ----------
def _strip_accents(s: str) -> str:
    return "".join(ch for ch in unicodedata.normalize("NFD", s) if unicodedata.category(ch) != "Mn")

def _norm_key(s: str) -> str:
    return re.sub(r"\s+", " ", _strip_accents(str(s or "")).strip().lower())

def _title_pl(s: str) -> str:
    s = str(s or "").strip()
    if not s:
        return s
    # woj/pow/gm piszemy jak w słowniku; resztę tytułujemy
    return s[:1].upper() + s[1:].lower()

VOIV_K2CANON = {_norm_key(v): v for v in VOIVODESHIPS}

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

def _canon_woj(txt: str) -> Tuple[str, bool]:
    """Zwraca (kanoniczna_nazwa_lub_wejście, czy_poprawne)."""
    raw = _clean_prefixes(txt, "woj")
    key = _norm_key(raw)
    if not key:
        return "", False
    # idealne dopasowanie
    if key in VOIV_K2CANON:
        return VOIV_K2CANON[key], True
    # miękkie dopasowanie (np. 'slaskie', 'mazowieckie   ')
    cand = get_close_matches(key, list(VOIV_K2CANON.keys()), n=1, cutoff=0.85)
    if cand:
        return VOIV_K2CANON[cand[0]], True
    # znane błędy typu "plockie" -> (Płock jest w Mazowieckim; nie zakładajmy tego
    # automatycznie). Zwracamy wyczyszczony tytuł – oznaczymy jako niepewne.
    return _title_pl(raw), False

# --------- IO ----------
ADDR_COLS = ["Województwo", "Powiat", "Gmina", "Miejscowość", "Dzielnica"]
SHEET_RAPORT = "raport"
SHEET_ODF = "raport_odfiltrowane"

def _ensure_columns(df: pd.DataFrame) -> pd.DataFrame:
    for c in ADDR_COLS:
        if c not in df.columns:
            df[c] = ""
    return df

def _ask_defaults() -> Optional[Dict[str, str]]:
    try:
        import tkinter as tk
        from tkinter import simpledialog
    except Exception:
        return None
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    fields = [
        ("Województwo", "Województwo"),
        ("Powiat", "Powiat"),
        ("Gmina", "Gmina"),
        ("Miejscowość", "Miejscowość"),
        ("Dzielnica", "Dzielnica"),
    ]
    out: Dict[str, str] = {}
    for label, key in fields:
        v = simpledialog.askstring("Uzupełnij braki (opcjonalnie)", f"{label}:")
        if v:
            out[key] = v.strip()
    try:
        root.destroy()
    except Exception:
        pass
    return out or None

def _normalize_address_block(df: pd.DataFrame, ask_defaults: bool) -> Tuple[pd.DataFrame, Dict[str, int]]:
    stats = {"woj_fixed": 0, "woj_invalid": 0, "filled": 0}
    df = _ensure_columns(df.copy())

    # przygotuj domyślne uzupełnienia (opcjonalnie)
    defaults = _ask_defaults() if ask_defaults else None

    # WOJEWÓDZTWO – walidacja/kanonizacja
    new_woj = []
    for val in df["Województwo"].astype(str).tolist():
        canon, ok = _canon_woj(val)
        if not val.strip() and defaults and defaults.get("Województwo"):
            canon, ok = _canon_woj(defaults["Województwo"])
            stats["filled"] += 1
        if ok:
            if canon != val:
                stats["woj_fixed"] += 1
            new_woj.append(canon)
        else:
            if val.strip():
                stats["woj_invalid"] += 1
            new_woj.append(_title_pl(_clean_prefixes(val, "woj")))
    df["Województwo"] = new_woj

    # POZOSTAŁE — czyszczenie prefiksów + tytułowanie + uzupełnianie pustych domyślnymi
    def _fix(col: str, level: str):
        nonlocal df, stats, defaults
        series = []
        for v in df[col].astype(str).tolist():
            w = _title_pl(_clean_prefixes(v, level))
            if not w and defaults and defaults.get(col):
                w = _title_pl(_clean_prefixes(defaults[col], level))
                stats["filled"] += 1
            series.append(w)
        df[col] = series

    _fix("Powiat", "pow")
    _fix("Gmina", "gm")
    _fix("Miejscowość", "")
    _fix("Dzielnica", "dz")

    return df, stats

def _process_sheet(xlsx: Path, sheet: str, ask_defaults: bool) -> Dict[str, int]:
    try:
        df = pd.read_excel(xlsx, sheet_name=sheet, engine="openpyxl")
    except Exception:
        return {"skipped": 1}

    df2, st = _normalize_address_block(df, ask_defaults)
    if not df2.equals(df):
        with pd.ExcelWriter(xlsx, engine="openpyxl", mode="a", if_sheet_exists="replace") as wr:
            df2.to_excel(wr, sheet_name=sheet, index=False)
    st["updated"] = int(not df2.equals(df))
    return st

def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="path", required=True, help="Ścieżka do pliku Excel")
    ap.add_argument("--no-dialog", action="store_true", help="Nie pytaj o wartości domyślne dla pustych pól")
    args = ap.parse_args(argv)

    xlsx = Path(args.path).expanduser()
    if not xlsx.exists():
        print(f"[ERR] Brak pliku: {xlsx}")
        sys.exit(1)

    sheets_to_touch = []
    # priorytet „raport”, potem „raport_odfiltrowane”, a jak brak – pierwszy arkusz
    try:
        xl = pd.ExcelFile(xlsx, engine="openpyxl")
        if SHEET_RAPORT in xl.sheet_names:
            sheets_to_touch.append(SHEET_RAPORT)
        else:
            sheets_to_touch.append(xl.sheet_names[0])
        if SHEET_ODF in xl.sheet_names:
            sheets_to_touch.append(SHEET_ODF)
    except Exception as e:
        print(f"[ERR] Nie można otworzyć Excela: {e}")
        sys.exit(2)

    total = {"updated": 0, "woj_fixed": 0, "woj_invalid": 0, "filled": 0}
    for sh in sheets_to_touch:
        st = _process_sheet(xlsx, sh, ask_defaults=(not args.no_dialog))
        for k, v in st.items():
            total[k] = total.get(k, 0) + int(v)

    print("[OK] Zakończono normalizację adresów.")
    print(f"  Zmienione arkusze: {total.get('updated',0)}")
    print(f"  Poprawione województwa (kanonizacja): {total.get('woj_fixed',0)}")
    print(f"  Podejrzane województwa (niekanoniczne): {total.get('woj_invalid',0)}")
    print(f"  Uzupełnione puste pola (domyślne): {total.get('filled',0)}")

if __name__ == "__main__":
    main()
