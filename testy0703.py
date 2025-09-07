#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Automat: pobiera ogłoszenia z WSZYSTKICH województw i pobiera ich dane.
Korzysta z:
 - linki_mieszkania.py
 - scraper_otodom_mieszkania.py (jeśli istnieje) lub scraper_otodom.py

Domyślnie NADPISUJE istniejące pliki (brak pomijania).

Przykłady:
    python automat.py
    python automat.py --only "Mazowieckie,Małopolskie"
    python automat.py --sleep 0.5
    python automat.py --merge
"""

from __future__ import annotations

import argparse
import csv
import os
import sys
import time
from pathlib import Path
from typing import Iterable, List

# ===== Lista województw =====
WOJEWODZTWA: List[str] = [
    "Warmińsko-Mazurskie", "Wielkopolskie", "Zachodniopomorskie",
]

# ===== Ścieżki bazowe =====
def _desktop() -> Path:
    return Path.home() / "Desktop"

def _base_dir() -> Path:
    d = _desktop()
    for name in ("baza danych", "Baza danych"):
        p = d / name
        if p.exists():
            return p
    p = d / "baza danych"
    p.mkdir(parents=True, exist_ok=True)
    return p

BASE_DIR = _base_dir()
LINKI_DIR = BASE_DIR / "linki"
WOJ_DIR = BASE_DIR / "województwa"
LINKI_DIR.mkdir(parents=True, exist_ok=True)
WOJ_DIR.mkdir(parents=True, exist_ok=True)

# ===== Pliki skryptów (ten sam folder co automat.py) =====
THIS_DIR = Path(__file__).resolve().parent
LINKI_SCRIPT = (THIS_DIR / "linki_mieszkania.py").resolve()
SCRAPER_MIESZ = (THIS_DIR / "scraper_otodom_mieszkania.py").resolve()
SCRAPER_STD = (THIS_DIR / "scraper_otodom.py").resolve()
SCALANIE = (THIS_DIR / "scalanie.py").resolve()  # opcjonalnie

def _choose_scraper() -> Path:
    # Preferuj scraper_otodom_mieszkania.py, jeśli jest
    if SCRAPER_MIESZ.exists():
        return SCRAPER_MIESZ
    return SCRAPER_STD

def _check_scripts() -> None:
    missing = []
    if not LINKI_SCRIPT.exists():
        missing.append(LINKI_SCRIPT.name)
    scraper = _choose_scraper()
    if not scraper.exists():
        missing.append(scraper.name)
    if missing:
        raise FileNotFoundError("Brak wymaganych plików: " + ", ".join(missing))

# ===== Utils =====
def _log(msg: str) -> None:
    ts = time.strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)

def _count_csv_rows(path: Path) -> int:
    if not path.exists():
        return 0
    try:
        with path.open("r", encoding="utf-8-sig", newline="") as f:
            return sum(1 for _ in csv.reader(f))
    except Exception:
        return sum(1 for _ in path.read_text(errors="ignore").splitlines())

def _rm_if_exists(path: Path) -> None:
    try:
        if path.exists():
            path.unlink()
    except Exception:
        pass

def _run(cmd: list[str]) -> int:
    import subprocess
    try:
        res = subprocess.run(cmd, check=False)
        return res.returncode
    except Exception:
        return 999

# ===== Główna pętla =====
def _iter_wojewodztwa(only: Iterable[str] | None) -> List[str]:
    if not only:
        return WOJEWODZTWA
    wanted = {w.strip() for w in only if w.strip()}

    def _norm(s: str) -> str:
        rep = str.maketrans("ąćęłńóśźżĄĆĘŁŃÓŚŹŻ", "acelnoszzACELNOSZZ")
        return s.translate(rep).lower()

    idx = {_norm(w): w for w in WOJEWODZTWA}
    out: List[str] = []
    for w in wanted:
        key = _norm(w)
        if key in idx:
            out.append(idx[key])
        else:
            _log(f"⚠ Nie rozpoznano województwa: {w}")
    return out

def main():
    _check_scripts()
    parser = argparse.ArgumentParser()
    parser.add_argument("--only", help="Lista województw rozdzielona przecinkami (domyślnie: wszystkie).")
    parser.add_argument("--sleep", type=float, default=1.0, help="Przerwa (s) między województwami.")
    parser.add_argument("--merge", action="store_true", help="Po zakończeniu uruchom scalanie CSV->Excel (scalanie.py).")
    args = parser.parse_args()

    selected = _iter_wojewodztwa(args.only.split(",") if args.only else None)
    if not selected:
        _log("Brak województw do przetworzenia (sprawdź parametr --only).")
        sys.exit(2)

    scraper = _choose_scraper()
    _log(f"Start automatu. Katalog bazowy: {BASE_DIR}")
    _log(f"Używany scraper: {scraper.name}")

    ok_cnt = 0
    for woj in selected:
        _log("=" * 64)
        linki_csv = LINKI_DIR / f"{woj}.csv"
        dane_csv = WOJ_DIR / f"{woj}.csv"

        # 1) Linki – zawsze od zera (nadpisz)
        _rm_if_exists(linki_csv)
        cmd_linki = [sys.executable, str(LINKI_SCRIPT), "--region", woj, "--output", str(linki_csv)]
        _log(f"[{woj}] Pobieram linki…")
        rc1 = _run(cmd_linki)
        if rc1 != 0 or not linki_csv.exists() or _count_csv_rows(linki_csv) == 0:
            _log(f"[{woj}] ✖ Błąd pobierania linków (kod {rc1}).")
            continue
        _log(f"[{woj}] ✔ Linki: {linki_csv} ({_count_csv_rows(linki_csv)} wierszy)")

        # 2) Dane – zawsze od zera (nadpisz)
        _rm_if_exists(dane_csv)
        cmd_scr = [
            sys.executable, str(scraper),
            "--region", woj,
            "--input", str(linki_csv),
            "--output", str(dane_csv),
        ]
        _log(f"[{woj}] Pobieram dane ogłoszeń…")
        rc2 = _run(cmd_scr)
        if rc2 != 0 or not dane_csv.exists() or _count_csv_rows(dane_csv) == 0:
            _log(f"[{woj}] ✖ Błąd scrapera (kod {rc2}).")
            continue
        _log(f"[{woj}] ✔ Dane: {dane_csv} ({_count_csv_rows(dane_csv)} wierszy)")

        ok_cnt += 1
        time.sleep(max(0.0, args.sleep))

    _log("=" * 64)
    _log(f"Zakończono. Poprawnie przetworzone województwa: {ok_cnt}/{len(selected)}")
    _log(f"Wyniki CSV: {WOJ_DIR}")

    if args.merge:
        if SCALANIE.exists():
            _log("Uruchamiam scalanie CSV → Excel…")
            _run([sys.executable, str(SCALANIE)])
        else:
            _log("⚠ Nie znaleziono scalanie.py – pomijam krok scalania.")

if __name__ == "__main__":
    main()