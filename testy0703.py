#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Automatyczne pobranie ogłoszeń z WSZYSTKICH województw
przy użyciu istniejących skryptów:
 - linki_mieszkania.py  --region <WOJ> --output <plik.csv>
 - scraper_otodom.py    --region <WOJ> --input <linki.csv> --output <oferty.csv>

Uruchomienie (domyślnie wszystkie):
    python automat.py

Przydatne opcje:
    --only "Mazowieckie,Małopolskie"   # ogranicz do wybranych
    --force                             # nadpisuj istniejące pliki
"""

from __future__ import annotations

import argparse
import csv
import os
import sys
import time
from pathlib import Path
from typing import Iterable, List

# ===== Lista województw (spójna z Twoimi skryptami GUI) =====
WOJEWODZTWA: List[str] = [
    "Dolnośląskie", "Kujawsko-Pomorskie", "Lubelskie", "Lubuskie",
    "Łódzkie", "Małopolskie", "Mazowieckie", "Opolskie",
    "Podkarpackie", "Podlaskie", "Pomorskie", "Śląskie",
    "Świętokrzyskie", "Warmińsko-Mazurskie", "Wielkopolskie", "Zachodniopomorskie",
]

# ===== Ścieżki katalogów wyjściowych =====
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

# ===== Ścieżki do skryptów pomocniczych (w tym samym folderze co automat.py) =====
THIS_DIR = Path(__file__).resolve().parent
LINKI_SCRIPT = (THIS_DIR / "linki_mieszkania.py").resolve()
SCRAPER_SCRIPT = (THIS_DIR / "scraper_otodom.py").resolve()

def _check_scripts() -> None:
    missing = [p.name for p in (LINKI_SCRIPT, SCRAPER_SCRIPT) if not p.exists()]
    if missing:
        raise FileNotFoundError(
            "Brak wymaganych plików w folderze programu: " + ", ".join(missing)
        )

# ===== Utils =====
def _log(msg: str) -> None:
    ts = time.strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)

def _count_csv_rows(path: Path) -> int:
    if not path.exists():
        return 0
    try:
        with path.open("r", encoding="utf-8-sig", newline="") as f:
            # bezpieczne liczenie (nie zakładamy nagłówka)
            return sum(1 for _ in csv.reader(f))
    except Exception:
        # awaryjnie liczymy po liniach
        return sum(1 for _ in path.read_text(errors="ignore").splitlines())

def _run_linki(woj: str, out_csv: Path, force: bool) -> bool:
    """
    Uruchamia linki_mieszkania.py dla danego województwa.
    Zwraca True, jeśli zakończyło się poprawnie (lub pominęliśmy, bo już istnieje i nie wymuszamy).
    """
    if out_csv.exists() and not force:
        _log(f"[{woj}] Linki już istnieją → pomijam (użyj --force, aby odświeżyć)")
        return True

    cmd = [sys.executable, str(LINKI_SCRIPT), "--region", woj, "--output", str(out_csv)]
    _log(f"[{woj}] Pobieram linki…")
    try:
        import subprocess
        subprocess.run(cmd, check=True)
        n = _count_csv_rows(out_csv)
        _log(f"[{woj}] ✔ Linki zapisane ({n} wierszy) → {out_csv}")
        return True
    except subprocess.CalledProcessError as e:
        _log(f"[{woj}] ✖ Błąd pobierania linków (kod {e.returncode})")
        return False
    except Exception as e:
        _log(f"[{woj}] ✖ Błąd uruchamiania linków: {e}")
        return False

def _run_scraper(woj: str, in_csv: Path, out_csv: Path, force: bool) -> bool:
    """
    Uruchamia scraper_otodom.py dla danego województwa.
    """
    if out_csv.exists() and not force:
        _log(f"[{woj}] Wynik CSV już istnieje → pomijam (użyj --force, aby odświeżyć)")
        return True

    if not in_csv.exists() or _count_csv_rows(in_csv) == 0:
        _log(f"[{woj}] Brak linków wejściowych ({in_csv}) – pomijam scraper")
        return False

    cmd = [
        sys.executable,
        str(SCRAPER_SCRIPT),
        "--region", woj,
        "--input", str(in_csv),
        "--output", str(out_csv),
    ]
    _log(f"[{woj}] Pobieram szczegóły ogłoszeń…")
    try:
        import subprocess
        subprocess.run(cmd, check=True)
        n = _count_csv_rows(out_csv)
        _log(f"[{woj}] ✔ Dane zapisane ({n} wierszy) → {out_csv}")
        return True
    except subprocess.CalledProcessError as e:
        _log(f"[{woj}] ✖ Błąd scrapera (kod {e.returncode})")
        return False
    except Exception as e:
        _log(f"[{woj}] ✖ Błąd uruchamiania scrapera: {e}")
        return False

# ===== Główna pętla =====
def _iter_wojewodztwa(only: Iterable[str] | None) -> List[str]:
    if not only:
        return WOJEWODZTWA
    # normalizacja: usuń spacje wokół i dopasuj pełne nazwy
    wanted = {w.strip() for w in only if w.strip()}
    # pozwalamy podać nazwę bez polskich znaków – zrobimy proste mapowanie
    def _norm(s: str) -> str:
        rep = str.maketrans("ąćęłńóśźżĄĆĘŁŃÓŚŹŻ", "acelnoszzACELNOSZZ")
        return s.translate(rep).lower()

    idx = {_norm(w): w for w in WOJEWODZTWA}
    resolved: List[str] = []
    for w in wanted:
        key = _norm(w)
        if key in idx:
            resolved.append(idx[key])
        else:
            _log(f"⚠ Nie rozpoznano województwa: {w}")
    return resolved

def main():
    _check_scripts()

    parser = argparse.ArgumentParser()
    parser.add_argument("--only", help="Lista województw rozdzielona przecinkami (domyślnie: wszystkie).")
    parser.add_argument("--force", action="store_true", help="Nadpisuj istniejące pliki wynikowe.")
    parser.add_argument("--sleep", type=float, default=1.0, help="Przerwa (sekundy) między województwami.")
    args = parser.parse_args()

    selected = _iter_wojewodztwa(args.only.split(",") if args.only else None)
    if not selected:
        _log("Brak województw do przetworzenia (sprawdź parametr --only).")
        sys.exit(2)

    _log(f"Start automatu. Katalog bazowy: {BASE_DIR}")
    ok_cnt = 0
    for woj in selected:
        _log("=" * 64)
        linki_csv = LINKI_DIR / f"{woj}.csv"
        oferty_csv = WOJ_DIR / f"{woj}.csv"

        ok1 = _run_linki(woj, linki_csv, force=args.force)
        ok2 = _run_scraper(woj, linki_csv, oferty_csv, force=args.force) if ok1 else False

        if ok1 and ok2:
            ok_cnt += 1

        time.sleep(max(0.0, args.sleep))

    _log("=" * 64)
    _log(f"Zakończono. Poprawnie przetworzonych województw: {ok_cnt}/{len(selected)}")
    _log(f"Wyniki: {WOJ_DIR}")
    # (opcjonalnie) tutaj możesz wywołać skrypt scalający do Excela,
    # jeśli masz go w projekcie:
    #   from subprocess import run
    #   run([sys.executable, str((THIS_DIR/'scalanie.py').resolve())], check=False)

if __name__ == "__main__":
    main()
