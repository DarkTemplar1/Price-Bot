#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import csv
import time
import random
from pathlib import Path

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from bazadanych import pobierz_dane_z_otodom  # import parsera pojedynczego ogłoszenia

print("[SCRAPER] Uruchomiono scraper_otodom_mieszkania.py")

WAIT_SECONDS = 20
KOLUMNY_WYJ = [
    "cena", "cena_za_metr", "metry", "liczba_pokoi", "pietro", "rynek", "rok_budowy",
    "material", "wojewodztwo", "powiat", "gmina", "miejscowosc", "dzielnica", "ulica", "link",
]

# ========================= UTILS =========================
def _detect_desktop() -> Path:
    home = Path.home()
    for name in ("Desktop", "Pulpit"):
        p = home / name
        if p.exists():
            return p
    return home

BASE_BAZA = _detect_desktop() / "baza danych"
BASE_WOJ_DIR = BASE_BAZA / "województwa"
BASE_WOJ_DIR.mkdir(parents=True, exist_ok=True)

def _ensure_results_csv_label(label: str) -> Path:
    p = BASE_WOJ_DIR / f"{label}.csv"
    if not p.exists():
        with p.open("w", newline="", encoding="utf-8-sig") as f:
            csv.writer(f).writerow(KOLUMNY_WYJ)
    return p

def _read_links_from_csv(path: Path) -> list[str]:
    links = []
    with path.open("r", newline="", encoding="utf-8") as f:
        r = csv.reader(f)
        header = next(r, None)
        has_header = bool(header and ("link" in [c.strip().lower() for c in header]))
        if not has_header and header and header[0].strip():
            links.append(header[0].strip())
        for row in r:
            if row and row[0].strip():
                links.append(row[0].strip())
    return links

def _append_rows_to_output_csv(output_csv: Path, rows: list[dict]):
    with output_csv.open("a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        for r in rows:
            w.writerow([r.get(col, "") for col in KOLUMNY_WYJ])

# ========================= PIPE =========================
def przetworz_linki_do_csv(input_csv: Path, out_csv: Path, limit: int = 0):
    links = _read_links_from_csv(input_csv)
    if limit and limit > 0:
        links = links[:limit]

    total = len(links)
    print(f"[INFO] Do przetworzenia linków: {total}")
    if not links:
        return

    dane_lista: list[dict] = []
    progress_file = Path("scraper_progress.txt")
    if progress_file.exists():
        progress_file.unlink()  # wyczyść poprzedni stan

    for i, link in enumerate(links, start=1):
        print(f"\n🌐 [{i}/{total}] {link}")
        # zapisz postęp do pliku
        try:
            progress_file.write_text(f"{i}/{total}", encoding="utf-8")
        except Exception:
            pass

        try:
            dane = pobierz_dane_z_otodom(link)
        except Exception as e:
            print(f"❌ Błąd przy linku: {e}")
            dane = None

        if not dane:
            continue

        for k in KOLUMNY_WYJ:
            dane.setdefault(k, "")

        if (dane.get("cena") or "").strip():
            dane_lista.append(dane)

        # losowe opóźnienie 4–8s
        delay = random.uniform(4.0, 8.0)
        print(f"[INFO] Czekam {delay:.2f}s przed kolejnym ogłoszeniem…")
        time.sleep(delay)

    _append_rows_to_output_csv(out_csv, dane_lista)

    # usuń plik postępu po zakończeniu
    if progress_file.exists():
        progress_file.unlink()

    print(f"\n✅ Zapisano {len(dane_lista)} rekordów do {out_csv}")

# ========================= CLI =========================
if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", "-i", type=str, required=True, help="Ścieżka do pliku CSV z linkami")
    ap.add_argument("--limit", "-l", type=int, default=0, help="Limit ogłoszeń (0 = wszystkie)")
    ap.add_argument("--label", "-b", type=str, default="Województwo", help="Etykieta pliku wynikowego")
    args = ap.parse_args()

    input_csv = Path(args.input)
    out_csv = _ensure_results_csv_label(args.label)
    przetworz_linki_do_csv(input_csv, out_csv, limit=args.limit)
