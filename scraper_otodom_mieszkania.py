#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations
import argparse
import csv
import time
import sys
from pathlib import Path
from typing import Iterable, List, Dict
from secrets import SystemRandom

# alias do modułu z pobieraniem danych z pojedynczego ogłoszenia
import scraper_otodom as scpr

BASE_DIR = Path(__file__).resolve().parent

# kolumny zapisu (CSV)
RESULT_HEADERS = [
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

_rng = SystemRandom()
DELAY_MIN = 3.0
DELAY_MAX = 4.0

def _sleep_random():
    delay = _rng.uniform(DELAY_MIN, DELAY_MAX)
    print(f"[sleep] odczekuję ~{delay:.2f} s…", flush=True)
    time.sleep(delay)

def _read_links_from_csv(path: Path) -> List[str]:
    links: List[str] = []
    if not path.exists() or path.stat().st_size == 0:
        print(f"[WARN] Brak pliku z linkami: {path}")
        return links
    with path.open("r", newline="", encoding="utf-8") as f:
        r = csv.reader(f)
        next(r, None)  # pomiń nagłówek
        for row in r:
            if not row:
                continue
            u = (row[0] or "").strip()
            if u:
                links.append(u)
    return links

def _ensure_row_keys(row: Dict[str, str]) -> Dict[str, str]:
    return {k: (row.get(k) or "").strip() for k in RESULT_HEADERS}

def _existing_result_links(out_csv: Path) -> set[str]:
    """Zwraca zestaw linków już obecnych w pliku wynikowym (aby nie dublować)."""
    existing: set[str] = set()
    if not out_csv.exists() or out_csv.stat().st_size == 0:
        return existing
    with out_csv.open("r", newline="", encoding="utf-8-sig") as f:
        r = csv.DictReader(f)
        for row in r:
            u = (row.get("link") or "").strip()
            if u:
                existing.add(u)
    return existing

def _append_rows_to_csv(out_csv: Path, rows: List[Dict[str, str]]) -> int:
    """Dopisuje wiersze do CSV, bez duplikacji po 'link'. Zwraca liczbę NOWO dodanych."""
    out_csv.parent.mkdir(parents=True, exist_ok=True)
    existing_links = _existing_result_links(out_csv)
    new_rows = [r for r in rows if (r.get("link") or "") not in existing_links]
    if not new_rows:
        return 0
    write_header = not out_csv.exists() or out_csv.stat().st_size == 0
    with out_csv.open("a", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=RESULT_HEADERS)
        if write_header:
            w.writeheader()
        for r in new_rows:
            w.writerow(_ensure_row_keys(r))
    return len(new_rows)

def _process_links_to_rows(links: Iterable[str]) -> List[Dict[str, str]]:
    results: List[Dict[str, str]] = []
    for idx, link in enumerate(links, start=1):
        print(f"\n[{idx}] Przetwarzam: {link}")
        try:
            data = None
            for attempt in range(3):
                try:
                    data = scpr.pobierz_dane_z_otodom(link)
                    break
                except Exception as e:
                    try:
                        transient = scpr._is_remote_closed(e)  # type: ignore[attr-defined]
                    except Exception:
                        s = (str(e) or "").lower()
                        transient = any(m in s for m in (
                            "remote end closed", "remotedisconnected", "connection reset", "chunked encoding"
                        ))
                    if transient and attempt < 2:
                        print(f"[WARN] Tymczasowy problem sieciowy: {e} — ponawiam…")
                        _sleep_random()
                        continue
                    print(f"[ERR] Nie udało się pobrać danych: {e}")
                    data = None
                    break

            if not data:
                _sleep_random()
                continue

            if not (data.get("cena") or "").strip():
                print("⏭️  Pusta cena – pomijam rekord.")
                _sleep_random()
                continue

            row = {
                "cena": data.get("cena", ""),
                "cena_za_metr": data.get("cena_za_metr", ""),
                "metry": data.get("metry", ""),
                "liczba_pokoi": data.get("liczba_pokoi", ""),
                "pietro": data.get("pietro", ""),
                "rynek": data.get("rynek", ""),
                "rok_budowy": data.get("rok_budowy", ""),
                "material": data.get("material", ""),
                "wojewodztwo": data.get("wojewodztwo", ""),
                "powiat": data.get("powiat", ""),
                "gmina": data.get("gmina", ""),
                "miejscowosc": data.get("miejscowosc", ""),
                "dzielnica": data.get("dzielnica", ""),
                "ulica": data.get("ulica", ""),
                "link": data.get("link", link),
            }
            results.append(_ensure_row_keys(row))

        finally:
            _sleep_random()

    return results

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--region", "-r", required=True,
                    help="Slug województwa (np. mazowieckie) lub etykieta (np. Mazowieckie).")
    args = ap.parse_args()

    # normalizacja – przyjmujemy zarówno 'Mazowieckie' jak i 'mazowieckie'
    reg = args.region.strip().lower()
    # mapowanie uproszczone (akceptujemy też gotowy slug)
    MAP = {
        "dolnośląskie": "dolnoslaskie", "dolnoslaskie": "dolnoslaskie",
        "kujawsko-pomorskie": "kujawsko-pomorskie",
        "lubelskie": "lubelskie",
        "lubuskie": "lubuskie",
        "łódzkie": "lodzkie", "lodzkie": "lodzkie",
        "małopolskie": "malopolskie", "malopolskie": "malopolskie",
        "mazowieckie": "mazowieckie",
        "opolskie": "opolskie",
        "podkarpackie": "podkarpackie",
        "podlaskie": "podlaskie",
        "pomorskie": "pomorskie",
        "śląskie": "slaskie", "slaskie": "slaskie",
        "świętokrzyskie": "swietokrzyskie", "swietokrzyskie": "swietokrzyskie",
        "warmińsko-mazurskie": "warminsko-mazurskie", "warminsko-mazurskie": "warminsko-mazurskie",
        "wielkopolskie": "wielkopolskie",
        "zachodniopomorskie": "zachodniopomorskie",
    }
    region_slug = MAP.get(reg, reg)
    if region_slug not in MAP.values():
        raise SystemExit(f"Nieznane województwo: {args.region}")

    intake = BASE_DIR / f"intake_{region_slug}.csv"
    out_csv = BASE_DIR / f"wyniki_{region_slug}.csv"

    print(f"[INFO] Wejście:  {intake.name}")
    print(f"[INFO] Wyjście:  {out_csv.name}")

    links = _read_links_from_csv(intake)
    if not links:
        print("[INFO] Brak linków do przetworzenia.")
        sys.exit(0)

    rows = _process_links_to_rows(links)
    if not rows:
        print("[INFO] Brak poprawnych rekordów do zapisania (np. brak cen).")
        sys.exit(0)

    added = _append_rows_to_csv(out_csv, rows)
    print(f"[OK] Dopisano {added} nowych rekordów do „{out_csv.name}”.")
    print("Gotowe.")
