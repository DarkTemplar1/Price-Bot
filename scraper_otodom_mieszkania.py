from __future__ import annotations

import csv
import time
import sys
from pathlib import Path
from typing import Iterable, List, Dict
from secrets import SystemRandom

# ✅ alias zmieniony na `scpr`
import scraper_otodom as scpr

from EXCELoperacje import (
    ensure_baza_mieszkania,
    append_mieszkania_rows,
    BAZA_MIESZKANIA_HEADERS,
)

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_INTAKE = BASE_DIR / "intake.csv"

DB_XLSX = BASE_DIR / "Baza_danych.xlsx"
DB_SHEET = "Mieszkania"

_rng = SystemRandom()
DELAY_MIN = 3.0
DELAY_MAX = 4.0


def _sleep_random():
    delay = _rng.uniform(DELAY_MIN, DELAY_MAX)
    print(f"[sleep] odczekuję ~{delay:.2f} s…", flush=True)
    time.sleep(delay)


def _read_links_from_csv(path: Path) -> List[str]:
    """Czyta linki z 1. kolumny intake.csv, zaczynając od 2. wiersza."""
    links: List[str] = []
    if not path.exists() or path.stat().st_size == 0:
        print(f"[WARN] Brak pliku z linkami: {path}")
        return links

    with path.open("r", newline="", encoding="utf-8") as f:
        r = csv.reader(f)
        next(r, None)  # pomiń 1. wiersz (nagłówek)
        for row in r:
            if not row:
                continue
            u = (row[0] or "").strip()
            if u:
                links.append(u)
    return links



def _ensure_row_keys(row: Dict[str, str]) -> Dict[str, str]:
    return {k: (row.get(k) or "").strip() for k in BAZA_MIESZKANIA_HEADERS}


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


def main():
    intake = Path(sys.argv[1]).expanduser() if len(sys.argv) > 1 else DEFAULT_INTAKE
    db_path = Path(sys.argv[2]).expanduser() if len(sys.argv) > 2 else DB_XLSX

    print(f"[INFO] Wejście: {intake}")
    print(f"[INFO] Baza danych: {db_path} (arkusz: {DB_SHEET})")

    links = _read_links_from_csv(intake)
    if not links:
        print("[INFO] Brak linków do przetworzenia.")
        return

    ensure_baza_mieszkania(db_path, sheet=DB_SHEET)

    rows = _process_links_to_rows(links)
    if not rows:
        print("[INFO] Brak poprawnych rekordów do zapisania (np. brak cen).")
        return

    append_mieszkania_rows(db_path, rows, sheet=DB_SHEET)
    print(f"[OK] Zapisano {len(rows)} rekordów do „{db_path.name}” / arkusz „{DB_SHEET}”.")
    print("Gotowe.")


if __name__ == "__main__":
    main()
