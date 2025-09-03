#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import csv
import time
import random
from http.client import RemoteDisconnected
from pathlib import Path

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from bazadanych import pobierz_dane_z_otodom  # <<< IMPORT FUNKCJI SCRAPERA

print("[SCRAPER] Uruchomiono scraper_otodom.py")

# ------------------------------ ŚCIEŻKI ------------------------------
def _detect_desktop() -> Path:
    home = Path.home()
    for name in ("Desktop", "Pulpit"):
        p = home / name
        if p.exists():
            return p
    return home

BASE_BAZA = _detect_desktop() / "baza danych"
BASE_LINKI_DIR = BASE_BAZA / "linki"
BASE_WOJ_DIR = BASE_BAZA / "województwa"
BASE_LINKI_DIR.mkdir(parents=True, exist_ok=True)
BASE_WOJ_DIR.mkdir(parents=True, exist_ok=True)

# ------------------------------ REGION UTILS ------------------------------
VOIVODESHIPS_LABEL_SLUG: list[tuple[str, str]] = [
    ("Dolnośląskie", "dolnoslaskie"),
    ("Kujawsko-Pomorskie", "kujawsko-pomorskie"),
    ("Lubelskie", "lubelskie"),
    ("Lubuskie", "lubuskie"),
    ("Łódzkie", "lodzkie"),
    ("Małopolskie", "malopolskie"),
    ("Mazowieckie", "mazowieckie"),
    ("Opolskie", "opolskie"),
    ("Podkarpackie", "podkarpackie"),
    ("Podlaskie", "podlaskie"),
    ("Pomorskie", "pomorskie"),
    ("Śląskie", "slaskie"),
    ("Świętokrzyskie", "swietokrzyskie"),
    ("Warmińsko-Mazurskie", "warminsko-mazurskie"),
    ("Wielkopolskie", "wielkopolskie"),
    ("Zachodniopomorskie", "zachodniopomorskie"),
]
SLUG_TO_LABEL = {slug: label for label, slug in VOIVODESHIPS_LABEL_SLUG}
LABEL_TO_SLUG = {label.lower(): slug for label, slug in VOIVODESHIPS_LABEL_SLUG}


def normalize_region_single(raw: str) -> tuple[str, str]:
    s = (raw or "").strip()
    low = s.lower()
    if low in SLUG_TO_LABEL:
        return SLUG_TO_LABEL[low], low
    if low in LABEL_TO_SLUG:
        slug = LABEL_TO_SLUG[low]
        return SLUG_TO_LABEL[slug], slug
    raise SystemExit(f"Nieznane województwo: {raw}")


# ------------------------------ KONFIG ------------------------------
WAIT_SECONDS = 20
KOLUMNY_WYJ = [
    "cena", "cena_za_metr", "metry", "liczba_pokoi", "pietro", "rynek", "rok_budowy",
    "material", "wojewodztwo", "powiat", "gmina", "miejscowosc", "dzielnica", "ulica", "link",
]


# ------------------------------ UTILS ------------------------------
def _is_remote_closed(exc: Exception) -> bool:
    s = (str(exc) or "").lower()
    if isinstance(exc, RemoteDisconnected):
        return True
    return any(m in s for m in (
        "remote end closed connection without response", "remotedisconnected",
        "connection reset", "chunked encoding"
    ))


def _pobierz_soup(url: str):
    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    driver = webdriver.Chrome(options=opts)
    try:
        driver.get(url)
        WebDriverWait(driver, WAIT_SECONDS).until(
            EC.any_of(
                EC.presence_of_element_located((By.CSS_SELECTOR, "meta[name='description']")),
                EC.presence_of_element_located((By.CSS_SELECTOR, "strong[data-testid='ad-price']")),
                EC.presence_of_element_located((By.CSS_SELECTOR, "link[rel='canonical']")),
            )
        )
        html = driver.page_source
    finally:
        driver.quit()
    return BeautifulSoup(html, "html.parser")


# ------------------------------ CSV I/O ------------------------------
def _links_paths_candidates(label: str, slug: str) -> list[Path]:
    return [
        BASE_LINKI_DIR / f"{label}.csv",
        BASE_LINKI_DIR / f"{slug}.csv",
        BASE_LINKI_DIR / f"intake_{label}.csv",
        BASE_LINKI_DIR / f"intake_{slug}.csv",
    ]


def _find_links_csv(label: str, slug: str) -> Path | None:
    for p in _links_paths_candidates(label, slug):
        if p.exists() and p.stat().st_size > 0:
            return p
    return None


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


# ------------------------------ PIPE ------------------------------
def przetworz_linki_do_csv(input_csv: Path, out_csv: Path, limit: int = 0):
    try:
        links = _read_links_from_csv(input_csv)
    except Exception as e:
        raise SystemExit(f"Nie mogę odczytać linków z {input_csv}: {e}")

    if limit and limit > 0:
        links = links[:limit]

    print(f"[INFO] Do przetworzenia linków: {len(links)}")
    if not links:
        print("[INFO] Brak linków — nic do zrobienia.")
        return

    dane_lista: list[dict] = []
    for i, link in enumerate(links, start=1):
        print(f"\n🌐 [{i}/{len(links)}] {link}")
        try:
            dane = pobierz_dane_z_otodom(link)
        except Exception as e:
            print(f"❌ Błąd przy linku: {e}")
            continue

        if not dane:
            continue

        for k in KOLUMNY_WYJ:
            dane.setdefault(k, "")

        if (dane.get("cena") or "").strip():
            dane_lista.append(dane)

        # >>> losowe opóźnienie 4–8s
        delay = random.uniform(4.0, 8.0)
        print(f"[INFO] Czekam {delay:.2f}s przed kolejnym ogłoszeniem…")
        time.sleep(delay)

    _append_rows_to_output_csv(out_csv, dane_lista)
    print(f"\n✅ Zapisano {len(dane_lista)} rekordów do {out_csv}")


# ------------------------------ CLI ------------------------------
if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--region", "-r", help="Etykieta lub slug")
    ap.add_argument("--input", "-i", type=str, help="Ścieżka do pliku CSV z linkami (zamiast --region)")
    ap.add_argument("--limit", "-l", type=int, default=0, help="Ilość ogłoszeń do przetworzenia (0 = wszystkie)")
    args = ap.parse_args()

    if args.input:
        input_csv = Path(args.input)
        if not input_csv.exists():
            raise SystemExit(f"❌ Podany plik nie istnieje: {input_csv}")
        label = input_csv.stem
    else:
        if not args.region:
            raise SystemExit("❌ Musisz podać --region lub --input")
        label, slug = normalize_region_single(args.region)
        input_csv = _find_links_csv(label, slug)
        if not input_csv:
            raise SystemExit("Nie znaleziono pliku z linkami.")

    out_csv = _ensure_results_csv_label(label)
    przetworz_linki_do_csv(input_csv, out_csv, limit=args.limit)
