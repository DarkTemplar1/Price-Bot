#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import csv
import os
import time
from typing import Dict, List
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin

HEADERS_POOL = [
    # kilka UA na rotację, żeby ograniczyć 403
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:124.0) Gecko/20100101 Firefox/124.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",
]

BASE = "https://www.otodom.pl"

def read_links(csv_path: str) -> List[str]:
    links = []
    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        if reader.fieldnames and "link" in [c.strip().lower() for c in reader.fieldnames]:
            col_map = {c.strip().lower(): c for c in reader.fieldnames}
            for row in reader:
                href = (row.get(col_map.get("link")) or "").strip()
                if href:
                    links.append(href)
        else:
            f.seek(0)
            raw = csv.reader(f)
            for i, row in enumerate(raw):
                if i == 0 and row and row[0].lower().strip() in ("link", "links"):
                    continue
                if not row:
                    continue
                href = (row[0] or "").strip()
                if href:
                    links.append(href)

    norm = []
    seen = set()
    for href in links:
        href2 = href.strip()
        if href2.startswith("/hpr"):
            href2 = href2[4:]
        if href2.startswith("/"):
            href2 = urljoin(BASE, href2)
        if href2.startswith("https://otodom.pl"):
            href2 = href2.replace("https://otodom.pl", "https://www.otodom.pl", 1)
        if href2 and href2 not in seen:
            seen.add(href2)
            norm.append(href2)
    return norm

def pick_headers(i: int) -> Dict[str, str]:
    ua = HEADERS_POOL[i % len(HEADERS_POOL)]
    return {
        "User-Agent": ua,
        "Accept-Language": "pl-PL,pl;q=0.9",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Referer": "https://www.otodom.pl/",
        "Cache-Control": "no-cache",
        "Pragma": "no-cache",
    }

def fetch(url: str, attempt: int = 0) -> requests.Response:
    time.sleep(0.4 if attempt == 0 else 1.2)
    return requests.get(url, headers=pick_headers(attempt), timeout=20, allow_redirects=True)

def extract_text(el) -> str:
    if not el:
        return ""
    return " ".join(el.stripped_strings)

def parse_details(soup: BeautifulSoup) -> Dict[str, str]:
    details: Dict[str, str] = {}
    grids = soup.select('[data-sentry-element="ItemGridContainer"]')
    for g in grids:
        cells = [c for c in g.find_all(recursive=False) if c.name == "div"]
        if len(cells) < 2:
            continue
        label = extract_text(cells[0]).replace(":", "").strip()
        value = extract_text(cells[1])
        if not label or not value:
            continue
        details[label] = value
    return details

def normalize_floor(val: str) -> str:
    """Normalizuje wartość pola 'Piętro'."""
    if not val:
        return ""
    val = val.strip()
    if "/" in val:
        part = val.split("/")[0].strip()
        if part.lower() == "parter":
            return "parter"
        return part
    if val.lower() == "parter":
        return "parter"
    return val

def parse_offer(url: str) -> Dict[str, str]:
    resp = None
    for attempt in range(3):
        try:
            resp = fetch(url, attempt)
            if resp.status_code == 200:
                break
            if resp.status_code in (403, 429, 500, 502, 503):
                continue
            break
        except requests.RequestException:
            continue

    if not resp or resp.status_code != 200:
        return {}

    soup = BeautifulSoup(resp.text, "html.parser")

    cena = extract_text(soup.select_one('strong[data-cy="adPageHeaderPrice"]'))
    cena_m2 = extract_text(soup.select_one('[aria-label="Cena za metr kwadratowy"]'))

    det = parse_details(soup)

    mapping = {
        "Powierzchnia": "metry",
        "Liczba pokoi": "liczba_pokoi",
        "Piętro": "pietro",
        "Rynek": "rynek",
        "Rok budowy": "rok_budowy",
        "Materiał budynku": "material",
    }

    row = {
        "cena": cena or "",
        "cena_za_metr": cena_m2 or "",
        "metry": "",
        "liczba_pokoi": "",
        "pietro": "",
        "rynek": "",
        "rok_budowy": "",
        "material": "",
        "link": url,
    }

    for label, col in mapping.items():
        if label in det:
            val = det[label]
            if col == "pietro":
                val = normalize_floor(val)
            row[col] = val

    return row

def save_rows(rows: List[Dict[str, str]], out_csv: str):
    os.makedirs(os.path.dirname(out_csv), exist_ok=True)
    cols = ["cena","cena_za_metr","metry","liczba_pokoi","pietro","rynek","rok_budowy","material","link"]
    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in cols})

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--region", required=True, help="Region tylko do logów (np. Kujawsko-Pomorskie)")
    parser.add_argument("--input", required=True, help="CSV z kolumną 'link' lub pojedynczą kolumną z linkami")
    parser.add_argument("--output", required=True, help="Dokąd zapisać wynik CSV")
    args = parser.parse_args()

    links = read_links(args.input)
    print(f"[INFO] Wczytano {len(links)} linków do przetworzenia")

    rows: List[Dict[str, str]] = []
    for i, url in enumerate(links, 1):
        data = parse_offer(url)
        if not data:
            print(f"[SCRAPER] ⚠️ Nie udało się pobrać/parsować: {url}")
            continue

        if not data.get("cena"):
            print(f"[SCRAPER] ⚠️ Brak ceny, pomijam {url}")
            continue

        rows.append(data)
        if i % 10 == 0:
            print(f"[INFO] Przetworzono {i}/{len(links)}")

    if rows:
        save_rows(rows, args.output)
        print(f"[OK] Zapisano dane do {args.output}")
    else:
        print("[INFO] Brak wyników do zapisania")

if __name__ == "__main__":
    main()
