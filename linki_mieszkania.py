#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import csv
import json
import math
import os
import re
from typing import Any, Optional

import requests
from bs4 import BeautifulSoup

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept-Language": "pl-PL,pl;q=0.9,en-US;q=0.8,en;q=0.7",
}

# ------------------ Pomocnicze: pobieranie i parsowanie JSON z Next.js ------------------ #

def fetch(url: str) -> str:
    r = requests.get(url, headers=HEADERS, timeout=20)
    r.raise_for_status()
    return r.text

def extract_next_data(html: str) -> Optional[dict]:
    """Wyciąga JSON z <script id="__NEXT_DATA__">."""
    soup = BeautifulSoup(html, "html.parser")
    tag = soup.find("script", id="__NEXT_DATA__")
    if not tag or not tag.string:
        return None
    try:
        return json.loads(tag.string)
    except json.JSONDecodeError:
        return None

def fetch_next_data_json(build_id: str, path_and_query: str) -> Optional[dict]:
    """
    Dla aplikacji Next.js pobiera bezpośredni JSON:
    https://host/_next/data/{buildId}/{sciezka}.json{?query}
    """
    if path_and_query.startswith("/"):
        path_and_query = path_and_query[1:]
    base = "https://www.otodom.pl/_next/data"
    url_json = f"{base}/{build_id}/{path_and_query}.json"
    try:
        r = requests.get(url_json, headers=HEADERS, timeout=20)
        r.raise_for_status()
        return r.json()
    except Exception:
        return None

def deep_find_total(d: Any) -> Optional[int]:
    """
    Rekurencyjnie szuka licznika w polach typu:
    totalCount, total, totalResults, resultsCount, count, itemsCount, hits, etc.
    Zwraca największą sensowną wartość.
    """
    KEY_HINTS = {"totalcount", "total", "totalresults", "results", "resultscount",
                 "count", "items", "hits", "offerscount", "listingscount"}
    best = None

    def walk(x: Any):
        nonlocal best
        if isinstance(x, dict):
            for k, v in x.items():
                lk = k.lower().replace("_", "").replace("-", "")
                if lk in KEY_HINTS or any(h in lk for h in KEY_HINTS):
                    if isinstance(v, (int, float)) and v >= 0:
                        val = int(v)
                        if best is None or val > best:
                            best = val
                walk(v)
        elif isinstance(x, list):
            for el in x:
                walk(el)

    walk(d)
    return best

def get_total_offers(search_url: str) -> Optional[int]:
    """
    Zwraca całkowitą liczbę ogłoszeń dla danego URL-a wyników wyszukiwania.
    """
    try:
        html = fetch(search_url)
    except Exception:
        return None

    total = None
    next_data = extract_next_data(html)

    if next_data:
        total = deep_find_total(next_data)
        if total is None:
            build_id = next_data.get("buildId")
            if build_id:
                path_and_query = re.sub(r"^https?://[^/]+/", "", search_url)
                json_payload = fetch_next_data_json(build_id, path_and_query)
                if json_payload:
                    total = deep_find_total(json_payload)

    if total is None:
        # Ostatnia próba – regex po HTML, np. "Zobacz 1234 ogłoszeń"
        m = re.search(r"(\d[\d\s]{2,})\s*(?:ogłoszeń|ogloszen|ofert|wyników|wynikow)", html, flags=re.I)
        if m:
            total = int(m.group(1).replace(" ", ""))

    return total

# ------------------ Logika zbierania linków ------------------ #

POLISH_MAP = str.maketrans({
    "ą": "a", "ć": "c", "ę": "e", "ł": "l", "ń": "n", "ó": "o", "ś": "s", "ż": "z", "ź": "z",
    "Ą": "A", "Ć": "C", "Ę": "E", "Ł": "L", "Ń": "N", "Ó": "O", "Ś": "S", "Ż": "Z", "Ź": "Z",
})

def region_to_url(region: str) -> str:
    """
    Konwertuje nazwę województwa na format otodom.pl:
    - zamienia polskie znaki (ł → l, ś → s, itd.)
    - spacje → '-'
    - każde pojedyncze '-' → '--' (np. Kujawsko--Pomorskie)
    """
    region_ascii = region.translate(POLISH_MAP)
    region_ascii = region_ascii.replace(" ", "-")
    region_ascii = region_ascii.replace("-", "--")
    return region_ascii.lower()

def make_search_url(region: str, page: int) -> str:
    base = f"https://www.otodom.pl/pl/wyniki/sprzedaz/mieszkanie/{region_to_url(region)}"
    # limit=72 – tyle wyników na stronę
    return f"{base}?limit=72&ownerTypeSingleSelect=ALL&by=DEFAULT&direction=DESC&page={page}"

def pobierz_linki(region: str, output_file: str):
    # 1) Ustal łączną liczbę ogłoszeń i liczbę stron
    first_page_url = make_search_url(region, page=1)
    print(f"[DEBUG] URL 1. strony: {first_page_url}")
    total = get_total_offers(first_page_url)

    if total is None:
        print("[WARN] Nie udało się odczytać liczby ogłoszeń z JSON/HTML. "
              "Przejdę tylko przez pierwszą stronę.")
        pages_to_check = 1
    else:
        pages_to_check = max(1, math.ceil(total / 72))
        print(f"[INFO] Łączna liczba ogłoszeń: {total}. Sprawdzę {pages_to_check} stron(y).")

    collected_links = []
    total_collected = 0

    # 2) Iteruj dokładnie przez wyliczoną liczbę stron
    for page in range(1, pages_to_check + 1):
        url = make_search_url(region, page)
        try:
            r = requests.get(url, headers=HEADERS, timeout=20)
        except Exception as e:
            print(f"[WARN] Błąd pobierania {url}: {e}")
            continue

        if r.status_code != 200:
            print(f"[WARN] Nie udało się pobrać {url}, kod {r.status_code}")
            continue

        soup = BeautifulSoup(r.text, "html.parser")
        offers = soup.select("a[data-cy='listing-item-link']")
        new_links = [a["href"] for a in offers if a.get("href")]

        total_collected += len(new_links)
        print(f"[INFO] Strona {page}: {len(new_links)} linków (łącznie {total_collected}).")
        collected_links.extend(new_links)

    # 3) Zapis do CSV
    outdir = os.path.dirname(output_file)
    if outdir:
        os.makedirs(outdir, exist_ok=True)

    with open(output_file, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["link"])
        for link in collected_links:
            if link.startswith("/"):
                link = "https://www.otodom.pl" + link
            writer.writerow([link])

    print(f"[OK] Zapisano {len(collected_links)} linków do {output_file}")

# ------------------ Uruchomienie ------------------ #

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--region", required=True, help="Nazwa województwa, np. 'Małopolskie' albo 'Kujawsko-Pomorskie'")
    parser.add_argument("--output", required=True, help="Ścieżka do pliku CSV z linkami")
    args = parser.parse_args()

    pobierz_linki(args.region, args.output)
