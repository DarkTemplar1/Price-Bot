#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Wypisuje (print) łączną liczbę ogłoszeń dla podanego URL wyników Otodom.
Domyślnie: mieszkania na sprzedaż, pomorskie.
"""
from __future__ import annotations
import argparse
import json
import re
import requests

DEFAULT_URL = "https://www.otodom.pl/pl/wyniki/sprzedaz/mieszkanie/pomorskie"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept-Language": "pl-PL,pl;q=0.9,en-US;q=0.8,en;q=0.7",
}

PREFERRED_KEYS = {
    "total", "totalCount", "totalResults",
    "totalElements", "searchTotal", "numberOfResults"
}

def to_int(s: str) -> int | None:
    s = re.sub(r"\D+", "", s or "")
    return int(s) if s else None

def extract_total_from_next_data(html: str) -> int | None:
    m = re.search(
        r'<script id="__NEXT_DATA__" type="application/json">(.*?)</script>',
        html, flags=re.S
    )
    if not m:
        return None
    try:
        data = json.loads(m.group(1))
    except json.JSONDecodeError:
        return None

    best = None
    stack = [data]
    while stack:
        obj = stack.pop()
        if isinstance(obj, dict):
            for k, v in obj.items():
                if isinstance(v, (int, float)) and k in PREFERRED_KEYS:
                    best = max(best or 0, int(v))
                elif isinstance(v, (dict, list)):
                    stack.append(v)
        elif isinstance(obj, list):
            stack.extend(obj)
    return best

def extract_total_from_text(html: str) -> int | None:
    text = re.sub(r"\s+", " ", html)
    # np. „1–36 ogłoszeń z 15663”
    m = re.search(r"ogłoszeń\s+z\s+([\d\s\u00A0]+)", text, flags=re.I)
    if m and (n := to_int(m.group(1))):
        return n
    # zapasowe
    m = re.search(r"Znaleziono\s+([\d\s\u00A0]+)\s+ogłosze", text, flags=re.I)
    if m and (n := to_int(m.group(1))):
        return n
    return None

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--url", "-u", default=DEFAULT_URL, help="URL listingu Otodom")
    args = ap.parse_args()

    r = requests.get(args.url, headers=HEADERS, timeout=20)
    r.raise_for_status()
    html = r.text

    total = extract_total_from_next_data(html) or extract_total_from_text(html)
    if total is None:
        raise SystemExit("Nie udało się znaleźć liczby ogłoszeń.")
    print(total)

if __name__ == "__main__":
    main()
