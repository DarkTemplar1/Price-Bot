#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
linki_mieszkania.py

Zbiera linki do ofert „mieszkania” z Otodom i zapisuje je do CSV
per-województwo: intake_<wojewodztwo>.csv (np. intake_mazowieckie.csv).

- Jeśli plik istnieje, NIE jest nadpisywany – dopisywane są tylko nowe linki.
- Dla każdego dopisanego linku zapisuje datę i godzinę pobrania (Europe/Warsaw).
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import time
import unicodedata
from pathlib import Path
from typing import Iterable
from urllib.parse import urljoin, urlparse, parse_qs, urlencode, urlunparse

import requests
from bs4 import BeautifulSoup

# --- (opcjonalnie) Selenium - nie używamy domyślnie ---
try:
    import undetected_chromedriver as uc  # noqa: F401
    from selenium.webdriver.common.by import By  # noqa: F401
    from selenium.webdriver.support.ui import WebDriverWait  # noqa: F401
    from selenium.webdriver.support import expected_conditions as EC  # noqa: F401
    _SELENIUM_AVAILABLE = True
except Exception:
    _SELENIUM_AVAILABLE = False

from datetime import datetime
try:
    from zoneinfo import ZoneInfo  # Python 3.9+
    _TZ = ZoneInfo("Europe/Warsaw")
except Exception:
    _TZ = None

from secrets import SystemRandom
_RNG = SystemRandom()

BASE_DIR = Path(__file__).resolve().parent

WAIT_SECONDS = 20
HEADLESS = False
RETRIES_NAV = 3
SCROLL_PAUSE = 0.6
USE_SELENIUM = False  # <= domyślnie BEZ przeglądarki (stabilniej/faster)

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

VOIVODESHIPS = {
    "dolnoslaskie": "Dolnośląskie",
    "kujawsko-pomorskie": "Kujawsko-Pomorskie",
    "lubelskie": "Lubelskie",
    "lubuskie": "Lubuskie",
    "lodzkie": "Łódzkie",
    "malopolskie": "Małopolskie",
    "mazowieckie": "Mazowieckie",
    "opolskie": "Opolskie",
    "podkarpackie": "Podkarpackie",
    "podlaskie": "Podlaskie",
    "pomorskie": "Pomorskie",
    "slaskie": "Śląskie",
    "swietokrzyskie": "Świętokrzyskie",
    "warminsko-mazurskie": "Warmińsko-Mazurskie",
    "wielkopolskie": "Wielkopolskie",
    "zachodniopomorskie": "Zachodniopomorskie",
}

def _slugify_region(s: str) -> str:
    s = s.strip().lower()
    # pozwalamy na podanie etykiety z polskimi znakami – mapujemy do slugów
    reverse = {v.lower(): k for k, v in VOIVODESHIPS.items()}
    return VOIVODESHIPS.get(s, None) or reverse.get(s, None) or s  # s może już być slugiem


# ------------------------------ Utils ------------------------------
def add_or_replace_query_param(url: str, key: str, value: str) -> str:
    parts = list(urlparse(url))
    q = parse_qs(parts[4], keep_blank_values=True)
    q[key] = [str(value)]
    parts[4] = urlencode(q, doseq=True)
    return urlunparse(parts)

def _strip_accents(text: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", text) if not unicodedata.combining(c))

def parse_total_offers(html: str) -> int | None:
    for key in ("totalResults", "numberOfItems", "total", "resultsCount", "offersCount"):
        m = re.search(rf'"{key}"\s*:\s*(\d+)', html)
        if m:
            try:
                return int(m.group(1))
            except ValueError:
                continue
    ascii_html = _strip_accents(html.lower())
    m = re.search(r'([\d\s\.,]+)\s*oglosz', ascii_html)
    if m:
        try:
            return int(re.sub(r"[^\d]", "", m.group(1)))
        except ValueError:
            pass
    return None

def parse_links_from_html(html: str, base_url: str) -> set[str]:
    soup = BeautifulSoup(html, "html.parser")
    links: set[str] = set()

    for a in soup.select('a[data-cy="listing-item-link"][href]'):
        href = (a.get("href") or "").strip()
        if not href:
            continue
        full = urljoin(base_url, href.split("?")[0])
        if "/pl/oferta/" in full:
            links.add(full)

    for script in soup.find_all("script", {"type": "application/ld+json"}):
        raw = script.string or ""
        if not raw.strip():
            continue
        try:
            data = json.loads(raw)
        except json.JSONDecodeError:
            continue

        def walk(node):
            if isinstance(node, dict):
                if node.get("@type") == "Product":
                    offers = node.get("offers")
                    if isinstance(offers, dict):
                        inner = offers.get("offers")
                        if isinstance(inner, list):
                            for it in inner:
                                u = (it or {}).get("url")
                                if isinstance(u, str) and "/pl/oferta/" in u:
                                    links.add(urljoin(base_url, u.split("?")[0]))
                for v in node.values():
                    walk(v)
            elif isinstance(node, list):
                for v in node:
                    walk(v)

        walk(data)
    return links

# ------------------------------ FETCHERS ------------------------------
class HttpFetcher:
    def __init__(self, wait_seconds: int = WAIT_SECONDS):
        self.sess = requests.Session()
        self.sess.headers.update(
            {
                "User-Agent": USER_AGENT,
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Accept-Language": "pl-PL,pl;q=0.9,en-US;q=0.8,en;q=0.7",
                "Connection": "keep-alive",
            }
        )
        self.timeout = wait_seconds
        self._html = ""
        self._url = ""

    def get(self, url: str) -> bool:
        last_err = None
        for attempt in range(RETRIES_NAV):
            try:
                r = self.sess.get(url, timeout=self.timeout)
                if r.status_code >= 400:
                    raise requests.HTTPError(f"{r.status_code}")
                self._html = r.text
                self._url = r.url
                return True
            except Exception as e:
                last_err = e
                time.sleep(1.5 * (attempt + 1))
        print(f"[WARN] Nie udało się wczytać: {url} ({last_err})")
        return False

    def page_source(self) -> str:
        return self._html

    def current_url(self) -> str:
        return self._url or ""

    def close(self):
        try:
            self.sess.close()
        except Exception:
            pass

# (SeleniumFetcher pomijam – jak w Twojej wersji, nie zmieniałem)

def _make_fetcher():
    return HttpFetcher(wait_seconds=WAIT_SECONDS)

def collect_links_from_n_pages(search_url: str, pages: int = 1,
                               delay_min: float = 1.5, delay_max: float = 6.0) -> list[str]:
    fetcher = _make_fetcher()
    seen, out = set(), []
    try:
        for page in range(1, pages + 1):
            url = search_url if page == 1 else add_or_replace_query_param(search_url, "page", page)
            if not fetcher.get(url):
                break
            html = fetcher.page_source()
            base_url = fetcher.current_url() or search_url
            links = parse_links_from_html(html, base_url)
            new = [u for u in links if u not in seen]
            if not new and page > 1:
                break
            seen.update(new)
            out.extend(new)
            time.sleep(_RNG.uniform(delay_min, delay_max))
    finally:
        fetcher.close()
    return out

def _read_existing_links(csv_path: Path) -> set[str]:
    existing: set[str] = set()
    if not csv_path.exists() or csv_path.stat().st_size == 0:
        return existing
    with csv_path.open("r", newline="", encoding="utf-8") as f:
        sample = f.read(1024)
        f.seek(0)
        has_header = "link" in (sample.splitlines()[0].lower() if sample else "")
        if has_header:
            r = csv.DictReader(f)
            for row in r:
                u = (row.get("link") or "").strip()
                if u:
                    existing.add(u)
        else:
            rr = csv.reader(f)
            for row in rr:
                if not row:
                    continue
                u = (row[0] or "").strip()
                if u and u.lower() != "link":
                    existing.add(u)
    return existing

def append_links_with_timestamp(links: Iterable[str], csv_path: Path) -> int:
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    existing = _read_existing_links(csv_path)
    new = [u for u in links if u not in existing]
    write_header = not csv_path.exists() or csv_path.stat().st_size == 0
    with csv_path.open("a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if write_header:
            w.writerow(["link", "captured_date", "captured_time"])
        for u in new:
            now = datetime.now(tz=_TZ) if _TZ else datetime.now()
            w.writerow([u, now.date().isoformat(), now.time().isoformat(timespec="seconds")])
    return len(new)

# --------------------------- CLI ---------------------------
if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--region", "-r", required=True,
                    help="Slug województwa (np. mazowieckie) lub etykieta (np. Mazowieckie).")
    ap.add_argument("--pages", "-p", type=int, default=1, help="Ile stron wyników pobrać (domyślnie 1).")
    args = ap.parse_args()

    region_slug = _slugify_region(args.region)
    if region_slug not in VOIVODESHIPS:
        raise SystemExit(f"Nieznane województwo: {args.region}")

    start_url = (
        f"https://www.otodom.pl/pl/wyniki/sprzedaz/mieszkanie/{region_slug}"
        "?limit=36&ownerTypeSingleSelect=ALL&by=DEFAULT&direction=DESC"
    )
    csv_path = BASE_DIR / f"intake_{region_slug}.csv"

    print(f"[INFO] Województwo: {VOIVODESHIPS[region_slug]} ({region_slug})")
    print(f"[INFO] URL startowy: {start_url}")
    print(f"[INFO] Plik intake: {csv_path.name}")

    links = collect_links_from_n_pages(start_url, pages=args.pages)
    added = append_links_with_timestamp(links, csv_path)
    print(f"[OK] Zebrano {len(links)} linków, dopisano {added} nowych do {csv_path.name}")
