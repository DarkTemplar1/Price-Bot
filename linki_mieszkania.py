#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations
import argparse
import csv
import json
import math
import re
import time
import unicodedata
from pathlib import Path
from typing import Iterable, List, Optional
from urllib.parse import urljoin, urlparse, parse_qs, urlencode, urlunparse

import requests
from bs4 import BeautifulSoup
from datetime import datetime
try:
    from zoneinfo import ZoneInfo  # Python 3.9+
    _TZ = ZoneInfo("Europe/Warsaw")
except Exception:
    _TZ = None

from secrets import SystemRandom
_RNG = SystemRandom()

# ------------------------------ WYJŚCIE: Desktop/Pulpit ------------------------------
def _detect_desktop() -> Path:
    home = Path.home()
    for name in ("Desktop", "Pulpit"):
        p = home / name
        if p.exists():
            return p
    return home

BASE_BAZA = _detect_desktop() / "baza danych"
BASE_LINKI = BASE_BAZA / "linki"
BASE_WOJ = BASE_BAZA / "województwa"
BASE_LINKI.mkdir(parents=True, exist_ok=True)
BASE_WOJ.mkdir(parents=True, exist_ok=True)
# -------------------------------------------------------------------------------------

WAIT_SECONDS = 20
RETRIES_NAV = 3
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

# Kolejność ustalona, by raport był przewidywalny
VOIVODESHIPS_ORDERED: List[tuple[str, str]] = [
    ("dolnoslaskie", "Dolnośląskie"),
    ("kujawsko--pomorskie", "Kujawsko-Pomorskie"),
    ("lubelskie", "Lubelskie"),
    ("lubuskie", "Lubuskie"),
    ("lodzkie", "Łódzkie"),
    ("malopolskie", "Małopolskie"),
    ("mazowieckie", "Mazowieckie"),
    ("opolskie", "Opolskie"),
    ("podkarpackie", "Podkarpackie"),
    ("podlaskie", "Podlaskie"),
    ("pomorskie", "Pomorskie"),
    ("slaskie", "Śląskie"),
    ("swietokrzyskie", "Świętokrzyskie"),
    ("warminsko-mazurskie", "Warmińsko-Mazurskie"),
    ("wielkopolskie", "Wielkopolskie"),
    ("zachodniopomorskie", "Zachodniopomorskie"),
]
VOIVODESHIPS = dict(VOIVODESHIPS_ORDERED)
DEFAULT_REGION = "all"  # domyślnie wszystkie

# ------------------------------ Utils ------------------------------
def _slugify_region_input(raw: str) -> List[str]:
    if not raw or raw.strip().lower() == "all":
        return [slug for slug, _ in VOIVODESHIPS_ORDERED]
    label_to_slug = {v.lower(): k for k, v in VOIVODESHIPS.items()}
    out: List[str] = []
    for part in raw.split(","):
        s = part.strip().lower()
        if not s:
            continue
        if s in VOIVODESHIPS:
            out.append(s)
        else:
            slug = label_to_slug.get(s)
            if slug:
                out.append(slug)
            else:
                out.append(s)
    bad = [s for s in out if s not in VOIVODESHIPS]
    if bad:
        valid = ", ".join(VOIVODESHIPS.keys())
        raise SystemExit(
            f"Nieznane województwo/slug: {', '.join(bad)}\n"
            f"Dozwolone slugi: {valid}\n"
            f"Albo użyj: --region all"
        )
    seen = set()
    uniq = []
    for s in out:
        if s not in seen:
            uniq.append(s)
            seen.add(s)
    return uniq

def add_or_replace_query_param(url: str, key: str, value: str) -> str:
    parts = list(urlparse(url))
    q = parse_qs(parts[4], keep_blank_values=True)
    q[key] = [str(value)]
    parts[4] = urlencode(q, doseq=True)
    return urlunparse(parts)

def _strip_accents(text: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", text) if not unicodedata.combining(c))

_ID_RE = re.compile(r"/pl/oferta/[^/]*-(\d{5,})")
def _offer_key(url: str) -> str:
    m = _ID_RE.search(url)
    return m.group(1) if m else url.rstrip("/")

# ------------------------------ Parsowanie ------------------------------
def parse_links_from_html(html: str, base_url: str) -> set[str]:
    soup = BeautifulSoup(html, "html.parser")
    links: set[str] = set()
    for a in soup.select(
        'a[data-cy="listing-item-link"][href], '
        'a[data-testid="listing-item-link"][href], '
        'a[href*="/pl/oferta/"]'
    ):
        href = (a.get("href") or "").strip()
        if not href:
            continue
        full = urljoin(base_url, href.split("?")[0].split("#")[0]).rstrip("/")
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
                                    full2 = urljoin(base_url, u.split("?")[0].split("#")[0]).rstrip("/")
                                    links.add(full2)
                for v in node.values():
                    walk(v)
            elif isinstance(node, list):
                for v in node:
                    walk(v)
        walk(data)
    return links

# ------------------------------ FETCHER ------------------------------
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
                time.sleep(1.25 * (attempt + 1))
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

def _make_fetcher():
    return HttpFetcher(wait_seconds=WAIT_SECONDS)

# ------------------------------ Liczba ogłoszeń → liczba stron ------------------------------
PREFERRED_KEYS = {"total","totalCount","totalResults","totalElements","searchTotal","numberOfResults"}

def _to_int(s: str) -> int | None:
    s = re.sub(r"\D+", "", s or "")
    return int(s) if s else None

def _extract_total_from_next_data(html: str) -> int | None:
    m = re.search(r'<script id="__NEXT_DATA__" type="application/json">(.*?)</script>', html, flags=re.S)
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

def _extract_total_from_text(html: str) -> int | None:
    text = re.sub(r"\s+", " ", html)
    m = re.search(r"ogłoszeń\s+z\s+([\d\s\u00A0]+)", text, flags=re.I)
    if m and (n := _to_int(m.group(1))):
        return n
    m = re.search(r"Znaleziono\s+([\d\s\u00A0]+)\s+ogłosze", text, flags=re.I)
    if m and (n := _to_int(m.group(1))):
        return n
    return None

def _fetch_total_results(url: str, user_agent: str) -> int:
    headers = {"User-Agent": user_agent, "Accept-Language": "pl-PL,pl;q=0.9,en-US;q=0.8,en;q=0.7"}
    r = requests.get(url, headers=headers, timeout=20)
    r.raise_for_status()
    html = r.text
    total = _extract_total_from_next_data(html) or _extract_total_from_text(html)
    if total is None:
        raise RuntimeError("Nie udało się znaleźć liczby ogłoszeń.")
    return total

def _limit_from_url(url: str, default: int = 36) -> int:
    q = parse_qs(urlparse(url).query)
    try:
        return int((q.get("limit") or [default])[0])
    except Exception:
        return default

def _pages_from_total(url: str, user_agent: str) -> tuple[int, int, int]:
    total = _fetch_total_results(url, user_agent)
    limit = _limit_from_url(url, default=36)  # zwykle 36 lub 72
    pages = math.ceil(total / max(limit, 1))
    return pages, total, limit

# ------------------------------ Core ------------------------------
def collect_links_all_pages(
    search_url: str,
    min_delay: float = 1.5,
    max_delay: float = 5.0,
    stop_after_k_empty_pages: int = 25,
    max_pages: Optional[int] = None,
) -> list[str]:
    fetcher = _make_fetcher()
    seen_keys, out = set(), []
    page = 1
    empty_streak = 0
    try:
        while True:
            if isinstance(max_pages, int) and page > max_pages:
                print(f"[INFO] Osiągnięto limit {max_pages} stron – koniec.")
                break

            url = search_url if page == 1 else add_or_replace_query_param(search_url, "page", page)
            ok = fetcher.get(url)
            if not ok:
                if page == 1:
                    print(f"[ERR] Nie udało się pobrać pierwszej strony: {url}")
                break

            html = fetcher.page_source()
            base_url = fetcher.current_url() or search_url
            links = parse_links_from_html(html, base_url)

            if not links:
                print(f"[INFO] Strona {page}: brak jakichkolwiek linków – koniec.")
                break

            new_links = []
            for u in links:
                k = _offer_key(u)
                if k not in seen_keys:
                    seen_keys.add(k)
                    new_links.append(u)

            if new_links:
                out.extend(new_links)
                empty_streak = 0
                print(f"[INFO] Strona {page}: {len(new_links)} nowych (łącznie {len(out)}).")
            else:
                empty_streak += 1
                print(f"[INFO] Strona {page}: 0 nowych (seria {empty_streak}/{stop_after_k_empty_pages}).")
                if empty_streak >= stop_after_k_empty_pages:
                    print("[INFO] Zbyt wiele stron bez nowych ofert – przerywam.")
                    break

            page += 1
            time.sleep(_RNG.uniform(min_delay, max_delay))
    finally:
        fetcher.close()
    return out

# ------------------------------ CSV ------------------------------
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
    existing_links = _read_existing_links(csv_path)
    existing_keys = {_offer_key(u) for u in existing_links}
    to_write = []
    for u in links:
        key = _offer_key(u)
        if (u not in existing_links) and (key not in existing_keys):
            to_write.append(u)
            existing_links.add(u)
            existing_keys.add(key)
    write_header = not csv_path.exists() or csv_path.stat().st_size == 0
    with csv_path.open("a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if write_header:
            w.writerow(["link", "captured_date", "captured_time"])
        for u in to_write:
            now = datetime.now(tz=_TZ) if _TZ else datetime.now()
            w.writerow([u, now.date().isoformat(), now.time().isoformat(timespec="seconds")])
    return len(to_write)

# ------------------------------ Runner ------------------------------
def run_for_region(slug: str) -> tuple[int, int]:
    start_url = (
        f"https://www.otodom.pl/pl/wyniki/sprzedaz/mieszkanie/{slug}"
        "?limit=72&ownerTypeSingleSelect=ALL&by=DEFAULT&direction=DESC"
    )
    csv_path = BASE_LINKI / f"intake_{slug}.csv"   # <<< ZAPIS NA PULPICIE / baza danych / linki

    print(f"\n========== {VOIVODESHIPS[slug]} ({slug}) ==========")
    print(f"[INFO] URL startowy: {start_url}")
    print(f"[INFO] Plik intake:  {csv_path}")

    try:
        pages, total, limit = _pages_from_total(start_url, USER_AGENT)
        print(f"[INFO] Ogłoszeń łącznie: {total} | limit/stronę: {limit} | stron do sprawdzenia: {pages}")
        max_pages = pages
    except Exception as e:
        print(f"[WARN] Nie udało się ustalić liczby stron ({e}). Lecę bez twardego limitu.")
        max_pages = None

    links = collect_links_all_pages(
        start_url,
        min_delay=1.5,
        max_delay=5.0,
        stop_after_k_empty_pages=25,
        max_pages=max_pages,
    )
    added = append_links_with_timestamp(links, csv_path)
    print(f"[OK] Zebrano {len(links)} linków, dopisano {added} nowych do {csv_path.name}")
    return (len(links), added)

# --------------------------- CLI ---------------------------
if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--region", "-r",
        default=DEFAULT_REGION,
        help=("Slug/etykieta województwa (np. 'mazowieckie' lub 'Mazowieckie'), "
              "lista po przecinku (np. 'mazowieckie,slaskie') lub 'all' (domyślnie).")
    )
    args = ap.parse_args()

    region_slugs = _slugify_region_input(args.region)
    total_found = 0
    total_added = 0
    for slug in region_slugs:
        found, added = run_for_region(slug)
        total_found += found
        total_added += added
        time.sleep(_RNG.uniform(1.0, 2.5))

    print("\n=========== PODSUMOWANIE ===========")
    print(f"Łącznie znalezionych linków: {total_found}")
    print(f"Łącznie NOWO dopisanych:     {total_added}")
