#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
linki_mieszkania.py

Zbiera linki do ofert „mieszkania” z Otodom i zapisuje je do CSV.
Dla KAŻDEGO dopisanego linku zapisuje także datę i godzinę pobrania
w osobnych kolumnach:
  - captured_date (YYYY-MM-DD, strefa Europe/Warsaw)
  - captured_time (HH:MM:SS, strefa Europe/Warsaw)

Plik wyjściowy: intake.csv (zgodny wstecz – scraper czyta tylko kolumnę `link`).

Uwaga:
- Skrypt próbuje działać bez przeglądarki (requests + BeautifulSoup).
- Jeśli Otodom zmieni layout, w razie potrzeby można przełączyć się na Selenium.
"""

from __future__ import annotations

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

# ------------------------------------------------------------------
from datetime import datetime
try:
    from zoneinfo import ZoneInfo  # Python 3.9+
    _TZ = ZoneInfo("Europe/Warsaw")
except Exception:
    _TZ = None  # fallback: naive localtime

from secrets import SystemRandom
_RNG = SystemRandom()

BASE_DIR = Path(__file__).resolve().parent
CSV_PATH = BASE_DIR / "intake.csv"  # zachowujemy zgodność ze scraperem

START_URL = (
    "https://www.otodom.pl/pl/wyniki/sprzedaz/mieszkanie/cala-polska"
    "?limit=36&ownerTypeSingleSelect=ALL&by=DEFAULT&direction=DESC"
)

WAIT_SECONDS = 20
HEADLESS = False
RETRIES_NAV = 3
SCROLL_PAUSE = 0.6
USE_SELENIUM = False  # <= domyślnie BEZ przeglądarki (stabilniej/faster)

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

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
    # spróbuj z JSON wstrzykniętego do strony
    for key in ("totalResults", "numberOfItems", "total", "resultsCount", "offersCount"):
        m = re.search(rf'"{key}"\s*:\s*(\d+)', html)
        if m:
            try:
                return int(m.group(1))
            except ValueError:
                continue
    # fallback: heurystyka po tekście
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

    # główny selektor kart ogłoszeń
    for a in soup.select('a[data-cy="listing-item-link"][href]'):
        href = (a.get("href") or "").strip()
        if not href:
            continue
        full = urljoin(base_url, href.split("?")[0])
        if "/pl/oferta/" in full:
            links.add(full)

    # JSON-LD (czasem lista ofert jest również w danych strukturalnych)
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
    """Prosty pobieracz przez requests (bez przeglądarki)."""

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


class SeleniumFetcher:
    """Opcjonalny fetcher przez undetected_chromedriver (użyj tylko gdy to konieczne)."""

    def __init__(self, headless: bool = HEADLESS, wait_seconds: int = WAIT_SECONDS):
        if not _SELENIUM_AVAILABLE:
            raise RuntimeError("Selenium/undetected_chromedriver nie są dostępne w środowisku.")
        opts = uc.ChromeOptions()
        if headless:
            opts.add_argument("--headless=new")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--disable-dev-shm-usage")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--start-maximized")
        if USER_AGENT:
            opts.add_argument(f"--user-agent={USER_AGENT}")
        self.driver = uc.Chrome(options=opts)
        self.wait = WebDriverWait(self.driver, wait_seconds)

    def get(self, url: str) -> bool:
        last_err = None
        for attempt in range(RETRIES_NAV):
            try:
                self.driver.get(url)
                try:
                    self.wait.until(
                        EC.any_of(
                            EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'a[data-cy="listing-item-link"]')),
                            EC.presence_of_element_located((By.CSS_SELECTOR, 'script[type="application/ld+json"]')),
                        )
                    )
                except Exception:
                    time.sleep(2)
                self._gentle_scroll()
                return True
            except Exception as e:
                last_err = e
                time.sleep(1.5 * (attempt + 1))
        print(f"[WARN] Nie udało się wczytać: {url} ({last_err})")
        return False

    def page_source(self) -> str:
        return self.driver.page_source

    def current_url(self) -> str:
        return self.driver.current_url

    def _gentle_scroll(self):
        try:
            h = self.driver.execute_script("return document.body.scrollHeight")
            pos = 0
            while pos < h:
                pos += int(h * 0.25)
                self.driver.execute_script(f"window.scrollTo(0, {pos});")
                time.sleep(SCROLL_PAUSE)
                h = self.driver.execute_script("return document.body.scrollHeight")
        except Exception:
            pass

    def close(self):
        try:
            self.driver.quit()
        except Exception:
            pass


# --------------------------- LOGIKA ---------------------------
def _make_fetcher():
    if USE_SELENIUM:
        try:
            return SeleniumFetcher(headless=HEADLESS, wait_seconds=WAIT_SECONDS)
        except Exception as e:
            print(f"[INFO] Przełączam na tryb HTTP (Selenium nie wystartował): {e}")
    return HttpFetcher(wait_seconds=WAIT_SECONDS)


def collect_links_from_n_pages(
    search_url: str,
    pages: int = 1,
    delay_min: float = 1.5,
    delay_max: float = 6.0,
) -> list[str]:
    """
    Zbiera linki z podanej liczby stron wyników. Między kolejnymi stronami
    wprowadza losowe opóźnienie z zakresu [delay_min, delay_max].
    """
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
            # losowe opóźnienie 1.5–6.0s
            time.sleep(_RNG.uniform(delay_min, delay_max))
    finally:
        fetcher.close()
    return out


def _read_existing_links(csv_path: Path) -> set[str]:
    """Czytaj istniejące linki z CSV (obsługuje nagłówki: 'link', 'link,captured_at', 'link,captured_date,captured_time')."""
    existing: set[str] = set()
    if not csv_path.exists() or csv_path.stat().st_size == 0:
        return existing

    with csv_path.open("r", newline="", encoding="utf-8") as f:
        # spróbuj DictReader
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
            # fallback: zwykły reader, pierwsza kolumna = link
            rr = csv.reader(f)
            for row in rr:
                if not row:
                    continue
                u = (row[0] or "").strip()
                if u and u.lower() != "link":
                    existing.add(u)
    return existing


def append_links_with_timestamp(links: Iterable[str], csv_path: Path = CSV_PATH) -> int:
    """
    Dopisz unikalne linki do CSV z kolumnami: link,captured_date,captured_time.
    Zwraca liczbę NOWO dodanych rekordów.
    """
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    existing = _read_existing_links(csv_path)
    new = [u for u in links if u not in existing]

    # ustal, czy trzeba dopisać nagłówki
    write_header = not csv_path.exists() or csv_path.stat().st_size == 0

    with csv_path.open("a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if write_header:
            w.writerow(["link", "captured_date", "captured_time"])
        for u in new:
            now = datetime.now(tz=_TZ) if _TZ else datetime.now()
            date_str = now.date().isoformat()                        # YYYY-MM-DD
            time_str = now.time().isoformat(timespec="seconds")      # HH:MM:SS
            w.writerow([u, date_str, time_str])

    return len(new)


# --------------------------- CLI ---------------------------
if __name__ == "__main__":
    pages = 1  # Możesz zwiększyć, np. 3–5
    print(f"[INFO] Pobieram linki z {pages} strony/stron…")
    all_links = collect_links_from_n_pages(START_URL, pages=pages)
    added = append_links_with_timestamp(all_links, CSV_PATH)
    print(f"[OK] Zebrano {len(all_links)} linków, dopisano {added} nowych do {CSV_PATH.name}")
