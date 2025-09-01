#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import csv
import json
import re
import time
from http.client import RemoteDisconnected
from pathlib import Path

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

print("[SCRAPER] Uruchomiono scraper_otodom_mieszkania.py")

# ------------------------------ ŚCIEŻKI ------------------------------
def _detect_desktop() -> Path:
    home = Path.home()
    for name in ("Desktop","Pulpit"):
        p = home / name
        if p.exists():
            return p
    return home

BASE_BAZA = _detect_desktop() / "baza danych"
BASE_LINKI_DIR = BASE_BAZA / "linki"          # CSV z linkami (ETYKIETY)
BASE_WOJ_DIR   = BASE_BAZA / "województwa"    # CSV z wynikami (ETYKIETY)
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
from adres_otodom import (
    set_contact_email,
    parsuj_adres_string,
    uzupelnij_braki_z_nominatim,
    dopelnij_powiat_gmina_jesli_brak,
    _clean_gmina,
    _consistency_pass_row,
)

CONTACT_EMAIL = "twoj_email@domena.pl"
WAIT_SECONDS = 20
set_contact_email(CONTACT_EMAIL)

KOLUMNY_WYJ = [
    "cena","cena_za_metr","metry","liczba_pokoi","pietro","rynek","rok_budowy","material",
    "wojewodztwo","powiat","gmina","miejscowosc","dzielnica","ulica","link",
]

szukane_pola = {
    "Cena": "cena",
    "Cena za m²": "cena_za_metr",
    "Powierzchnia": "metry",
    "Liczba pokoi": "liczba_pokoi",
    "Piętro": "pietro",
    "Rynek": "rynek",
    "Rok budowy": "rok_budowy",
    "Materiał budynku": "material",
}

LABEL_SYNONYMS = {
    "Liczba pokoi": ("liczba_pokoi", ["Liczba pokoi"]),
    "Rynek": ("rynek", ["Rynek", "Typ rynku"]),
    "Rok budowy": ("rok_budowy", ["Rok budowy", "Rok bud."]),
    "Materiał budynku": ("material", ["Materiał budynku", "Materiał", "Materiał bud."]),
    "Piętro": ("pietro", ["Piętro", "Kondygnacja", "Poziom"]),
}

def format_pln_int(x) -> str | None:
    if x is None: return None
    try: return f"{int(x):,} zł".replace(",", " ")
    except Exception: return None

def fmt_metry(value) -> str | None:
    if value is None: return None
    try:
        v = float(str(value).replace(",", "."))
        s = f"{v:.2f}".rstrip("0").rstrip(".").replace(".", ",")
        return f"{s} m²"
    except Exception:
        return None

def _num(s: str):
    if not s: return None
    digits = re.sub(r"[^\d]", "", s)
    return int(digits) if digits else None

def oblicz_cena_za_metr(cena_str: str, metry_str: str) -> str:
    try:
        cena = int(re.sub(r"[^\d]", "", cena_str or ""))
        metry = float((metry_str or "").replace(",", ".").replace("m²", "").strip())
        if not cena or not metry: return ""
        return f"{round(cena / metry)} zł/m²"
    except Exception:
        return ""

def _norm_txt(s: str | None) -> str:
    return re.sub(r"\s+", " ", (s or "").replace("\xa0", " ")).strip()

_REMOTE_CLOSED_MARKERS = (
    "remote end closed connection without response",
    "remotedisconnected",
    "connection reset",
    "chunked encoding",
)
def _is_remote_closed(exc: Exception) -> bool:
    s = (str(exc) or "").lower()
    if isinstance(exc, RemoteDisconnected): return True
    return any(m in s for m in _REMOTE_CLOSED_MARKERS)

def _safe_enrich_address(adr: dict, tries: int = 5, base_sleep: float = 1.2) -> dict:
    cur = dict(adr or {})
    for i in range(tries):
        try:
            cur = uzupelnij_braki_z_nominatim(cur)
            cur = dopelnij_powiat_gmina_jesli_brak(cur)
            return cur
        except Exception as e:
            if _is_remote_closed(e):
                time.sleep(base_sleep * (i + 1))
                continue
            raise
    print("[WARN] Nominatim: zrezygnowano po wielokrotnych próbach")
    return cur

PRICE_SELECTORS = [
    "strong[data-testid='ad-price']",
    "[data-testid='ad-price']",
    "[data-cy='ad-price']",
]
def _extract_price_from_dom(soup, wynik: dict):
    for sel in PRICE_SELECTORS:
        el = soup.select_one(sel)
        if el:
            txt = _norm_txt(el.get_text(" ", strip=True))
            if txt:
                wynik["cena"] = txt
                return

def _address_from_jsonld(soup) -> str:
    for s in soup.find_all("script", attrs={"type": "application/ld+json"}):
        raw = (s.string or "").strip()
        if not raw: continue
        try:
            data = json.loads(raw)
        except Exception:
            try:
                data = json.loads(raw.replace("\n", " ").replace("\t", " "))
            except Exception:
                continue
        items = data if isinstance(data, list) else [data]
        for it in items:
            addr = it.get("address")
            if isinstance(addr, dict):
                parts = []
                street = addr.get("streetAddress")
                locality = addr.get("addressLocality") or addr.get("addressRegion")
                country = addr.get("addressCountry")
                if street: parts.append(_norm_txt(street))
                if locality: parts.append(_norm_txt(locality))
                if country: parts.append(country if isinstance(country, str) else _norm_txt(country.get("name", "")))
                s = ", ".join([p for p in parts if p])
                if s: return s
    return ""

def _pobierz_soup(url: str):
    opts = Options()
    opts.add_argument("--start-maximized")
    # opts.add_argument("--headless=new")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    opts.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )

    driver = webdriver.Chrome(options=opts)
    try:
        try:
            driver.execute_cdp_cmd(
                "Page.addScriptToEvaluateOnNewDocument",
                {"source": "Object.defineProperty(navigator,'webdriver',{get:() => undefined})"},
            )
        except Exception:
            pass

        last_err = None
        for attempt in range(3):
            try:
                driver.get(url)
                break
            except Exception as e:
                last_err = e
                time.sleep(1.5 * (attempt + 1))
        if last_err and attempt == 2:
            raise last_err

        wait = WebDriverWait(driver, WAIT_SECONDS)
        try:
            wait.until(
                EC.any_of(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "meta[name='description']")),
                    EC.presence_of_element_located((By.CSS_SELECTOR, "strong[data-testid='ad-price']")),
                    EC.presence_of_element_located((By.CSS_SELECTOR, "link[rel='canonical']")),
                )
            )
        except Exception:
            time.sleep(min(5, WAIT_SECONDS))

        html = driver.page_source
    finally:
        driver.quit()
    return BeautifulSoup(html, "html.parser")

def _wyciagnij_adres_string(soup) -> str:
    mapa_link = soup.select_one("a[href='#map']") or soup.select_one("a[href*='mapa']") \
        or soup.find("span", attrs={"data-cy": "adPageAdLocalisation"})
    if mapa_link:
        t = mapa_link.get_text(strip=True)
        print("📍 Adres z UI:", t)
        return t
    return ""

def _fill_from_jsonld(obj, wynik: dict):
    if not isinstance(obj, (dict, list)): return
    items = obj if isinstance(obj, list) else [obj]
    for it in items:
        if not isinstance(it, dict): continue
        offers = it.get("offers") or it.get("offer")
        if isinstance(offers, dict):
            if "cena" not in wynik and offers.get("price"):
                wynik["cena"] = format_pln_int(offers.get("price")) or wynik.get("cena")
        floor_size = it.get("floorSize") or it.get("floorArea")
        if isinstance(floor_size, dict) and "metry" not in wynik:
            v = floor_size.get("value") or it.get("valueReference")
            m2 = fmt_metry(v)
            if m2: wynik["metry"] = m2
        if "liczba_pokoi" not in wynik and it.get("numberOfRooms"):
            wynik["liczba_pokoi"] = str(_num(str(it.get("numberOfRooms"))))
        if "rok_budowy" not in wynik and it.get("yearBuilt"):
            wynik["rok_budowy"] = str(it["yearBuilt"])
        if "material" not in wynik and it.get("material"):
            mat = it.get("material")
            if isinstance(mat, str): wynik["material"] = mat
        if "rynek" not in wynik and it.get("itemCondition"):
            cond = str(it["itemCondition"]).lower()
            if "new" in cond: wynik["rynek"] = "pierwotny"
            elif "used" in cond: wynik["rynek"] = "wtórny"

def pobierz_dane_z_otodom(url: str) -> dict:
    soup = _pobierz_soup(url)
    wynik = {}
    _extract_price_from_dom(soup, wynik)

    opis = soup.find("meta", attrs={"name": "description"})
    if opis and opis.has_attr("content"):
        content = opis["content"]
        if "cena" not in wynik:
            cena_match = re.search(r"za cenę ([\d\s]+zł)", content)
            if cena_match: wynik["cena"] = cena_match.group(1).strip()
        metry_match = re.search(r"ma ([\d.,]+ m²)", content)
        if metry_match: wynik["metry"] = metry_match.group(1).strip()
        pietro_match = re.search(r"na ([^,\.]+? piętrze)", content)
        if pietro_match:
            pietro_raw = (pietro_match.group(1) or "").strip().lower()
            if "parter" in pietro_raw:
                wynik["pietro"] = "parter"
            else:
                tylko_numer = re.search(r"\d+", pietro_raw)
                wynik["pietro"] = tylko_numer.group() if tylko_numer else ""

    for sel in [
        "strong[data-testid='ad-price-per-m']","div[data-testid='ad-price-per-m']",
        "p[data-testid='ad-price-per-m']","span[data-testid='ad-price-per-m']",
        "[data-testid='ad-price-per-m2']","[data-cy='ad-price-per-m']",
    ]:
        el = soup.select_one(sel)
        if el:
            wynik["cena_za_metr"] = el.get_text(strip=True)
            break

    szczegoly = soup.select("div[data-testid='table-value']")
    for item in szczegoly:
        label_elem = item.select_one("div[data-testid='table-value-title']")
        value_elem = item.select_one("div[data-testid='table-value-value']")
        if label_elem and value_elem:
            label = label_elem.get_text(strip=True)
            value = value_elem.get_text(strip=True)
            if label in szukane_pola and szukane_pola[label] not in wynik:
                wynik[szukane_pola[label]] = value

    try:
        for s in soup.find_all("script", attrs={"type": "application/ld+json"}):
            raw = (s.string or "").strip()
            if not raw: continue
            try:
                data = json.loads(raw)
                _fill_from_jsonld(data, wynik)
            except Exception:
                try:
                    data = json.loads(raw.replace("\n", " ").replace("\t", " "))
                    _fill_from_jsonld(data, wynik)
                except Exception:
                    pass
    except Exception:
        pass

    adres_string = _wyciagnij_adres_string(soup)
    adr_jsonld_str = _address_from_jsonld(soup)
    if adr_jsonld_str and len(adr_jsonld_str) > len(adres_string):
        adres_string = adr_jsonld_str

    adr = parsuj_adres_string(adres_string)
    try:
        adr = _safe_enrich_address(adr)
    except Exception as e:
        print(f"[WARN] Enrichment failed: {e}")

    if not (adr.get("ulica_nazwa") or "").strip():
        try:
            for s in soup.find_all("script", attrs={"type": "application/ld+json"}):
                raw = (s.string or "").strip()
                if not raw: continue
                data = json.loads(raw.replace("\n", " ").replace("\t", " "))
                items = data if isinstance(data, list) else [data]
                for it in items:
                    addr = it.get("address")
                    if isinstance(addr, dict):
                        street = addr.get("streetAddress")
                        if street:
                            nazwa = re.sub(r"\s+\d.*$", "", street).strip()
                            adr["ulica_nazwa"] = nazwa or street.strip()
                            break
                else:
                    continue
                break
        except Exception:
            pass

    adr["gmina"] = _clean_gmina(adr.get("gmina"))
    dzielnica_csv = adr.get("dzielnica") or ""
    ulica_csv = adr.get("ulica_nazwa") or ""

    canonical = soup.find("link", rel="canonical")
    wynik["link"] = canonical["href"] if canonical and canonical.has_attr("href") else url

    if not wynik.get("cena_za_metr"):
        wynik["cena_za_metr"] = oblicz_cena_za_metr(wynik.get("cena", ""), wynik.get("metry", ""))

    out = {
        "cena": wynik.get("cena", ""),
        "cena_za_metr": wynik.get("cena_za_metr", ""),
        "metry": wynik.get("metry", ""),
        "liczba_pokoi": wynik.get("liczba_pokoi", ""),
        "pietro": wynik.get("pietro", ""),
        "rynek": wynik.get("rynek", ""),
        "rok_budowy": wynik.get("rok_budowy", ""),
        "material": wynik.get("material", ""),
        "wojewodztwo": adr.get("wojewodztwo") or "",
        "powiat": adr.get("powiat") or "",
        "gmina": adr.get("gmina") or "",
        "miejscowosc": adr.get("miasto") or "",
        "dzielnica": dzielnica_csv,
        "ulica": ulica_csv,
        "link": wynik["link"],
    }
    _consistency_pass_row(out)
    return out

# ------------------------------ CSV I/O (ETYKIETY) ------------------------------
def _links_paths_candidates(label: str, slug: str) -> list[Path]:
    return [
        BASE_LINKI_DIR / f"{label}.csv",
        BASE_LINKI_DIR / f"{slug}.csv",            # legacy
        BASE_LINKI_DIR / f"intake_{label}.csv",    # legacy
        BASE_LINKI_DIR / f"intake_{slug}.csv",     # legacy
    ]

def _find_links_csv(label: str, slug: str) -> Path | None:
    for p in _links_paths_candidates(label, slug):
        if p.exists() and p.stat().st_size > 0:
            return p
    return None

def _ensure_results_csv_label(label: str) -> Path:
    p = BASE_WOJ_DIR / f"{label}.csv"
    if not p.exists():
        p.parent.mkdir(parents=True, exist_ok=True)
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
    _ensure_results_csv_label(output_csv.stem)
    with output_csv.open("a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        for r in rows:
            w.writerow([r.get(col, "") for col in KOLUMNY_WYJ])

# ------------------------------ PIPE ------------------------------
def przetworz_linki_do_csv(input_csv: Path, out_csv: Path):
    try:
        links = _read_links_from_csv(input_csv)
    except Exception as e:
        raise SystemExit(f"Nie mogę odczytać linków z {input_csv}: {e}")

    print(f"[INFO] Do przetworzenia linków: {len(links)}")
    if not links:
        print("[INFO] Brak linków — nic do zrobienia.")
        _ensure_results_csv_label(out_csv.stem)
        return

    dane_lista: list[dict] = []
    for i, link in enumerate(links, start=1):
        print(f"\n🌐 [{i}/{len(links)}] {link}")
        dane = None
        for attempt in range(2):
            try:
                dane = pobierz_dane_z_otodom(link)
                break
            except Exception as e:
                if _is_remote_closed(e) and attempt == 0:
                    print(f"[WARN] Problem sieciowy: {e} — ponawiam…")
                    time.sleep(2.0)
                    continue
                print(f"❌ Błąd przy linku: {e}")
                dane = None
                break
        if not dane:
            continue
        for k in KOLUMNY_WYJ:
            dane.setdefault(k, "")
        if (dane.get("cena") or "").strip():
            dane_lista.append(dane)

    _append_rows_to_output_csv(out_csv, dane_lista)
    print(f"\n✅ Zapisano {len(dane_lista)} rekordów do {out_csv}")

# ------------------------------ CLI ------------------------------
if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--region", "-r", required=True, help="Etykieta (np. 'Mazowieckie') lub slug (np. 'mazowieckie')")
    args = ap.parse_args()

    label, slug = normalize_region_single(args.region)
    input_csv = _find_links_csv(label, slug)
    if not input_csv:
        raise SystemExit(
            "Nie znaleziono pliku z linkami.\n"
            f"Szukane: {BASE_LINKI_DIR / (label + '.csv')} oraz warianty legacy."
        )

    out_csv = BASE_WOJ_DIR / f"{label}.csv"
    print(f"[INFO] Czytam linki z: {input_csv}")
    print(f"[INFO] Zapis wyników do: {out_csv}")
    przetworz_linki_do_csv(input_csv, out_csv)
