from __future__ import annotations
"""
adres_otodom.py — wersja BEZ zależności od SOAP/zeep

Co zostało:
- Te same funkcje publiczne jak wcześniej, aby scraper działał bez zmian:
  set_contact_email, set_teryt_credentials (no-op), parsuj_adres_string,
  uzupelnij_braki_z_nominatim, dopelnij_powiat_gmina_jesli_brak,
  _clean_gmina, _tylko_dzielnica, _consistency_pass_row.

Co się zmieniło:
- Usunięto import i użycie biblioteki 'zeep' (SOAP TERYT).
- „Uzupełnianie” odbywa się lokalnie (heurystyki):
    • Województwo – z mapy znanych miast (VOIVODESHIP_BY_CITY).
    • Miasta na prawach powiatu – ustaw powiat i gminę = miasto.
    • Porządki nazw (prefiksy „woj./powiat/gmina”, wielkość liter, itp.).
- Funkcja set_teryt_credentials istnieje, ale nic nie robi (dla kompatybilności).
"""

import dataclasses
import re
from typing import Any, Dict, Optional

# ================== KONFIG / STAŁE ==================
CONTACT_EMAIL = "twoj_email@domena.pl"  # można nadpisać przez set_contact_email()

def set_contact_email(email: str):
    """Ustaw e-mail do identyfikacji (zachowane dla kompatybilności)."""
    global CONTACT_EMAIL
    if email and email.strip():
        CONTACT_EMAIL = email.strip()

def set_teryt_credentials(*args, **kwargs):
    """No-op (zachowane dla kompatybilności)."""
    return

# ================== LISTY / SŁOWNIKI ADMIN ==================
VOIVODESHIPS = [
    "Dolnośląskie","Kujawsko-Pomorskie","Lubelskie","Lubuskie","Łódzkie","Małopolskie",
    "Mazowieckie","Opolskie","Podkarpackie","Podlaskie","Pomorskie","Śląskie",
    "Świętokrzyskie","Warmińsko-Mazurskie","Wielkopolskie","Zachodniopomorskie",
]
VOIVODESHIPS_LOWER = {v.lower(): v for v in VOIVODESHIPS}

VOIVODESHIP_BY_CITY = {
    "Warszawa":"Mazowieckie","Kraków":"Małopolskie","Łódź":"Łódzkie","Wrocław":"Dolnośląskie",
    "Poznań":"Wielkopolskie","Gdańsk":"Pomorskie","Gdynia":"Pomorskie","Sopot":"Pomorskie",
    "Szczecin":"Zachodniopomorskie","Bydgoszcz":"Kujawsko-Pomorskie","Toruń":"Kujawsko-Pomorskie",
    "Lublin":"Lubelskie","Białystok":"Podlaskie","Katowice":"Śląskie","Kielce":"Świętokrzyskie",
    "Rzeszów":"Podkarpackie","Olsztyn":"Warmińsko-Mazurskie","Opole":"Opolskie",
    "Zielona Góra":"Lubuskie","Gorzów Wielkopolski":"Lubuskie","Radom":"Mazowieckie","Częstochowa":"Śląskie",
}

CITY_COUNTY = {
    "Warszawa","Kraków","Łódź","Wrocław","Poznań","Gdańsk","Gdynia","Sopot","Szczecin",
    "Bydgoszcz","Toruń","Lublin","Białystok","Katowice","Kielce","Rzeszów","Olsztyn",
    "Opole","Zielona Góra","Gorzów Wielkopolski","Jelenia Góra","Wałbrzych","Ruda Śląska",
    "Gliwice","Zabrze","Bytom","Chorzów","Tychy","Dąbrowa Górnicza","Sosnowiec",
    "Jaworzno","Piekary Śląskie","Świętochłowice","Siemianowice Śląskie","Mysłowice",
    "Tarnów","Nowy Sącz","Legnica","Słupsk","Koszalin","Elbląg","Grudziądz",
}

# ================== HELPERY (wewnętrzne) ==================
def _czysc(t: str) -> str:
    return re.sub(r"\s+", " ", (t or "")).strip()

def _is_voivodeship(seg: str):
    if not seg:
        return None
    s = _czysc(seg).lower().replace("województwo", "").replace("woj.", "").strip()
    return VOIVODESHIPS_LOWER.get(s)

def _clean_gmina(name: str | None) -> str | None:
    if not name:
        return name
    n = name.strip()
    n = re.sub(r'^(gmina|gm\.)\s*(miejska|wiejska|miejsko-wiejska)?\s*', '', n, flags=re.I)
    return _czysc(n)

def _tylko_dzielnica(txt: str | None) -> str:
    if not txt:
        return ""
    return re.split(r'\s*[,/]\s*', txt.strip(), maxsplit=1)[0]

# ================== PARSER ADRESU ==================
def parsuj_adres_string(tekst: str) -> dict:
    """
    Heurystyczny parser adresu z łańcucha Otodom/UI/JSON-LD.
    Zwraca słownik pól: ulica_typ, ulica_nazwa, nr, poddzielnica, dzielnica, miasto,
    gmina, powiat, wojewodztwo, oryginal.
    """
    out = {
        'ulica_typ': None, 'ulica_nazwa': None, 'nr': None,
        'poddzielnica': None, 'dzielnica': None, 'miasto': None, 'gmina': None,
        'powiat': None, 'wojewodztwo': None, 'oryginal': tekst or ''
    }
    if not tekst:
        return out
    t = f" {tekst} "

    # etykietowane
    m = re.search(r'(?:woj\.?|województwo)\s*([A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż \-]+?)(?=,|$)', t, re.I)
    if m:
        w = _czysc(m.group(1))
        out['wojewodztwo'] = VOIVODESHIPS_LOWER.get(w.lower(), w)

    m = re.search(r'powiat\s*([A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż \-]+?)(?=,|$)', t, re.I)
    if m: out['powiat'] = _czysc(m.group(1))

    m = re.search(r'gmina\s*(?:m\.\s*st\.\s*)?([A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż \-]+?)(?=,|$)', t, re.I)
    if m: out['gmina'] = _czysc(m.group(1))

    # ulica + numer
    street_patterns = [
        r'(?:ul\.?|ulica)\s+([A-ZĄĆĘŁŃÓŚŹŻ0-9][\wĄĆĘŁŃÓŚŹŻąćęłńóśźż\.\- ]+?)(?:\s+(\d+[A-Za-z]?(?:/\d+[A-Za-z]?)?))?(?=,|$)',
        r'(?:al\.?|aleja|aleje)\s+([A-ZĄĆĘŁŃÓŚŹŻ0-9][\wĄĆĘŁŃÓŚŹŻąćęłńóśźż\.\- ]+?)(?:\s+(\d+[A-Za-z]?(?:/\d+[A-Za-z]?)?))?(?=,|$)',
        r'(?:pl\.?|plac)\s+([A-ZĄĆĘŁŃÓŚŹŻ0-9][\wĄĆĘŁŃÓŚŹŻąćęłńóśźż\.\- ]+?)(?:\s+(\d+[A-Za-z]?(?:/\d+[A-Za-z]?)?))?(?=,|$)',
        r'(?:os\.?|osiedle)\s+([A-ZĄĆĘŁŃÓŚŹŻ0-9][\wĄĆĘŁŃÓŚŹŻąćęłńóśźż\.\- ]+?)(?:\s+(\d+[A-Za-z]?(?:/\d+[A-Za-z]?)?))?(?=,|$)',
    ]
    for p in street_patterns:
        m = re.search(p, t, re.I)
        if m:
            typ_fragment = re.search(r'^(ul\.?|ulica|al\.?|aleja|aleje|pl\.?|plac|os\.?|osiedle)', m.group(0), re.I)
            out['ulica_typ'] = typ_fragment.group(1).replace('.', '').lower() if typ_fragment else None
            out['ulica_nazwa'] = _czysc(m.group(1))
            if m.group(2): out['nr'] = _czysc(m.group(2))
            break

    # segmenty po przecinkach
    parts = [_czysc(x) for x in t.split(',') if x.strip()]

    # województwo jako osobny segment
    loc_clean = []
    for seg in parts:
        w = _is_voivodeship(seg)
        if w and not out['wojewodztwo']:
            out['wojewodztwo'] = w
        elif not re.search(r'\b(ul\.?|ulica|al\.?|aleja|aleje|pl\.?|plac|os\.?|osiedle|gmina|powiat|Polska)\b', seg, re.I):
            loc_clean.append(seg)

    # heurystyka: ... [poddzielnica], [dzielnica], [miasto]
    if loc_clean:
        out['miasto'] = loc_clean[-1]
        if len(loc_clean) >= 2:
            out['dzielnica'] = _tylko_dzielnica(loc_clean[-2])
        if len(loc_clean) >= 3:
            out['poddzielnica'] = _tylko_dzielnica(loc_clean[-3])

    # korekta kolejności
    if out.get('miasto') and out.get('dzielnica'):
        dz = _czysc(out['dzielnica'])
        if dz in VOIVODESHIP_BY_CITY or dz in CITY_COUNTY:
            out['miasto'], out['dzielnica'] = dz, _tylko_dzielnica(out['miasto'])

    if not out['wojewodztwo'] and out['miasto'] in VOIVODESHIP_BY_CITY:
        out['wojewodztwo'] = VOIVODESHIP_BY_CITY[out['miasto']]

    if out.get('dzielnica'):
        out['dzielnica'] = _tylko_dzielnica(out['dzielnica'])
    if out.get('poddzielnica'):
        out['poddzielnica'] = _tylko_dzielnica(out['poddzielnica'])

    return out

# ====== „Uzupełnianie” HEURYSTYCZNE (zamiast TERYT/Nominatim) ======
def uzupelnij_braki_z_heurystyk(ad: dict) -> dict:
    """Bez sieci – drobne uzupełnienia i porządki."""
    ad = dict(ad or {})

    # jeśli brak miasta – spróbuj z dzielnicy/poddzielnicy
    if not _czysc(ad.get('miasto')) and _czysc(ad.get('dzielnica')):
        ad['miasto'] = _czysc(ad['dzielnica'])

    # województwo z mapy dużych miast
    if not _czysc(ad.get('wojewodztwo')) and _czysc(ad.get('miasto')) in VOIVODESHIP_BY_CITY:
        ad['wojewodztwo'] = VOIVODESHIP_BY_CITY[_czysc(ad['miasto'])]

    # porządki nazw
    if ad.get('gmina'):
        ad['gmina'] = _clean_gmina(ad['gmina'])
    if ad.get('dzielnica'):
        ad['dzielnica'] = _tylko_dzielnica(ad['dzielnica'])
    if ad.get('poddzielnica'):
        ad['poddzielnica'] = _tylko_dzielnica(ad['poddzielnica'])

    _consistency_pass_row(ad)
    return ad

def dopelnij_powiat_gmina_jesli_brak(ad: dict) -> dict:
    """Heurystyczne dopięcie powiat/gmina – miasta na prawach powiatu."""
    ad = dict(ad or {})
    msc = _czysc(ad.get('miasto'))
    if not msc:
        return ad
    if msc in CITY_COUNTY:
        if not _czysc(ad.get('powiat')):
            ad['powiat'] = msc
        if not _czysc(ad.get('gmina')):
            ad['gmina'] = msc
    _consistency_pass_row(ad)
    return ad

# Zachowane aliasy nazw:
def uzupelnij_braki_z_nominatim(ad: dict) -> dict:
    """Alias – lokalne heurystyki (bez sieci)."""
    return uzupelnij_braki_z_heurystyk(ad)

# ====== PORZĄDKI ADMIN (na rekord) ======
def _consistency_pass_row(row: dict):
    """miejscowosc vs powiat/gmina; miasta na prawach powiatu"""
    msc = _czysc(row.get("miejscowosc"))
    powiat = _czysc(row.get("powiat"))
    gmina = _czysc(row.get("gmina"))

    if not msc and gmina:
        row["miejscowosc"] = gmina
        msc = gmina

    if msc and powiat and msc.lower() == powiat.lower() and gmina:
        row["miejscowosc"] = gmina
        msc = gmina

    if msc in CITY_COUNTY:
        if not powiat or powiat.lower() != msc.lower():
            row["powiat"] = msc
        if not gmina or gmina.lower() != msc.lower():
            row["gmina"] = msc

# ====== (opcjonalnie) typy danych zgodne z poprzednią wersją ======
@dataclasses.dataclass
class TerytUnit:  # zgodność, nieużywane
    woj: Optional[str] = None
    powiat: Optional[str] = None
    gmina: Optional[str] = None
    terc: Optional[str] = None

@dataclasses.dataclass
class TerytLocality:  # zgodność, nieużywane
    nazwa: str = ""
    simc: str = ""
    rodzaj_simc: Optional[str] = None
    jednostka: Optional[TerytUnit] = None

class TerytClient:  # STUB, gdyby ktoś stworzył obiekt – informujemy jasno
    def __init__(self, *args, **kwargs):
        raise RuntimeError("TerytClient nie jest dostępny w tej wersji bez SOAP. "
                           "Korzystaj z funkcji heurystycznych (uzupelnij_braki_z_nominatim / dopelnij_powiat_gmina_jesli_brak).")

# ====== EXPORT ======
__all__ = [
    "set_contact_email",
    "set_teryt_credentials",
    "parsuj_adres_string",
    "uzupelnij_braki_z_nominatim",
    "dopelnij_powiat_gmina_jesli_brak",
    "_clean_gmina",
    "_tylko_dzielnica",
    "_consistency_pass_row",
    "TerytClient",           # pozostawione dla kompatybilności (rzuca błąd przy użyciu)
    "TerytUnit", "TerytLocality",
]
