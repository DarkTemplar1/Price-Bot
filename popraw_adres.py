#!/usr/bin/env python3
from __future__ import annotations

import argparse
import sys
from typing import List, Dict, Any, Optional
import pandas as pd
from pathlib import Path

COL_NR_KW = "Nr KW"
ADDR_COLS = ["Województwo", "Powiat", "Gmina", "Miejscowość", "Dzielnica", "Ulica"]
PRICE_COLS = [
    "Średnia cena za m2 ( z bazy)",
    "Średnia skorygowana cena za m2",
    "Statyczna wartość nieruchomości",
]

# —————————————————— Wbudowane mapowanie starych/błędnych nazw → nowe ——————————————————
# klucze po lewej są normalizowane (casefold, bez zbędnych spacji), wartości po prawej to pożądana forma
# Sekcja "global" działa dla wszystkich kolumn adresowych, sekcja per-kolumna — tylko w tej kolumnie.
BUILTIN_NAME_MAP: Dict[Optional[str], Dict[str, str]] = {
    None: {  # globalnie (miasta — korekta polskich znaków i popularnych zapisów)
        "lodz": "Łódź",
        "krakow": "Kraków",
        "poznan": "Poznań",
        "wroclaw": "Wrocław",
        "bialystok": "Białystok",
        "radom": "Radom",
        "szczecin": "Szczecin",
        "zielona gora": "Zielona Góra",
        "gorzow wlkp.": "Gorzów Wielkopolski",
        "gorzow wielkopolski": "Gorzów Wielkopolski",
        "koszalin": "Koszalin",
        "bielsko biala": "Bielsko-Biała",
        "warszawa": "Warszawa",
        "gdansk": "Gdańsk",
        "gdynia": "Gdynia",
        "sopot": "Sopot",
        "rzeszow": "Rzeszów",
        "olsztyn": "Olsztyn",
        "torun": "Toruń",
        "bydgoszcz": "Bydgoszcz",
        "katowice": "Katowice",
        "opole": "Opole",
        "kielce": "Kielce",
        "biala podlaska": "Biała Podlaska",
        "nowy sacz": "Nowy Sącz",
    },
    "Województwo": {  # poprawne nazwy województw
        "dolnoslaskie": "Dolnośląskie",
        "kujawsko pomorskie": "Kujawsko-Pomorskie",
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
        "warminsko mazurskie": "Warmińsko-Mazurskie",
        "warminsko-mazurskie": "Warmińsko-Mazurskie",
        "wielkopolskie": "Wielkopolskie",
        "zachodniopomorskie": "Zachodniopomorskie",
    },
    "Dzielnica": {
        "srodmiescie": "Śródmieście",
        "praga polnoc": "Praga-Północ",
        "praga poludnie": "Praga-Południe",
        "bialoleka": "Białołęka",
        "bemowo": "Bemowo",
        "wola": "Wola",
        "ursynow": "Ursynów",
        "wlochy": "Włochy",
        "targowek": "Targówek",
        "mokotow": "Mokotów",
        "ziebice": "Ziębice",
    },
}

# —————————————————— Narzędzia normalizacji / mapowania ——————————————————
def norm_val(v: Any) -> str:
    if pd.isna(v):
        return ""
    s = " ".join(str(v).strip().split())  # zredukuj wielokrotne spacje
    if s.lower() in {"", "nan", "none", "null"}:
        return ""
    return s.casefold()

def ensure_columns(df: pd.DataFrame, required: List[str], context: str) -> None:
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            f"W pliku '{context}' brakuje kolumn: {missing}. "
            f"Dostępne kolumny: {list(df.columns)}"
        )

def build_mask(ter: pd.DataFrame, probe: Dict[str, Any], cols_for_matching: List[str]) -> pd.Series:
    mask = pd.Series(True, index=ter.index)
    for c in cols_for_matching:
        if c in probe:
            v = norm_val(probe[c])
            if v != "":
                mask = mask & (ter[f"{c}__norm"] == v)
    return mask

def concat_address(row: pd.Series, cols: List[str]) -> str:
    parts = []
    for c in ["Województwo", "Powiat", "Gmina", "Miejscowość", "Dzielnica", "Ulica"]:
        if c in cols:
            val = str(row.get(c, "")).strip()
            if val and val.lower() not in {"nan", "none", "null"}:
                parts.append(val)
    return ", ".join(parts)

def _strip_ulica_prefixes(name: str) -> str:
    # usuń częste prefiksy: ul., UL., ul  — oraz standaryzuj skróty al./pl./os.
    s = name.strip()
    low = s.casefold()
    for pref in ["ul.", "ul ", "ul. ", "ulica ", "ul "]:
        if low.startswith(pref):
            s = s[len(pref):].lstrip()
            break
    return " ".join(s.split())

def load_custom_mapping(path: Path) -> Dict[Optional[str], Dict[str, str]]:
    """
    Wczytuje CSV/XLSX o kolumnach:
      - 'stara' (wymagana)
      - 'nowa'  (wymagana)
      - 'kolumna' (opcjonalna: np. 'Województwo', 'Miejscowość'; jeśli brak → mapowanie globalne)
    Zwraca strukturę jak BUILTIN_NAME_MAP.
    """
    if not path.exists():
        raise FileNotFoundError(f"Nie znaleziono pliku mapowania: {path}")
    if path.suffix.lower() in {".xlsx", ".xlsm"}:
        df = pd.read_excel(path, dtype=str)
    else:
        df = pd.read_csv(path, dtype=str, sep=None, engine="python")
    for required in ["stara", "nowa"]:
        if required not in df.columns:
            raise ValueError(f"Plik mapowania musi mieć kolumnę '{required}'.")
    result: Dict[Optional[str], Dict[str, str]] = {}
    for _, r in df.iterrows():
        old = norm_val(r["stara"])
        new = str(r["nowa"]).strip()
        col = str(r["kolumna"]).strip() if "kolumna" in df.columns and pd.notna(r["kolumna"]) else None
        if not old or not new:
            continue
        bucket = result.setdefault(col if col in ADDR_COLS else None, {})
        bucket[old] = new
    return result

def make_name_resolver(custom_map: Optional[Dict[Optional[str], Dict[str, str]]] = None):
    """
    Tworzy funkcję fix_name(value, column) zwracającą poprawioną nazwę wg:
    1) mapowań kolumnowych (custom → builtin),
    2) mapowań globalnych,
    3) lekkiej normalizacji (spacje, prefiksy 'ul.').
    """
    merged: Dict[Optional[str], Dict[str, str]] = {}
    # start od wbudowanych
    for k, v in BUILTIN_NAME_MAP.items():
        merged[k] = dict(v)
    # nadpisz/custom
    if custom_map:
        for k, v in custom_map.items():
            merged.setdefault(k, {})
            merged[k].update({norm_val(kk): vv for kk, vv in v.items()})

    def fix_name(value: Any, column: str) -> str:
        raw = "" if pd.isna(value) else str(value)
        clean = " ".join(raw.strip().split())
        key = norm_val(clean)

        # specyficzne dla ulic
        if column == "Ulica":
            clean = _strip_ulica_prefixes(clean)
            key = norm_val(clean)

        # 1) kolumnowe
        colmap = merged.get(column, {})
        if key in colmap:
            return colmap[key]

        # 2) globalne
        glob = merged.get(None, {})
        if key in glob:
            return glob[key]

        # 3) brak mapy — zwróć oczyszczoną wartość (np. bez 'ul.')
        return clean

    return fix_name

# —————————————————— Główna procedura ——————————————————
def main() -> int:
    parser = argparse.ArgumentParser(description="Dopasowanie adresów z TERYT + korekta starych nazw na nowe.")
    parser.add_argument("--we", "--wejscie", dest="wejscie", required=True, help="Ścieżka do pliku wejściowego Excel.")
    parser.add_argument("--teryt", dest="teryt", required=True, help="Ścieżka do pliku TERYT.xlsx.")
    parser.add_argument("--arkusz", dest="arkusz", default=0, help="Nazwa lub indeks arkusza (domyślnie pierwszy).")
    parser.add_argument("--zapis", dest="zapis", required=True, help="Ścieżka pliku wyjściowego Excel.")
    parser.add_argument("--mapa", dest="mapa", default=None,
                        help="OPCJONALNIE: plik CSV/XLSX z kolumnami: stara, nowa[, kolumna] – dodatkowe mapowania nazw.")
    args = parser.parse_args()

    # wejście
    try:
        df = pd.read_excel(args.wejscie, sheet_name=args.arkusz, dtype=str)
    except Exception as e:
        print(f"Nie udało się wczytać pliku wejściowego: {e}", file=sys.stderr)
        return 2

    try:
        teryt = pd.read_excel(args.teryt, dtype=str)
    except Exception as e:
        print(f"Nie udało się wczytać pliku TERYT: {e}", file=sys.stderr)
        return 2

    ensure_columns(df, [COL_NR_KW] + ADDR_COLS + PRICE_COLS, context=args.wejscie)
    base_teryt_cols = ["Województwo", "Powiat", "Gmina", "Miejscowość", "Dzielnica"]
    ensure_columns(teryt, base_teryt_cols, context=args.teryt)

    # przygotuj kolumny znormalizowane w TERYT (do filtrowania)
    teryt_has_ulica = "Ulica" in teryt.columns
    for c in base_teryt_cols + (["Ulica"] if teryt_has_ulica else []):
        teryt[f"{c}__norm"] = teryt[c].apply(norm_val) if c in teryt.columns else ""

    # ewentualnie dodaj brakujące kolumny cenowe w wejściu
    for col in PRICE_COLS:
        if col not in df.columns:
            df[col] = ""

    # wczytaj dodatkowe mapowania (jeśli podano)
    custom_map = None
    if args.mapa:
        try:
            custom_map = load_custom_mapping(Path(args.mapa))
        except Exception as e:
            print(f"Uwaga: nie udało się wczytać mapowania '{args.mapa}': {e}", file=sys.stderr)

    # stwórz resolver nazw i zastosuj go do kolumn adresowych wejścia
    fix_name = make_name_resolver(custom_map)
    for c in ADDR_COLS:
        if c in df.columns:
            df[c] = df[c].apply(lambda v, col=c: fix_name(v, col))

    processed = 0
    for idx, row in df.iterrows():
        nr_kw = norm_val(row[COL_NR_KW])
        if nr_kw == "":
            continue
        processed += 1

        addr_values = {c: row.get(c, "") for c in ADDR_COLS}
        filled_count = sum(1 for c, v in addr_values.items() if norm_val(v) != "")

        if filled_count == 0:
            for col in PRICE_COLS:
                df.at[idx, col] = "brak adresu"
            continue

        if filled_count == 2:
            probe = {c: addr_values.get(c, "") for c in base_teryt_cols}
            cols_for_matching = [c for c in base_teryt_cols if norm_val(probe[c]) != ""]
            if not cols_for_matching:
                for col in PRICE_COLS:
                    df.at[idx, col] = "brak adresu"
                continue

            mask = build_mask(teryt, probe, cols_for_matching)
            candidates = teryt[mask]

            if len(candidates) == 1:
                for c in base_teryt_cols:
                    if norm_val(df.at[idx, c]) == "":
                        df.at[idx, c] = candidates.iloc[0][c]
                if teryt_has_ulica and norm_val(df.at[idx, "Ulica"]) == "" and "Ulica" in candidates.columns:
                    df.at[idx, "Ulica"] = candidates.iloc[0].get("Ulica", "")
            elif len(candidates) > 1:
                addr_list = []
                for _, r in candidates.iterrows():
                    addr_list.append(concat_address(r, base_teryt_cols + (["Ulica"] if teryt_has_ulica else [])))
                unique_addrs = sorted(set(a for a in addr_list if a))
                multi = " | ".join(unique_addrs[:50])
                for col in PRICE_COLS:
                    df.at[idx, col] = multi if multi else "brak adresu"
            else:
                for col in PRICE_COLS:
                    df.at[idx, col] = "brak adresu"
            continue

        if filled_count >= 3:
            probe = {c: addr_values.get(c, "") for c in base_teryt_cols + (["Ulica"] if teryt_has_ulica else [])}
            cols_for_matching = [c for c, v in probe.items() if norm_val(v) != ""]
            if not cols_for_matching:
                for col in PRICE_COLS:
                    df.at[idx, col] = "brak adresu"
                continue

            mask = build_mask(teryt, probe, cols_for_matching)
            candidates = teryt[mask]

            if len(candidates) == 1:
                for c in base_teryt_cols:
                    if norm_val(df.at[idx, c]) == "":
                        df.at[idx, c] = candidates.iloc[0][c]
                if teryt_has_ulica:
                    if norm_val(df.at[idx, "Ulica"]) == "" and "Ulica" in candidates.columns:
                        df.at[idx, "Ulica"] = candidates.iloc[0].get("Ulica", "")
            elif len(candidates) > 1:
                addr_list = []
                for _, r in candidates.iterrows():
                    addr_list.append(concat_address(r, base_teryt_cols + (["Ulica"] if teryt_has_ulica else [])))
                unique_addrs = sorted(set(a for a in addr_list if a))
                multi = " | ".join(unique_addrs[:50])
                for col in PRICE_COLS:
                    df.at[idx, col] = multi if multi else "brak adresu"
            else:
                for col in PRICE_COLS:
                    df.at[idx, col] = "brak adresu"
            continue

    try:
        with pd.ExcelWriter(args.zapis, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
    except Exception as e:
        print(f"Nie udało się zapisać pliku wyjściowego: {e}", file=sys.stderr)
        return 2

    print(f"Zakończono. Przetworzono wierszy z Nr KW: {processed}. Zapisano: {args.zapis}")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
