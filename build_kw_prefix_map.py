# -*- coding: utf-8 -*-
import re
import unicodedata
from pathlib import Path
from typing import Optional, List, Dict

import pandas as pd
import requests
from bs4 import BeautifulSoup

ELI_URL = "https://eli.gov.pl/api/acts/DU/2015/1613/text.html"
KW_RE = re.compile(r"^[A-Z]{2}\d[A-Z0-9]{1,2}$")

DIAC_MAP = str.maketrans({
    "ą":"a","ć":"c","ę":"e","ł":"l","ń":"n","ó":"o","ś":"s","ź":"z","ż":"z",
    "Ą":"A","Ć":"C","Ę":"E","Ł":"L","Ń":"N","Ó":"O","Ś":"S","Ź":"Z","Ż":"Z",
})

STOP_SPLITS = [
    r"\s*[-–]?\s*dla\b",
    r"\s*w\s+granicach\b",
    r"\s*z\s+siedzibą\b",
    r"\s*w\s+wydziale\b",
    r"\s*obejmuj(?:ących|ącą|ący|ące)\b",
    r"\s*oraz\b",
    r"\s*na\s+obszarze\b",
    r"\s*w\s+tym\b",
]

def clean_text(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    s = unicodedata.normalize("NFC", s)
    s = s.replace("\xa0"," ").replace("\u202f"," ").replace("\u2009"," ")
    for ch in ("\u00ad","\u200b","\u200c","\u200d","\ufeff"):
        s = s.replace(ch, "")
    s = re.sub(r"\s+", " ", s).strip()
    return s

def norm_key(s: str) -> str:
    s = clean_text(s).lower()
    s = s.translate(DIAC_MAP)
    return s

def cut_descriptors(s: str) -> str:
    s = clean_text(s)
    s = re.split("|".join(STOP_SPLITS), s, flags=re.I)[0]
    return s.strip(" .,-;:")

def fetch_eli_html(url: str = ELI_URL) -> str:
    r = requests.get(url, timeout=60, headers={"User-Agent":"Mozilla/5.0"})
    if not r.encoding or r.encoding.lower() in ("ascii","iso-8859-1"):
        r.encoding = r.apparent_encoding or "utf-8"
    else:
        r.encoding = "utf-8"
    r.raise_for_status()
    return r.text

def parse_codes_from_table(html: str) -> pd.DataFrame:
    try:
        soup = BeautifulSoup(html, "lxml")
    except Exception:
        soup = BeautifulSoup(html, "html.parser")

    rows = []
    for table in soup.find_all("table"):
        for tr in table.find_all("tr"):
            cells = [clean_text(td.get_text(" ", strip=True)) for td in tr.find_all(["td","th"])]
            if len(cells) < 2:
                continue
            code = cells[-1]
            if KW_RE.match(code):
                src = cells[-2]
                parts = re.findall(r"\bw\s+([^,;]+)", src, flags=re.I)
                town = clean_text(parts[-1]) if parts else src
                town = cut_descriptors(town)
                town = re.split(r"\s*[-–]\s*wydział.*|\s*z siedzibą.*|\s*w wydziale.*", town, flags=re.I)[0]
                town = re.sub(r"[()\[\]]", "", town).strip(" .")
                rows.append({"KW_PREFIX": code, "Miejscowość": town})

    if not rows:
        text = clean_text(BeautifulSoup(html, "html.parser").get_text("\n", strip=True))
        last_place = None
        for line in [l for l in text.splitlines() if l.strip()]:
            if re.match(r"^(w|we)\s", line, re.I) or "wydział zamiejscowy z siedzibą w" in line.lower():
                last_place = line
            m = KW_RE.match(line)
            if m and last_place:
                parts = re.findall(r"\bw\s+([^,;]+)", last_place, flags=re.I)
                town = clean_text(parts[-1] if parts else last_place)
                town = cut_descriptors(town)
                town = re.split(r"\s*[-–]\s*wydział.*|\s*z siedzibą.*|\s*w wydziale.*", town, flags=re.I)[0]
                town = re.sub(r"[()\[\]]", "", town).strip(" .")
                rows.append({"KW_PREFIX": m.group(0), "Miejscowość": town})
                last_place = None

    return pd.DataFrame(rows).drop_duplicates(subset=["KW_PREFIX"]).reset_index(drop=True)

def load_teryt_df(teryt_path: Optional[Path]) -> Optional[pd.DataFrame]:
    if not teryt_path or not Path(teryt_path).exists():
        return None
    df = pd.read_excel(teryt_path)
    required = ["Województwo","Powiat","Gmina","Miejscowość","Dzielnica"]
    for c in required:
        if c not in df.columns:
            raise ValueError(f"Brak kolumny '{c}' w TERYT.xlsx")
        df[c] = df[c].fillna("").map(lambda s: str(s).strip())
    return df

def guess_nominative(locative: str, teryt_df: Optional[pd.DataFrame]) -> str:
    base = re.sub(r"^(w|we)\s+", "", clean_text(locative), flags=re.I).strip()
    if not base:
        return ""

    if teryt_df is not None:
        sub = teryt_df[teryt_df["Miejscowość"].str.lower() == base.lower()]
        if len(sub) > 0:
            return sub["Miejscowość"].iloc[0]

    cand = [base]
    low = base.lower()
    if low.endswith("owie"): cand.append(base[:-4] + "ów")
    if low.endswith("nie"):
        cand.append(base[:-2])         # Olsztynie -> Olsztyn
        cand.append(base[:-3] + "no")  # Gnieźnie -> Gniezno
    if low.endswith("wie"):
        cand.append(base[:-2] + "a")   # Warszawie -> Warszawa
        cand.append(base[:-3] + "wa")
    if low.endswith("ni"): cand.append(base + "a")   # Gdyni -> Gdynia
    if low.endswith("i") and not low.endswith("ni"): cand.append(base + "a")
    if low.endswith("y"): cand.append(base[:-1] + "a")  # Oleśnicy -> Oleśnica
    if low.endswith("ie"):
        cand.append(base[:-2] + "ia")
        cand.append(base[:-2])
    if low.endswith("u"):
        cand.append(base[:-1] + "iec")  # Żywcu -> Żywiec
        cand.append(base[:-1] + "ów")

    if teryt_df is not None and not teryt_df.empty:
        towns = teryt_df["Miejscowość"].dropna().unique().tolist()
        towns_norm = [norm_key(t) for t in towns]
        for c in cand:
            nk = norm_key(c)
            best_i, best_score = -1, -10**9
            for i, tn in enumerate(towns_norm):
                pref = 0
                for a, b in zip(nk, tn):
                    if a == b: pref += 1
                    else: break
                score = pref - abs(len(nk) - len(tn))
                if score > best_score:
                    best_score, best_i = score, i
            if best_i >= 0 and best_score >= max(1, len(nk)-2):
                return towns[best_i]

    return cand[0]

def build_teryt_index(df: pd.DataFrame) -> Dict[str, Dict[str, List[str]]]:
    df = df.copy()
    df["key"] = df["Miejscowość"].map(norm_key)
    grp = df.groupby("key", dropna=False)
    index = {}
    for k, g in grp:
        index[k] = {
            "woj": list(dict.fromkeys(g["Województwo"].tolist())),
            "pow": list(dict.fromkeys(g["Powiat"].tolist())),
            "gmi": list(dict.fromkeys(g["Gmina"].tolist())),
            "town_variants": list(dict.fromkeys(g["Miejscowość"].tolist()))
        }
    return index

def choose_powiat(miejsc: str, cand_powiaty: List[str]) -> str:
    if not cand_powiaty:
        return ""
    if len(cand_powiaty) == 1:
        return cand_powiaty[0]
    nk = norm_key(miejsc)
    for p in cand_powiaty:
        if norm_key(p).startswith("m ") and nk in norm_key(p):
            return p
    for p in cand_powiaty:
        if nk in norm_key(p):
            return p
    return ""

def choose_gmina(miejsc: str, cand_gminy: List[str]) -> str:
    if not cand_gminy:
        return ""
    if len(cand_gminy) == 1:
        return cand_gminy[0]
    nk = norm_key(miejsc)
    for g in cand_gminy:
        if norm_key(g) == nk:
            return g
    for g in cand_gminy:
        if nk in norm_key(g):
            return g
    return ""

def enrich_with_teryt(df_codes: pd.DataFrame, teryt_df: Optional[pd.DataFrame]) -> pd.DataFrame:
    df_codes = df_codes.copy()
    df_codes["Miejscowość"] = df_codes["Miejscowość"].map(lambda x: guess_nominative(x, teryt_df))

    if teryt_df is None or teryt_df.empty:
        df_codes["Województwo"] = ""
        df_codes["Powiat"] = ""
        df_codes["Gmina"] = ""
        return df_codes

    idx = build_teryt_index(teryt_df)

    woj_out, pow_out, gmi_out = [], [], []
    for _, row in df_codes.iterrows():
        key = norm_key(row["Miejscowość"])
        bucket = idx.get(key)
        if not bucket:
            woj_out.append(""); pow_out.append(""); gmi_out.append("")
            continue
        woj = bucket["woj"][0] if len(bucket["woj"]) == 1 else ""
        powiat = choose_powiat(row["Miejscowość"], bucket["pow"])
        gmina = choose_gmina(row["Miejscowość"], bucket["gmi"])
        woj_out.append(woj); pow_out.append(powiat); gmi_out.append(gmina)

    df_codes["Województwo"] = woj_out
    df_codes["Powiat"] = pow_out
    df_codes["Gmina"] = gmi_out
    return df_codes

def build_and_save(out_path: Path, teryt_path: Optional[Path] = Path("TERYT.xlsx")) -> pd.DataFrame:
    html = fetch_eli_html()
    df = parse_codes_from_table(html)
    teryt_df = load_teryt_df(teryt_path) if teryt_path else None
    df = enrich_with_teryt(df, teryt_df)
    cols = ["KW_PREFIX","Województwo","Powiat","Gmina","Miejscowość"]
    df = df[cols].sort_values("KW_PREFIX").reset_index(drop=True)
    df.to_excel(out_path, index=False, engine="openpyxl")
    return df

if __name__ == "__main__":
    out = Path("KW_prefix_map.xlsx")
    teryt = Path("TERYT.xlsx") if Path("TERYT.xlsx").exists() else None
    df = build_and_save(out, teryt)
    print(f"Zapisano: {out} (rekordów: {len(df)})")
