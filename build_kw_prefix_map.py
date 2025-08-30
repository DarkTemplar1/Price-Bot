# -*- coding: utf-8 -*-
import argparse
import re
import unicodedata
from pathlib import Path
from typing import Optional, List, Dict
import difflib
import pandas as pd
import requests
from bs4 import BeautifulSoup

ELI_URL = "https://eli.gov.pl/api/acts/DU/2015/1613/text.html"
KW_RE = re.compile(r"^[A-Z]{2}\d[A-Z0-9]{1,2}$")
FUZZY_THRESHOLD = 0.72

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

IRREG_LOC2NOM = {
    "bielsku-białej":"Bielsko-Biała",
    "białymstoku":"Białystok",
    "bielsku podlaskim":"Bielsk Podlaski",
    "starogardzie gdańskim":"Starogard Gdański",
    "kartuzach":"Kartuzy",
    "gliwicach":"Gliwice",
    "jastrzębiu-zdroju":"Jastrzębie-Zdrój",
    "raciborzu":"Racibórz",
    "rudzie śląskiej":"Ruda Śląska",
    "tarnowskich górach":"Tarnowskie Góry",
    "wodzisławiu śląskim":"Wodzisław Śląski",
    "żorach":"Żory",
    "rybniku":"Rybnik",
    "gorzowie wielkopolskim":"Gorzów Wielkopolski",
    "strzelcach krajeńskich":"Strzelce Krajeńskie",
    "jeleniej górze":"Jelenia Góra",
    "kamiennej górze":"Kamienna Góra",
    "lwówku śląskim":"Lwówek Śląski",
    "katowicach":"Katowice",
    "mysłowicach":"Mysłowice",
    "tychach":"Tychy",
    "busku-zdroju":"Busko-Zdrój",
    "starachowicach":"Starachowice",
    "kielcach":"Kielce",
    "ostrowcu świętokrzyskim":"Ostrowiec Świętokrzyski",
    "skarżysku-kamiennej":"Skarżysko-Kamienna",
    "wadowicach":"Wadowice",
    "myślenicach":"Myślenice",
    "oświęcimiu":"Oświęcim",
    "nowym mieście lubawskim":"Nowe Miasto Lubawskie",
    "środzie śląskiej":"Środa Śląska",
    "świnoujściu":"Świnoujście",
    "stargardzie szczecińskim":"Stargard",
    "ostrówie wielkopolskim":"Ostrów Wielkopolski",
    "zabrzu":"Zabrze",
    "żyrardowie":"Żyrardów",
    "piotrkowie trybunalskim":"Piotrków Trybunalski",
    "tomaszowie mazowieckim":"Tomaszów Mazowiecki",
    "grodzisku mazowieckim":"Grodzisk Mazowiecki",
    "sochaczewie":"Sochaczew",
    "poznaniu":"Poznań",
    "poznańiu":"Poznań",
}

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
    return clean_text(s).lower().translate(DIAC_MAP)

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
    req = ["Województwo","Powiat","Gmina","Miejscowość","Dzielnica"]
    for c in req:
        if c not in df.columns:
            raise ValueError(f"Brak kolumny '{c}' w {teryt_path}")
    return df[req].fillna("").astype(str)

def hyphen_loc2nom_component(w: str) -> str:
    lw = w.lower()
    if lw.endswith("sku"):
        return w[:-3] + "sko"
    if lw.endswith("cku"):
        return w[:-3] + "cko"
    if lw.endswith("ej"):
        return w[:-2] + "a"
    if lw.endswith("owie"):
        return w[:-4] + "ów"
    return w

def hyphen_loc2nom(s: str) -> str:
    return "-".join(hyphen_loc2nom_component(p) for p in s.split("-"))

def generate_candidates(locative: str) -> List[str]:
    b = clean_text(re.sub(r"^(w|we|na)\s+", "", locative, flags=re.I))
    cands = {b}
    if "-" in b:
        cands.add(hyphen_loc2nom(b))
    low = b.lower()
    if low.endswith("owie"): cands.add(b[:-4] + "ów")
    if low.endswith("niu"):  cands.add(b[:-2])
    if low.endswith("iu"):   cands.add(b[:-2])
    if low.endswith("ie"):   cands.update([b[:-2], b[:-2]+"ia"])
    if low.endswith("y"):    cands.add(b[:-1] + "a")
    if low.endswith("u"):    cands.update([b[:-1], b[:-1]+"ów"])
    if low.endswith("ach"):  cands.update([b[:-3]+"e", b[:-3]+"y"])
    return [x.strip(" .") for x in cands if x.strip()]

def to_nominative(name_loc: str, teryt_df: Optional[pd.DataFrame]) -> str:
    if not name_loc:
        return ""
    base = clean_text(name_loc)
    v = IRREG_LOC2NOM.get(base.lower())
    if v:
        return v
    if teryt_df is not None and not teryt_df.empty:
        mask = teryt_df["Miejscowość"].str.lower() == base.lower()
        if mask.any():
            return teryt_df.loc[mask, "Miejscowość"].iloc[0]
    cands = generate_candidates(base)
    if teryt_df is not None and not teryt_df.empty:
        towns = teryt_df["Miejscowość"].astype(str).tolist()
        towns_norm = [norm_key(t) for t in towns]
        for cn in [norm_key(c) for c in cands]:
            if cn in towns_norm:
                return towns[towns_norm.index(cn)]
        target = norm_key(base)
        best_name, best_ratio = None, 0.0
        for name in towns:
            ratio = difflib.SequenceMatcher(None, target, norm_key(name)).ratio()
            if ratio > best_ratio:
                best_ratio, best_name = ratio, name
        if best_ratio >= FUZZY_THRESHOLD:
            return best_name
    return cands[0] if cands else base

def fill_from_teryt(miejsc: str, teryt_df: Optional[pd.DataFrame],
                    pow_hint: str = "", gmi_hint: str = "", woj_hint: str = "") -> Dict[str,str]:
    # FIX 1: poprawny warunek wstępny
    if teryt_df is None or teryt_df.empty or not miejsc:
        return {"Województwo":"","Powiat":"","Gmina":"","Miejscowość":miejsc or ""}
    sub = teryt_df[teryt_df["Miejscowość"].str.lower() == miejsc.lower()]
    if sub.empty:
        return {"Województwo":"","Powiat":"","Gmina":"","Miejscowość":miejsc}

    def pick_unique(d: pd.DataFrame) -> pd.Series:
        if len(d) == 1:
            return d.iloc[0]
        # FIX 2: wektoryzowane porównania .map(norm_key)
        if pow_hint:
            mask = d["Powiat"].map(norm_key) == norm_key(pow_hint)
            d2 = d[mask]
            if len(d2) == 1: return d2.iloc[0]
            if not d2.empty: d = d2
        if gmi_hint:
            mask = d["Gmina"].map(norm_key) == norm_key(gmi_hint)
            d2 = d[mask]
            if len(d2) == 1: return d2.iloc[0]
            if not d2.empty: d = d2
        if woj_hint:
            mask = d["Województwo"].map(norm_key) == norm_key(woj_hint)
            d2 = d[mask]
            if len(d2) == 1: return d2.iloc[0]
            if not d2.empty: d = d2
        return d.sort_values(["Województwo","Powiat","Gmina","Miejscowość"]).iloc[0]

    rec = pick_unique(sub)
    return {"Województwo": rec["Województwo"],
            "Powiat":      rec["Powiat"],
            "Gmina":       rec["Gmina"],
            "Miejscowość": rec["Miejscowość"]}

def enrich_from_teryt(df_codes: pd.DataFrame, teryt_df: Optional[pd.DataFrame]) -> pd.DataFrame:
    out_rows = []
    # FIX 3: iterrows — pewne odczytywanie kolumn z PL znakami
    for _, r in df_codes.iterrows():
        woj_hint = clean_text(r.get("Województwo", ""))
        pow_hint = clean_text(r.get("Powiat", ""))
        gmi_hint = clean_text(r.get("Gmina", ""))
        miejsc_loc = clean_text(r.get("Miejscowość", ""))

        miejsc_nom = to_nominative(miejsc_loc, teryt_df)
        canon = fill_from_teryt(miejsc_nom, teryt_df, pow_hint=pow_hint, gmi_hint=gmi_hint, woj_hint=woj_hint)

        row = {c: r.get(c, "") for c in df_codes.columns}
        for c in ("Województwo","Powiat","Gmina","Miejscowość"):
            if c in row:
                row[c] = canon.get(c, row[c])
        if not row.get("Miejscowość"):
            row["Miejscowość"] = miejsc_nom
        out_rows.append(row)

    return pd.DataFrame(out_rows, columns=df_codes.columns)

def load_input_table(path: str) -> pd.DataFrame:
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Nie znaleziono pliku wejściowego: {p.resolve()}")
    if p.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(p, dtype=str).fillna("")
    errors = []
    for enc in ("utf-8-sig", "utf-8", "cp1250"):
        try:
            return pd.read_csv(p, dtype=str, sep=None, engine="python", encoding=enc).fillna("")
        except Exception as e_auto:
            errors.append(f"{enc}/auto: {e_auto}")
            for sep in (";", ","):
                try:
                    return pd.read_csv(p, dtype=str, sep=sep, encoding=enc).fillna("")
                except Exception as e_sep:
                    errors.append(f"{enc}/{sep}: {e_sep}")
    raise RuntimeError("Nie udało się wczytać CSV. Spróbuj zapisać plik w UTF-8 lub jako .xlsx.")

def build_from_eli(out_path: str, teryt_path: Optional[str]) -> None:
    html = fetch_eli_html()
    df = parse_codes_from_table(html)
    teryt = load_teryt_df(Path(teryt_path)) if teryt_path else None
    df["Województwo"] = ""
    df["Powiat"] = ""
    df["Gmina"] = ""
    df = df[["KW_PREFIX","Województwo","Powiat","Gmina","Miejscowość"]]
    df = enrich_from_teryt(df, teryt)
    df.sort_values("KW_PREFIX", inplace=True)
    df.to_excel(out_path, index=False)

def process_from_file(in_path: str, teryt_path: Optional[str], out_path: str) -> None:
    df = load_input_table(in_path)
    need = ["KW_PREFIX","Województwo","Powiat","Gmina","Miejscowość"]
    for c in need:
        if c not in df.columns:
            raise ValueError(f"Brak kolumny '{c}' w {in_path}")
    teryt = load_teryt_df(Path(teryt_path)) if teryt_path else None
    df = df[need].fillna("")
    df = enrich_from_teryt(df, teryt)
    df.sort_values("KW_PREFIX", inplace=True)
    df.to_excel(out_path, index=False)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--mode", choices=["file","eli"], default="file")
    ap.add_argument("--in", dest="input_path", default="KW_prefix_map.xlsx",
                    help="Wejście dla trybu file: CSV/XLSX z kolumnami KW_PREFIX,Województwo,Powiat,Gmina,Miejscowość")
    ap.add_argument("--teryt", dest="teryt_path", default="TERYT.xlsx",
                    help="Ścieżka do TERYT.xlsx (arkusz z kolumnami Województwo,Powiat,Gmina,Miejscowość,Dzielnica)")
    ap.add_argument("--out", dest="out_path", default="Nr KW.xlsx",
                    help="Plik wyjściowy .xlsx")
    args = ap.parse_args()

    if args.mode == "eli":
        build_from_eli(args.out_path, args.teryt_path if Path(args.teryt_path).exists() else None)
    else:
        process_from_file(args.input_path, args.teryt_path if Path(args.teryt_path).exists() else None, args.out_path)

if __name__ == "__main__":
    main()
