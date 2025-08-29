# -*- coding: utf-8 -*-
import sys
from pathlib import Path
import pandas as pd
from tkinter import Tk, filedialog, messagebox

OUT_XLSX = "TERYT.xlsx"

def pick_file(title, patterns=(("CSV files","*.csv"),("All","*.*"))):
    Tk().withdraw()
    f = filedialog.askopenfilename(title=title, filetypes=patterns)
    return f

def read_csv_any(path):
    for enc in ("utf-8", "cp1250", "iso-8859-2"):
        for sep in (";", ",", "\t"):
            try:
                df = pd.read_csv(path, encoding=enc, sep=sep, dtype=str, keep_default_na=False)
                return df
            except Exception:
                pass
    raise RuntimeError(f"Nie mogę wczytać CSV: {path} (spróbuj zapisać jako UTF-8; separator ';')")

def zfill(s, n):
    s = "" if s is None else str(s).strip()
    return s.zfill(n) if s else s

def build_lookups_terc(df_terc):
    df = df_terc.copy()
    req = ["WOJ","POW","GMI","RODZ","NAZWA"]
    miss = [c for c in req if c not in df.columns]
    if miss:
        raise RuntimeError("TERC.csv musi zawierać kolumny: " + ", ".join(req))

    df["WOJ"] = df["WOJ"].map(lambda s: zfill(s,2))
    df["POW"] = df["POW"].map(lambda s: zfill(s,2))
    df["GMI"] = df["GMI"].map(lambda s: zfill(s,3))
    df["RODZ"] = df["RODZ"].map(lambda s: zfill(s,1))

    woj = df[(df["POW"]=="") & (df["GMI"]=="")][["WOJ","NAZWA"]].drop_duplicates()
    woj_map = dict(zip(woj["WOJ"], woj["NAZWA"]))

    powiat = df[(df["GMI"]=="") & (df["POW"]!="")][["WOJ","POW","NAZWA"]].drop_duplicates()
    pow_map = {(r.WOJ, r.POW): r.NAZWA for r in powiat.itertuples()}

    gmina = df[(df["GMI"]!="")][["WOJ","POW","GMI","RODZ","NAZWA"]].drop_duplicates()
    gmi_map = {(r.WOJ, r.POW, r.GMI, r.RODZ): r.NAZWA for r in gmina.itertuples()}
    return woj_map, pow_map, gmi_map

def detect_simc_cols(df_simc):
    simc = df_simc.copy()
    cols = {c.lower(): c for c in simc.columns}
    def pick(*names):
        for n in names:
            if n.lower() in cols:
                return cols[n.lower()]
        return None
    c_WOJ = pick("WOJ")
    c_POW = pick("POW")
    c_GMI = pick("GMI")
    c_RODZ_GMI = pick("RODZ_GMI","RODZ","RODZAJ_GMINY")
    c_SYM = pick("SYM","SYM_SIMC","ID_SIMC")
    c_SYM_POD = pick("SYM_POD","SYMPOD","SYM_PARENT","SYM_NADRZ")
    c_NAZWA = pick("NAZWA","NAZWA_MIEJSCOWOSCI","NAZWA_SIMC","NAZWA_SIMC_PL")
    need = [c_WOJ,c_POW,c_GMI,c_RODZ_GMI,c_SYM,c_NAZWA]
    if any(v is None for v in need):
        raise RuntimeError("SIMC.csv: brak wymaganych kolumn (WOJ, POW, GMI, RODZ_GMI, SYM, NAZWA).")
    return {
        "WOJ": c_WOJ, "POW": c_POW, "GMI": c_GMI, "RODZ_GMI": c_RODZ_GMI,
        "SYM": c_SYM, "SYM_POD": c_SYM_POD, "NAZWA": c_NAZWA
    }

def build_teryt_xlsx(terc_csv=None, simc_csv=None, out_xlsx=OUT_XLSX):
    if not terc_csv:
        terc_csv = pick_file("Wskaż pełny TERC.csv")
    if not simc_csv:
        simc_csv = pick_file("Wskaż pełny SIMC.csv")
    if not terc_csv or not simc_csv:
        raise SystemExit(0)

    df_terc = read_csv_any(terc_csv)
    df_simc = read_csv_any(simc_csv)

    woj_map, pow_map, gmi_map = build_lookups_terc(df_terc)

    c = detect_simc_cols(df_simc)
    simc = df_simc.copy()
    simc["WOJ"] = simc[c["WOJ"]].map(lambda s: zfill(s,2))
    simc["POW"] = simc[c["POW"]].map(lambda s: zfill(s,2))
    simc["GMI"] = simc[c["GMI"]].map(lambda s: zfill(s,3))
    simc["RODZ_GMI"] = simc[c["RODZ_GMI"]].map(lambda s: zfill(s,1))
    simc["SYM"] = simc[c["SYM"]].map(lambda s: zfill(s,7))
    simc["SYM_POD"] = simc[c["SYM_POD"]].map(lambda s: zfill(s,7)) if c["SYM_POD"] else ""
    simc["NAZWA"] = simc[c["NAZWA"]].astype(str).map(lambda s: s.strip())

    sym_to_name = dict(zip(simc["SYM"], simc["NAZWA"]))

    rows = []
    for r in simc.itertuples():
        woj = woj_map.get(r.WOJ, "")
        powiat = pow_map.get((r.WOJ, r.POW), "")
        gmina = gmi_map.get((r.WOJ, r.POW, r.GMI, r.RODZ_GMI), "")
        if r.SYM_POD and r.SYM_POD != "":
            parent = sym_to_name.get(r.SYM_POD, "")
            miejsc = parent if parent else r.NAZWA
            dziel = r.NAZWA if parent else ""
        else:
            miejsc = r.NAZWA
            dziel = ""
        rows.append([woj, powiat, gmina, miejsc, dziel])

    out_df = pd.DataFrame(rows, columns=["Województwo","Powiat","Gmina","Miejscowość","Dzielnica"]).drop_duplicates()
    for c in out_df.columns:
        out_df[c] = out_df[c].fillna("").map(lambda s: str(s).strip())

    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as xw:
        out_df.to_excel(xw, sheet_name="TERYT", index=False)

    return out_xlsx, len(out_df)

if __name__ == "__main__":
    try:
        out, n = build_teryt_xlsx()
        try:
            Tk().withdraw()
            messagebox.showinfo("OK", f"Wygenerowano {out} (wierszy: {n}).")
        except Exception:
            print(f"Wygenerowano {out} (wierszy: {n}).")
    except Exception as e:
        try:
            Tk().withdraw()
            messagebox.showerror("Błąd", str(e))
        except Exception:
            print("BŁĄD:", e, file=sys.stderr)
            sys.exit(1)
