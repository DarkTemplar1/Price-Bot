from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook

RAPORT_SHEET = "raport"
RAPORT_ODF = "raport_odfiltrowane"

REQ_COLS = [
    "Nr KW", "Typ Księgi", "Stan Księgi", "Województwo", "Powiat", "Gmina",
    "Miejscowość", "Dzielnica", "Położenie", "Nr działek po średniku", "Obręb po średniku",
    "Ulica", "Sposób korzystania", "Obszar", "Ulica(dla budynku)",
    "przeznaczenie (dla budynku)", "Ulica(dla lokalu)", "Nr budynku( dla lokalu)",
    "Przeznaczenie (dla lokalu)", "Cały adres (dla lokalu)",
    "Czy udziały?"
]

EXTRA_VAL_COLS = [
    "Średnia cena za m2 ( z bazy)",
    "Średnia skorygowana cena za m2",
    "Statyczna wartość nieruchomości"
]

ADDR_COLS = ["Województwo", "Powiat", "Gmina", "Miejscowość", "Dzielnica"]

OLD_TO_NEW_VOIV = {
    "białostockie": "podlaskie", "bielskie": "śląskie", "bydgoskie": "kujawsko-pomorskie",
    "chełmskie": "lubelskie", "ciechanowskie": "mazowieckie", "częstochowskie": "śląskie",
    "elbląskie": "warmińsko-mazurskie", "gdańskie": "pomorskie", "gorzowskie": "lubuskie",
    "jeleniogórskie": "dolnośląskie", "kaliskie": "wielkopolskie", "katowickie": "śląskie",
    "kieleckie": "świętokrzyskie", "koszalińskie": "zachodniopomorskie", "krakowskie": "małopolskie",
    "krośnieńskie": "podkarpackie", "legnickie": "dolnośląskie", "leszczyńskie": "wielkopolskie",
    "łomżyńskie": "podlaskie", "nowosądeckie": "małopolskie", "olsztyńskie": "warmińsko-mazurskie",
    "opole": "opolskie", "ostrołęckie": "mazowieckie", "piotrkowskie": "łódzkie",
    "płockie": "mazowieckie", "poznańskie": "wielkopolskie", "przemyskie": "podkarpackie",
    "radomskie": "mazowieckie", "rzeszowskie": "podkarpackie", "siedleckie": "mazowieckie",
    "sieradzkie": "łódzkie", "skierniewickie": "łódzkie", "słupskie": "pomorskie",
    "suwałskie": "podlaskie", "szczecińskie": "zachodniopomorskie", "tarnobrzeskie": "podkarpackie",
    "tarnowskie": "małopolskie", "toruńskie": "kujawsko-pomorskie", "wałbrzyskie": "dolnośląskie",
    "warszawskie": "mazowieckie", "włocławskie": "kujawsko-pomorskie", "wrocławskie": "dolnośląskie",
    "zamojsko-chełmskie": "lubelskie", "zielonogórskie": "lubuskie"
}

CITY_TO_VOIV = {
    "warszawa": "mazowieckie", "kraków": "małopolskie", "łódź": "łódzkie", "wrocław": "dolnośląskie",
    "poznań": "wielkopolskie", "gdańsk": "pomorskie", "szczecin": "zachodniopomorskie",
    "katowice": "śląskie", "lublin": "lubelskie", "białystok": "podlaskie",
    "rzeszów": "podkarpackie", "kielce": "świętokrzyskie", "olsztyn": "warmińsko-mazurskie",
    "opole": "opolskie", "toruń": "kujawsko-pomorskie", "bydgoszcz": "kujawsko-pomorskie",
    "gorzów wielkopolski": "lubuskie", "zielona góra": "lubuskie"
}


def get_header_map(ws):
    headers = {}
    for idx, cell in enumerate(ws[1], start=1):
        if cell.value:
            headers[str(cell.value).strip()] = idx
    return headers


def normalize(s):
    return "" if s is None else str(s).strip()


def choose_file(title, patterns):
    root = tk.Tk();
    root.withdraw()
    f = filedialog.askopenfilename(title=title, filetypes=[("Excel", "*.xlsx")])
    root.destroy()
    return f


def load_teryt(teryt_path):
    wb = load_workbook(teryt_path, data_only=True)
    ws = wb.active
    h = get_header_map(ws)
    required = ["Województwo", "Powiat", "Gmina", "Miejscowość", "Dzielnica"]
    for r in required:
        if r not in h:
            raise RuntimeError("Plik TERYT.xlsx musi mieć kolumny: " + ", ".join(required))
    data = []
    for r in range(2, ws.max_row + 1):
        row = {k: normalize(ws.cell(row=r, column=h[k]).value) for k in required}
        data.append(row)
    return data


def load_kw_map(kw_map_path):
    wb = load_workbook(kw_map_path, data_only=True)
    ws = wb.active
    h = get_header_map(ws)
    if "KW_PREFIX" not in h:
        raise RuntimeError("Plik KW_prefix_map.xlsx musi zawierać kolumnę 'KW_PREFIX'")
    cols = ["Województwo", "Powiat", "Gmina", "Miejscowość"]
    data = {}
    for r in range(2, ws.max_row + 1):
        pref = normalize(ws.cell(row=r, column=h["KW_PREFIX"]).value).upper()
        if not pref:
            continue
        rec = {}
        for c in cols:
            if c in h:
                rec[c] = normalize(ws.cell(row=r, column=h[c]).value)
        data[pref] = rec
    return data


def intersect_candidates(cands, kw_rec):
    if not kw_rec:
        return cands
    out = []
    for c in cands:
        ok = True
        for k, v in kw_rec.items():
            if v and c.get(k, "") and c.get(k, "").lower() != v.lower():
                ok = False;
                break
        if ok:
            out.append(c)
    return out


def fill_from_candidates(ws, r, hmap, candidates):
    stable = {}
    keys = ["Województwo", "Powiat", "Gmina", "Miejscowość", "Dzielnica"]
    for k in keys:
        vals = {c.get(k, "").lower() for c in candidates if c.get(k)}
        if len(vals) == 1:
            stable[k] = next(iter(vals))
    for k, v in stable.items():
        if not normalize(ws.cell(row=r, column=hmap[k]).value):
            ws.cell(row=r, column=hmap[k], value=v)


def popraw():
    root = tk.Tk();
    root.withdraw()
    path = filedialog.askopenfilename(title="Wybierz raport (.xlsx)", filetypes=[("Excel", "*.xlsx")])
    root.destroy()
    if not path:
        return

    wb = load_workbook(path)
    if RAPORT_SHEET not in wb.sheetnames:
        messagebox.showerror("Błąd", f"Brak arkusza '{RAPORT_SHEET}'")
        return
    ws = wb[RAPORT_SHEET]
    hmap = get_header_map(ws)

    for needed in ADDR_COLS + EXTRA_VAL_COLS + ["Nr KW"]:
        if needed not in hmap:
            messagebox.showerror("Błąd", f"Brak wymaganej kolumny: {needed}")
            return

    teryt_path = None
    kw_map_path = None
    if Path("TERYT.xlsx").exists():
        teryt_path = "TERYT.xlsx"
    if Path("KW_prefix_map.xlsx").exists():
        kw_map_path = "KW_prefix_map.xlsx"

    if not teryt_path:
        teryt_path = choose_file("Wskaż TERYT.xlsx (Województwo,Powiat,Gmina,Miejscowość,Dzielnica)",
                                 ["*.xlsx"]) or None
    if not kw_map_path:
        kw_map_path = choose_file("Wskaż KW_prefix_map.xlsx (KW_PREFIX, ...)", ["*.xlsx"]) or None

    teryt_data = []
    kw_map = {}
    if teryt_path and Path(teryt_path).exists():
        try:
            teryt_data = load_teryt(teryt_path)
        except Exception as e:
            messagebox.showwarning("TERYT", f"Nie udało się załadować TERYT: {e}")
    if kw_map_path and Path(kw_map_path).exists():
        try:
            kw_map = load_kw_map(kw_map_path)
        except Exception as e:
            messagebox.showwarning("KW map", f"Nie udało się załadować mapy KW: {e}")

    idxs = [hmap[c] for c in ADDR_COLS]
    idx_kw = hmap["Nr KW"]
    idx_avg, idx_corr, idx_static = [hmap[c] for c in EXTRA_VAL_COLS]

    for r in range(2, ws.max_row + 1):
        cells = [normalize(ws.cell(row=r, column=i).value) for i in idxs]
        if all(c == "" for c in cells):
            ws.cell(row=r, column=idx_avg, value="Brak adresu")
            ws.cell(row=r, column=idx_corr, value="Brak adresu")
            ws.cell(row=r, column=idx_static, value="Brak adresu")

    woj_idx = hmap["Województwo"]
    for r in range(2, ws.max_row + 1):
        v = normalize(ws.cell(row=r, column=woj_idx).value).lower()
        if v in OLD_TO_NEW_VOIV:
            ws.cell(row=r, column=woj_idx, value=OLD_TO_NEW_VOIV[v])

    for r in range(2, ws.max_row + 1):
        values = {c: normalize(ws.cell(row=r, column=hmap[c]).value) for c in ADDR_COLS}
        if all(values[c] for c in ADDR_COLS):
            continue

        kw = normalize(ws.cell(row=r, column=idx_kw).value).upper()
        kw_pref = "".join([ch for ch in kw if ch.isalnum()])[:4] if kw else ""
        kw_rec = kw_map.get(kw_pref, {}) if kw_pref else {}

        if values.get("Miejscowość") and not values.get("Województwo"):
            mv = values["Miejscowość"].lower()
            if mv in CITY_TO_VOIV:
                ws.cell(row=r, column=hmap["Województwo"], value=CITY_TO_VOIV[mv])
                values["Województwo"] = CITY_TO_VOIV[mv]

        if teryt_data:
            cands = teryt_data
            for key in ADDR_COLS:
                if values.get(key):
                    cands = [x for x in cands if x.get(key, "").lower() == values[key].lower()]
            cands = intersect_candidates(cands, kw_rec)
            if len(cands) == 1:
                row = cands[0]
                for k in ADDR_COLS:
                    if not values.get(k):
                        ws.cell(row=r, column=hmap[k], value=row.get(k))
            elif len(cands) > 1:
                fill_from_candidates(ws, r, hmap, cands)

    wb.save(path)
    messagebox.showinfo("Zakończono", "Adresy zostały poprawione.")


if __name__ == "__main__":
    popraw()
