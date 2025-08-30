# -*- coding: utf-8 -*-
import sys
from pathlib import Path
from typing import List
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import Workbook, load_workbook

RAPORT_SHEET = "raport"
RAPORT_ODF = "raport_odfiltrowane"

REQ_COLS: List[str] = [
    "Nr KW","Typ Księgi","Stan Księgi","Województwo","Powiat","Gmina",
    "Miejscowość","Dzielnica","Położenie","Nr działek po średniku","Obręb po średniku",
    "Ulica","Sposób korzystania","Obszar","Ulica(dla budynku)",
    "przeznaczenie (dla budynku)","Ulica(dla lokalu)","Nr budynku( dla lokalu)",
    "Przeznaczenie (dla lokalu)","Cały adres (dla lokalu)","Czy udziały?"
]

VALUE_COLS: List[str] = [
    "Średnia cena za m2 ( z bazy)",
    "Średnia skorygowana cena za m2",
    "Statyczna wartość nieruchomości"
]

SUPPORTED = {".xlsx", ".xlsm"}

def _ensure_base_headers(ws) -> None:
    # jeżeli pusty arkusz — wpisz od razu bazowe REQ_COLS
    if ws.max_row == 1 and ws.max_column == 1 and (ws.cell(1,1).value in (None, "")):
        for c, h in enumerate(REQ_COLS, start=1):
            ws.cell(row=1, column=c, value=h)
        return

    # dopisz brakujące bazowe nagłówki na końcu (bez VALUE_COLS)
    existing = [cell.value or "" for cell in ws[1]]
    for col in REQ_COLS:
        if col not in existing:
            ws.cell(row=1, column=ws.max_column + 1, value=col)
            existing.append(col)

def _ensure_value_cols_after_anchor(ws, anchor="Czy udziały?") -> None:
    # odczytaj nagłówki
    headers = [cell.value or "" for cell in ws[1]]
    # zapewnij, że anchor istnieje
    if anchor not in headers:
        ws.cell(row=1, column=ws.max_column + 1, value=anchor)
        headers = [cell.value or "" for cell in ws[1]]

    anchor_idx = headers.index(anchor) + 1  # 1-based
    # sprawdź czy trzy kolumny za anchor mają już właściwe nazwy
    want = VALUE_COLS
    ok = True
    for offset, name in enumerate(want, start=1):
        cell_val = ws.cell(row=1, column=anchor_idx + offset).value
        if cell_val != name:
            ok = False
            break
    if ok:
        return  # już jest dobrze

    # wstaw 3 kolumny za anchor
    ws.insert_cols(anchor_idx + 1, amount=3)
    # ustaw nagłówki nowo wstawionych kolumn
    for i, name in enumerate(want, start=1):
        ws.cell(row=1, column=anchor_idx + i, value=name)

def ensure_sheet_and_columns(xlsx_path: Path) -> None:
    try:
        wb = load_workbook(xlsx_path)
    except Exception:
        wb = Workbook()
        wb.active.title = RAPORT_SHEET

    for sheet_name in (RAPORT_SHEET, RAPORT_ODF):
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.create_sheet(sheet_name)
        _ensure_base_headers(ws)
        _ensure_value_cols_after_anchor(ws, anchor="Czy udziały?")

    wb.save(xlsx_path)

def pick_file_via_gui() -> Path:
    root = tk.Tk(); root.withdraw()
    p = filedialog.askopenfilename(title="Wybierz plik Excel",
                                   filetypes=[("Excel", "*.xlsx *.xlsm"), ("Wszystkie pliki","*.*")])
    root.destroy()
    if not p:
        raise SystemExit(0)
    return Path(p)

def info(msg: str) -> None:
    try:
        r = tk.Tk(); r.withdraw()
        messagebox.showinfo("Gotowe", msg, parent=r)
        r.destroy()
    except Exception:
        print(msg)

def error(msg: str) -> None:
    try:
        r = tk.Tk(); r.withdraw()
        messagebox.showerror("Błąd", msg, parent=r)
        r.destroy()
    except Exception:
        print("BŁĄD:", msg, file=sys.stderr)

def main():
    if len(sys.argv) >= 2:
        xlsx_path = Path(sys.argv[1])
    else:
        xlsx_path = pick_file_via_gui()

    if xlsx_path.suffix.lower() not in SUPPORTED:
        error(f"Nieobsługiwane rozszerzenie: {xlsx_path.suffix}. Zapisz plik jako .xlsx lub .xlsm.")
        raise SystemExit(2)

    try:
        ensure_sheet_and_columns(xlsx_path)
        info(f"Przygotowano arkusze '{RAPORT_SHEET}' i '{RAPORT_ODF}' oraz dodano kolumny za 'Czy udziały?' w:\n{xlsx_path}")
    except Exception as e:
        error(f"Nie udało się przygotować pliku:\n{e}")
        raise SystemExit(1)

if __name__ == "__main__":
    main()
