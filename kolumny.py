
from openpyxl import Workbook, load_workbook

RAPORT_SHEET = "raport"

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

def ensure_sheet_and_columns(xlsx_path: str, sheet_name: str = RAPORT_SHEET):
    if not xlsx_path:
        raise ValueError("Brak ścieżki do pliku .xlsx")
    try:
        wb = load_workbook(xlsx_path)
    except Exception:
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.create_sheet(sheet_name)

    if ws.max_row == 1 and ws.max_column == 1 and ws.cell(row=1, column=1).value is None:
        headers = REQ_COLS + EXTRA_VAL_COLS
        for c, h in enumerate(headers, start=1):
            ws.cell(row=1, column=c, value=h)
    else:
        headers = [cell.value or "" for cell in ws[1]]
        def add_col_if_missing(col_name: str):
            if col_name not in headers:
                ws.cell(row=1, column=ws.max_column+1, value=col_name)
        for col in REQ_COLS + EXTRA_VAL_COLS:
            add_col_if_missing(col)

    wb.save(xlsx_path)
