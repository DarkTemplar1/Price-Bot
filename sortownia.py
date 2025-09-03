#!/usr/bin/env python3
# -*- coding: utf-8 -*-


from __future__ import annotations
import sys
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


APP_TITLE = "Sortownia – uruchamianie filtrów"

# Opcje w comboboxie → odpowiadające im pliki skryptów
FILTER_OPTIONS = [
    # (etykieta widoczna w comboboxie, nazwa pliku skryptu do uruchomienia)
    ("Filtr 1 — przykład", "filtr_1.py"),
    ("Filtr 2 — przykład", "filtr_2.py"),
    ("Filtr 3 — przykład", "filtr_3.py"),
]

def _script_path(script_name: str) -> Path:
    """
    Znajdź plik skryptu wg kilku standardowych lokalizacji:
    - obok bieżącego pliku,
    - ścieżka bieżąca (cwd),
    - bezpośrednio jeśli podano pełną ścieżkę.
    """
    candidates = [
        Path(__file__).with_name(script_name),
        Path.cwd() / script_name,
        Path(script_name),
    ]
    for c in candidates:
        try:
            if c.exists():
                return c.resolve()
        except Exception:
            pass
    return Path(script_name)

def _nice_path(p: Path | None) -> str:
    return str(p) if p else "(nie wybrano)"

class SortowniaApp:
    def __init__(self, root: tk.Tk, start_path: str | None = None):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("700x320")
        self.root.minsize(680, 300)

        # bieżący plik Excel do przekazania w skryptach
        self.excel_path: Path | None = Path(start_path).expanduser() if start_path else None

        # UI state
        self.var_selected = tk.StringVar(value=FILTER_OPTIONS[0])

        self._build_ui()

        # Zamknięcie okna krzyżykiem -> powrót do main.py
        self.root.protocol("WM_DELETE_WINDOW", self._back_to_main)

    def _back_to_main(self):
        """Uruchom main.py i zamknij bieżące okno."""
        main_path = _script_path("main.py")
        if not main_path.exists():
            messagebox.showerror("Nie znaleziono skryptu", f"Nie mogę znaleźć pliku:\n{main_path}")
            return
        try:
            subprocess.Popen([sys.executable, str(main_path)])
        except Exception as e:
            messagebox.showerror("Błąd uruchamiania", f"Nie udało się uruchomić main.py:\n{e}")
            return
        self.root.destroy()

    def _build_ui(self):
        # Nagłówek
        header = ttk.Frame(self.root, padding=12)
        header.pack(fill="x")
        ttk.Button(header, text="⟵ Zamknij", command=self._back_to_main).pack(side="left")
        ttk.Button(header, text="Wybierz plik…", command=self._choose_file).pack(side="left", padx=(8, 8))
        ttk.Label(header, text="Plik:").pack(side="left")

        self.path_var = tk.StringVar(value=str(self.excel_path) if self.excel_path else "(nie wybrano)")
        ttk.Label(header, textvariable=self.path_var, width=70).pack(side="left", padx=8)

        # Wybór filtra
        body = ttk.Frame(self.root, padding=12)
        body.pack(fill="both", expand=True)

        ttk.Label(body, text="Wybierz filtr:").pack(anchor="w")

        box = ttk.Frame(body)
        box.pack(anchor="w", pady=(6, 0))

        self.cmb = ttk.Combobox(
            box,
            textvariable=self.var_selected,
            state="readonly",
            width=50,
            values=[label for (label, _) in FILTER_OPTIONS],
        )
        self.cmb.current(0)
        self.cmb.pack(side="left", padx=8)

        self.btn_go = ttk.Button(box, text="Aktywuj filtr", command=self._activate)
        self.btn_go.pack(side="left", padx=(8, 0))

        # DODANE: przycisk poprawa_adresu
        btn_fix = ttk.Button(box, text="poprawa_adresu", command=self._run_poprawa_adresu)
        btn_fix.pack(side="left", padx=(12, 0))

        # DODANE: przycisk cofnij filtry
        btn_undo = ttk.Button(box, text="cofnij filtry", command=self._run_cofnij)
        btn_undo.pack(side="left", padx=(8, 0))

        # Status
        status = ttk.Frame(self.root, padding=12)
        status.pack(fill="x", side="bottom")
        self.msg = tk.StringVar(value="Gotowy.")
        ttk.Label(status, textvariable=self.msg).pack(anchor="w")

    def _choose_file(self):
        fp = filedialog.askopenfilename(
            title="Wskaż plik Excela",
            filetypes=[
                ("Excel", "*.xlsx *.xlsm *.xls"),
                ("Wszystkie pliki", "*.*"),
            ],
        )
        if fp:
            self.excel_path = Path(fp)
            self.path_var.set(_nice_path(self.excel_path))

    def _activate(self):
        idx = self.cmb.current()
        label, script = FILTER_OPTIONS[idx]
        self._run_selected_script(script)

    def _run_poprawa_adresu(self):
        self._run_selected_script("poprawa_adresu.py")

    def _run_cofnij(self):
        self._run_selected_script("cofnij_filtry.py")

    def _run_selected_script(self, script_name: str):
        spath = _script_path(script_name)
        if not spath.exists():
            messagebox.showerror("Nie znaleziono skryptu", f"Nie mogę znaleźć pliku:\n{spath}")
            return
        # budujemy polecenie
        args = [sys.executable, str(spath)]
        if self.excel_path:
            args.append(str(self.excel_path))

        try:
            subprocess.Popen(args)
            self.msg.set(f"Uruchomiono: {spath.name}")
        except Exception as e:
            messagebox.showerror("Błąd uruchamiania", f"Nie udało się uruchomić skryptu:\n{spath}\n\n{e}")


# ====== API dla main.py / standalone ======
def run_modal(file_path: str | None = None):
    root = tk.Tk()
    app = SortowniaApp(root, start_path=file_path)
    root.mainloop()


if __name__ == "__main__":
    # Użycie: python sortownia.py [opcjonalnie_ścieżka_do_excela]
    start = sys.argv[1] if len(sys.argv) > 1 else None
    run_modal(start)
