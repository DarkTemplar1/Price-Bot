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
    "jeden właściciel",
    "lokal mieszkalny",
    "Jeden właściciel i Lokal mieszkalny",
]

FILTER_TO_SCRIPT = {
    "jeden właściciel": "jeden_właściciel.py",
    "lokal mieszkalny": "LOKAL_MIESZKALNY.py",
    "Jeden właściciel i Lokal mieszkalny": "jeden_właściciel_i_LOKAL_MIESZKALNY.py",
}


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
    return candidates[0]


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

    def _build_ui(self):
        # Nagłówek
        header = ttk.Frame(self.root, padding=12)
        header.pack(fill="x")
        ttk.Button(header, text="⟵ Zamknij", command=self.root.destroy).pack(side="left")
        ttk.Button(header, text="Wybierz plik…", command=self._choose_file).pack(side="left", padx=(8, 8))
        ttk.Label(header, text="Plik:").pack(side="left")

        self.path_var = tk.StringVar(value=str(self.excel_path) if self.excel_path else "(nie wybrano)")
        ttk.Label(header, textvariable=self.path_var, width=70).pack(side="left", padx=8)

        # Wybór filtra
        box = ttk.LabelFrame(self.root, text="Filtr do uruchomienia", padding=12)
        box.pack(fill="x", padx=12, pady=(8, 0))

        ttk.Label(box, text="Wybierz:").pack(side="left")
        self.cmb = ttk.Combobox(
            box, state="readonly", width=42, textvariable=self.var_selected, values=FILTER_OPTIONS
        )
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
        status.pack(fill="both", expand=True)
        self.info_var = tk.StringVar(
            value="Wybierz plik Excela, potem filtr z listy i kliknij „Aktywuj filtr”.\n"
                  "Możesz też użyć „poprawa_adresu” (popraw_adres.py) lub „cofnij filtry” (cofnij.py)."
        )
        ttk.Label(status, textvariable=self.info_var, justify="left").pack(anchor="w")

        try:
            style = ttk.Style(self.root)
            if "clam" in style.theme_names():
                style.theme_use("clam")
        except Exception:
            pass

    # ---- Akcje UI ----
    def _choose_file(self):
        path = filedialog.askopenfilename(
            title="Wybierz plik Excel",
            filetypes=[
                ("Pliki Excel", "*.xlsx *.xls *.xlsm *.xlsb"),
                ("Wszystkie pliki", "*.*"),
            ],
        )
        if not path:
            return
        self.excel_path = Path(path)
        self.path_var.set(str(self.excel_path))

    def _activate(self):
        """Uruchom odpowiedni skrypt z przekazaniem ścieżki do Excela."""
        if self.excel_path is None:
            messagebox.showwarning("Brak pliku", "Najpierw wybierz plik Excela.")
            return

        choice = self.var_selected.get()
        script_name = FILTER_TO_SCRIPT.get(choice)
        if not script_name:
            messagebox.showerror("Błąd", f"Nieznana opcja: {choice}")
            return

        self._run_script(script_name)

    def _run_poprawa_adresu(self):
        """Uruchamia popraw_adres.py z parametrem --in <plik.xlsx>."""
        if self.excel_path is None:
            messagebox.showwarning("Brak pliku", "Najpierw wybierz plik Excela.")
            return
        self._run_script("popraw_adres.py")

    def _run_cofnij(self):
        """Uruchamia cofnij.py z parametrem --in <plik.xlsx> (przywrócenie po filtrach)."""
        if self.excel_path is None:
            messagebox.showwarning("Brak pliku", "Najpierw wybierz plik Excela.")
            return
        self._run_script("cofnij.py")

    def _run_script(self, script_name: str):
        """Wspólna logika uruchamiania zewnętrznego skryptu w nowym procesie."""
        spath = _script_path(script_name)
        if not spath.exists():
            messagebox.showerror("Nie znaleziono skryptu", f"Nie mogę znaleźć pliku:\n{spath}")
            return
        cmd = [sys.executable, str(spath), "--in", str(self.excel_path)]
        try:
            subprocess.Popen(cmd)
            self.info_var.set(f"Uruchomiono: {spath.name}  (plik: {self.excel_path})")
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
