#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations  # <-- musi być tu, na górze pliku, i tylko raz

import sys
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

APP_TITLE = "Sortownia – uruchamianie filtrów"

# (etykieta, plik skryptu)
FILTER_OPTIONS = [
    ("Czy udział? = nie", "jeden_właściciel.py"),
    ("Lokal mieszkalny", "LOKAL_MIESZKALNY.py"),
    ("Lokal mieszkalny", "jeden_właściciel_i_LOKAL_MIESZKALNY.py"),
]

def _script_path(script_name: str) -> Path:
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
    def __init__(self, window: tk.Toplevel, start_path: str | None = None):
        self.root = window
        self.root.title(APP_TITLE)
        self.root.minsize(680, 300)

        self.excel_path: Path | None = Path(start_path).expanduser() if start_path else None
        self.var_selected = tk.StringVar(value=FILTER_OPTIONS[0][0])

        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._back_to_main)

    def _back_to_main(self):
        self.root.destroy()

    def _build_ui(self):
        header = ttk.Frame(self.root, padding=12)
        header.pack(fill="x")
        ttk.Button(header, text="⟵ Zamknij", command=self._back_to_main).pack(side="left")
        ttk.Button(header, text="Wybierz plik…", command=self._choose_file).pack(side="left", padx=(8, 8))
        ttk.Label(header, text="Plik:").pack(side="left")

        self.path_var = tk.StringVar(value=str(self.excel_path) if self.excel_path else "(nie wybrano)")
        ttk.Label(header, textvariable=self.path_var, width=70).pack(side="left", padx=8)

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

        ttk.Button(box, text="Aktywuj filtr", command=self._activate).pack(side="left", padx=(8, 0))
        ttk.Button(box, text="poprawa_adresu", command=self._run_poprawa_adresu).pack(side="left", padx=(12, 0))
        ttk.Button(box, text="cofnij filtry", command=self._run_cofnij).pack(side="left", padx=(8, 0))

        status = ttk.Frame(self.root, padding=12)
        status.pack(fill="x", side="bottom")
        self.msg = tk.StringVar(value="Gotowy.")
        ttk.Label(status, textvariable=self.msg).pack(anchor="w")

    def _choose_file(self):
        fp = filedialog.askopenfilename(
            title="Wskaż plik Excela",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("Wszystkie pliki", "*.*")],
        )
        if fp:
            self.excel_path = Path(fp)
            self.path_var.set(_nice_path(self.excel_path))

    def _activate(self):
        idx = self.cmb.current()
        _, script = FILTER_OPTIONS[idx]
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

        args = [sys.executable, str(spath)]
        if self.excel_path:
            args.append(str(self.excel_path))

        try:
            subprocess.Popen(args)
            self.msg.set(f"Uruchomiono: {spath.name}")
        except Exception as e:
            messagebox.showerror("Błąd uruchamiania", f"Nie udało się uruchomić skryptu:\n{spath}\n\n{e}")

# ====== API dla main.py / standalone ======
def run_modal(parent: tk.Tk, file_path: str | None = None):
    """
    Otwiera sortownię jako MODALNE okno potomne (Toplevel) bez dodatkowego mainloop.
    """
    win = tk.Toplevel(parent)
    win.transient(parent)

    # wycentruj względem parenta
    win.update_idletasks()
    try:
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        ww, wh = 700, 320
        x = px + max(0, (pw - ww) // 2)
        y = py + max(0, (ph - wh) // 2)
        win.geometry(f"{ww}x{wh}+{x}+{y}")
    except Exception:
        win.geometry("700x320")

    SortowniaApp(win, start_path=file_path)

    # modalność + fokus
    win.grab_set()
    try:
        win.attributes("-topmost", True)
        win.lift()
        win.focus_force()
        win.after(200, lambda: win.attributes("-topmost", False))
    except tk.TclError:
        pass

    parent.wait_window(win)

if __name__ == "__main__":
    # tryb standalone – można uruchomić z konsoli do testów
    root = tk.Tk()
    root.withdraw()
    run_modal(root, sys.argv[1] if len(sys.argv) > 1 else None)
    root.destroy()
