#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
bazadanych.py  (UWAGA: jeśli inne pliki odwołują się do „bazydanych.py”,
zapisz ten plik pod obiema nazwami lub zaktualizuj odwołania.)

GUI z 5 przyciskami:
- mieszkania
- dom
- działki
- inwestycje
- Hale i Magazyny

Dla wybranej kategorii uruchamia sekwencyjnie:
1) linki_<kategoria>.py  → po zakończeniu pokazuje „Etap 1 wykonano”
2) scraper_otodom_<kategoria>.py

Nazwy skryptów (pliki w tym samym folderze co ten plik):
- linki_mieszkania.py
- linki_dom.py
- linki_dzialki.py
- linki_inwestycje.py
- linki_hale_i_magazyny.py

- scraper_otodom_mieszkania.py
- scraper_otodom_dom.py
- scraper_otodom_dzialki.py
- scraper_otodom_inwestycje.py
- scraper_otodom_hale_i_magazyny.py
"""
from __future__ import annotations

import sys
import threading
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox

APP_TITLE = "Baza danych – pobieranie i scraping"

# Mapa: etykieta → slug (używany w nazwach plików)
CATEGORIES: list[tuple[str, str]] = [
    ("mieszkania", "mieszkania"),
    ("dom", "dom"),
    ("działki", "dzialki"),
    ("inwestycje", "inwestycje"),
    ("Hale i Magazyny", "hale_i_magazyny"),
]

PYTHON = sys.executable


def _resolve_script(script_name: str) -> Path | None:
    """Zwraca ścieżkę do skryptu w tym samym folderze co ten plik, jeśli istnieje."""
    base = Path(__file__).resolve().parent
    p = (base / script_name).resolve()
    return p if p.exists() else None


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("620x380")
        self.resizable(False, False)

        self.var_status = tk.StringVar(value="Wybierz kategorię, aby rozpocząć.")
        self.var_phase1 = tk.StringVar(value="Etap 1 (linki): —")
        self.var_phase2 = tk.StringVar(value="Etap 2 (scraper): —")
        self._running_thread: threading.Thread | None = None

        self._build_ui()

    def _build_ui(self):
        wrapper = ttk.Frame(self, padding=16)
        wrapper.pack(fill="both", expand=True)

        ttk.Label(wrapper, text="Co chcesz pobrać?", font=("Segoe UI", 12, "bold")).pack(anchor="w")

        btns = ttk.Frame(wrapper)
        btns.pack(anchor="w", pady=(10, 4))

        for label, slug in CATEGORIES:
            b = ttk.Button(btns, text=label, width=22, command=lambda s=slug, l=label: self._start_pipeline(l, s))
            b.pack(side="left", padx=6, pady=6)

        sep = ttk.Separator(wrapper, orient="horizontal")
        sep.pack(fill="x", pady=(10, 8))

        info = ttk.LabelFrame(wrapper, text="Status", padding=10)
        info.pack(fill="x")

        ttk.Label(info, textvariable=self.var_status, justify="left").pack(anchor="w")
        ttk.Label(info, textvariable=self.var_phase1, justify="left").pack(anchor="w", pady=(6, 0))
        ttk.Label(info, textvariable=self.var_phase2, justify="left").pack(anchor="w")

        controls = ttk.Frame(wrapper)
        controls.pack(fill="x", pady=(12, 0))
        ttk.Button(controls, text="Zamknij", command=self._safe_close).pack(side="right")

    # ---------------- Pipeline ----------------

    def _start_pipeline(self, label: str, slug: str):
        if self._running_thread and self._running_thread.is_alive():
            messagebox.showinfo("W toku", "Inny proces jest w trakcie. Poczekaj na zakończenie.")
            return

        linki_name = f"linki_{slug}.py"
        scraper_name = f"scraper_otodom_{slug}.py"

        linki_path = _resolve_script(linki_name)
        scraper_path = _resolve_script(scraper_name)

        missing = [name for name, p in [(linki_name, linki_path), (scraper_name, scraper_path)] if p is None]
        if missing:
            messagebox.showerror(
                "Brak pliku",
                "Nie znaleziono następujących skryptów w katalogu aplikacji:\n- " + "\n- ".join(missing),
            )
            return

        self.var_status.set(f"Uruchamiam proces dla: {label}")
        self.var_phase1.set("Etap 1 (linki): uruchamianie…")
        self.var_phase2.set("Etap 2 (scraper): oczekiwanie…")

        def worker():
            # 1) LINKI
            ok1, msg1 = self._run_script_blocking(linki_path)
            self._set_after(lambda: self.var_phase1.set(f"Etap 1 (linki): {'wykonano ✅' if ok1 else 'błąd ❌'}{msg1}"))

            if not ok1:
                self._set_after(lambda: self.var_status.set("Zatrzymano z powodu błędu w etapie 1."))
                return

            # 2) SCRAPER
            self._set_after(lambda: self.var_phase2.set("Etap 2 (scraper): uruchamianie…"))
            ok2, msg2 = self._run_script_blocking(scraper_path)
            self._set_after(lambda: self.var_phase2.set(f"Etap 2 (scraper): {'wykonano ✅' if ok2 else 'błąd ❌'}{msg2}"))

            if ok2:
                self._set_after(lambda: self.var_status.set(f"Zakończono proces dla: {label}"))
            else:
                self._set_after(lambda: self.var_status.set("Zakończono z błędami (etap 2)."))

        t = threading.Thread(target=worker, daemon=True)
        self._running_thread = t
        t.start()

    def _run_script_blocking(self, path: Path) -> tuple[bool, str]:
        """Uruchamia skrypt Pythona i czeka na zakończenie. Zwraca (ok, ' (kod=...)' / ' (timeout)')."""
        try:
            proc = subprocess.run([PYTHON, str(path)], capture_output=False, timeout=None)
            return (proc.returncode == 0, f" (kod wyjścia={proc.returncode})")
        except FileNotFoundError:
            return (False, " (nie znaleziono pliku wykonywalnego)")
        except Exception as e:
            return (False, f" ({e})")

    def _set_after(self, fn):
        try:
            self.after(0, fn)
        except Exception:
            pass

    # ---------------- Lifecycle ----------------

    def _safe_close(self):
        if self._running_thread and self._running_thread.is_alive():
            if not messagebox.askyesno("Wciąż trwa", "Proces jeszcze działa. Na pewno zamknąć okno?"):
                return
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
