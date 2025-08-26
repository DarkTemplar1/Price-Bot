#!/usr/bin/env python3
# -*- coding: utf-8 -*-
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

# 16 województw (etykieta do UI -> slug do URL/plików)
VOIVODESHIPS: list[tuple[str, str]] = [
    ("Dolnośląskie", "dolnoslaskie"),
    ("Kujawsko-Pomorskie", "kujawsko-pomorskie"),
    ("Lubelskie", "lubelskie"),
    ("Lubuskie", "lubuskie"),
    ("Łódzkie", "lodzkie"),
    ("Małopolskie", "malopolskie"),
    ("Mazowieckie", "mazowieckie"),
    ("Opolskie", "opolskie"),
    ("Podkarpackie", "podkarpackie"),
    ("Podlaskie", "podlaskie"),
    ("Pomorskie", "pomorskie"),
    ("Śląskie", "slaskie"),
    ("Świętokrzyskie", "swietokrzyskie"),
    ("Warmińsko-Mazurskie", "warminsko-mazurskie"),
    ("Wielkopolskie", "wielkopolskie"),
    ("Zachodniopomorskie", "zachodniopomorskie"),
]

PYTHON = sys.executable


def _resolve_script(script_name: str) -> Path | None:
    base = Path(__file__).resolve().parent
    p = (base / script_name).resolve()
    return p if p.exists() else None


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("680x430")
        self.resizable(False, False)

        self.var_status = tk.StringVar(value="Wybierz kategorię i województwo, aby rozpocząć.")
        self.var_phase1 = tk.StringVar(value="Etap 1 (linki): —")
        self.var_phase2 = tk.StringVar(value="Etap 2 (scraper): —")
        self._running_thread: threading.Thread | None = None

        # województwo
        self.var_region_label = tk.StringVar(value="Mazowieckie")
        self.var_region_slug = tk.StringVar(value="mazowieckie")

        self._build_ui()

    def _build_ui(self):
        wrapper = ttk.Frame(self, padding=16)
        wrapper.pack(fill="both", expand=True)

        ttk.Label(wrapper, text="Co chcesz pobrać?", font=("Segoe UI", 12, "bold")).pack(anchor="w")

        # wybór województwa
        region_box = ttk.Frame(wrapper)
        region_box.pack(fill="x", pady=(10, 0))
        ttk.Label(region_box, text="Województwo:").pack(side="left")

        combo = ttk.Combobox(
            region_box,
            textvariable=self.var_region_label,
            values=[label for label, _ in VOIVODESHIPS],
            state="readonly",
            width=28,
        )
        combo.pack(side="left", padx=8)

        def _on_region_change(*_):
            label = self.var_region_label.get()
            slug = dict(VOIVODESHIPS).get(label, "mazowieckie")
            self.var_region_slug.set(slug)

        combo.bind("<<ComboboxSelected>>", _on_region_change)
        _on_region_change()

        # przyciski kategorii
        btns = ttk.Frame(wrapper)
        btns.pack(anchor="w", pady=(10, 4))
        for label, slug in CATEGORIES:
            ttk.Button(
                btns, text=label, width=22,
                command=lambda s=slug, l=label: self._start_pipeline(l, s)
            ).pack(side="left", padx=6, pady=6)

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

        region_label = self.var_region_label.get()
        region_slug = self.var_region_slug.get()

        self.var_status.set(f"Uruchamiam proces dla: {label} — woj. {region_label}")
        self.var_phase1.set("Etap 1 (linki): uruchamianie…")
        self.var_phase2.set("Etap 2 (scraper): oczekiwanie…")

        def worker():
            # 1) LINKI
            ok1, msg1 = self._run_script_blocking(linki_path, args=["--region", region_slug])
            self._set_after(lambda: self.var_phase1.set(f"Etap 1 (linki): {'wykonano ✅' if ok1 else 'błąd ❌'}{msg1}"))

            if not ok1:
                self._set_after(lambda: self.var_status.set("Zatrzymano z powodu błędu w etapie 1."))
                return

            # 2) SCRAPER
            self._set_after(lambda: self.var_phase2.set("Etap 2 (scraper): uruchamianie…"))
            ok2, msg2 = self._run_script_blocking(scraper_path, args=["--region", region_slug])
            self._set_after(lambda: self.var_phase2.set(f"Etap 2 (scraper): {'wykonano ✅' if ok2 else 'błąd ❌'}{msg2}"))

            if ok2:
                self._set_after(lambda: self.var_status.set(f"Zakończono proces dla: {label} — woj. {region_label}"))
            else:
                self._set_after(lambda: self.var_status.set("Zakończono z błędami (etap 2)."))

        t = threading.Thread(target=worker, daemon=True)
        self._running_thread = t
        t.start()

    def _run_script_blocking(self, path: Path, args: list[str] | None = None) -> tuple[bool, str]:
        try:
            cmd = [PYTHON, str(path)] + (args or [])
            proc = subprocess.run(cmd, capture_output=False, timeout=None)
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
