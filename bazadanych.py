#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
import csv
import threading
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox

import re
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

APP_TITLE = "Baza danych – pobieranie i scraping"
PYTHON = sys.executable

# =================== GUI KONFIG ===================
CATEGORIES: list[tuple[str, str]] = [
    ("mieszkania", "mieszkania"),
    ("dom", "dom"),
    ("działki", "dzialki"),
    ("inwestycje", "inwestycje"),
    ("Hale i Magazyny", "hale_i_magazyny"),
]
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

def _resolve_script(script_name: str) -> Path | None:
    base = Path(__file__).resolve().parent
    p = (base / script_name).resolve()
    return p if p.exists() else None

def _detect_desktop() -> Path:
    home = Path.home()
    for name in ("Desktop","Pulpit"):
        p = home / name
        if p.exists():
            return p
    return home

def _ensure_per_region_files(label: str) -> tuple[Path, Path]:
    base = _detect_desktop() / "baza danych"
    linki_dir = base / "linki"
    woj_dir   = base / "województwa"
    linki_dir.mkdir(parents=True, exist_ok=True)
    woj_dir.mkdir(parents=True, exist_ok=True)

    linki_csv = linki_dir / f"{label}.csv"
    wyniki_csv = woj_dir / f"{label}.csv"

    if not linki_csv.exists():
        with linki_csv.open("w", newline="", encoding="utf-8-sig") as f:
            csv.writer(f).writerow(["link"])
    if not wyniki_csv.exists():
        header = [
            "cena","cena_za_metr","metry","liczba_pokoi","pietro","rynek","rok_budowy","material",
            "wojewodztwo","powiat","gmina","miejscowosc","dzielnica","ulica","link",
        ]
        with wyniki_csv.open("w", newline="", encoding="utf-8-sig") as f:
            csv.writer(f).writerow(header)

    return linki_csv, wyniki_csv

# =================== APP ===================
class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("900x430")
        self.resizable(False, False)

        self.var_status = tk.StringVar(value="Wybierz kategorię i województwo, aby rozpocząć.")
        self.var_phase1 = tk.StringVar(value="Etap 1 (linki): —")
        self.var_phase2 = tk.StringVar(value="Etap 2 (scraper): —")
        self._running_thread: threading.Thread | None = None

        self.var_region_label = tk.StringVar(value="Mazowieckie")
        self.var_region_slug = tk.StringVar(value="mazowieckie")

        self._build_ui()

    def _build_ui(self):
        wrapper = ttk.Frame(self, padding=16)
        wrapper.pack(fill="both", expand=True)

        ttk.Label(wrapper, text="Co chcesz pobrać?", font=("Segoe UI", 12, "bold")).pack(anchor="w")

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

        btns = ttk.Frame(wrapper)
        btns.pack(anchor="w", pady=(10, 4))
        for label, slug_cat in CATEGORIES:
            ttk.Button(
                btns, text=label, width=22,
                command=lambda s=slug_cat, l=label: self._start_pipeline(l, s)
            ).pack(side="left", padx=6, pady=6)

        ttk.Separator(wrapper, orient="horizontal").pack(fill="x", pady=(10, 8))

        info = ttk.LabelFrame(wrapper, text="Status", padding=10)
        info.pack(fill="x")
        ttk.Label(info, textvariable=self.var_status, justify="left").pack(anchor="w")
        ttk.Label(info, textvariable=self.var_phase1, justify="left").pack(anchor="w", pady=(6, 0))
        ttk.Label(info, textvariable=self.var_phase2, justify="left").pack(anchor="w")

        controls = ttk.Frame(wrapper)
        controls.pack(fill="x", pady=(12, 0))
        ttk.Button(controls, text="Zamknij", command=self._safe_close).pack(side="right")

    # ---------------- Pipeline ----------------
    def _start_pipeline(self, label_cat: str, slug_cat: str):
        if self._running_thread and self._running_thread.is_alive():
            messagebox.showinfo("W toku", "Inny proces jest w trakcie. Poczekaj.")
            return

        # poprawione slug_cat
        linki_candidates = [f"linki_{slug_cat}.py", f"linki_{slug_cat.rstrip('a')}.py"]
        scraper_candidates = [f"scraper_otodom_{slug_cat}.py", "scraper_otodom.py"]
        linki_path = next((p for n in linki_candidates if (p := _resolve_script(n)) is not None), None)
        scraper_path = next((p for n in scraper_candidates if (p := _resolve_script(n)) is not None), None)
        if not linki_path or not scraper_path:
            messagebox.showerror("Brak pliku", "Nie znaleziono wymaganych skryptów.")
            return

        region_label = self.var_region_label.get()
        linki_csv, _ = _ensure_per_region_files(region_label)

        self.var_status.set(f"Uruchamiam: {label_cat} — woj. {region_label}")
        self.var_phase1.set("Etap 1 (linki): uruchamianie…")
        self.var_phase2.set("Etap 2 (scraper): oczekiwanie…")

        def worker():
            ok1, msg1 = self._run_script_blocking(linki_path, args=["--region", region_label])
            self._set_after(lambda: self.var_phase1.set(f"Etap 1 (linki): {'✅' if ok1 else '❌'}{msg1}"))
            if not ok1:
                self._set_after(lambda: self.var_status.set("Stop: błąd w etapie 1.")); return

            self._set_after(lambda: self.var_phase2.set("Etap 2 (scraper): uruchamianie…"))
            self._set_after(self._poll_progress)  # <<< START odczytu postępu

            args = ["--input", str(linki_csv), "--label", region_label]
            ok2, msg2 = self._run_script_blocking(scraper_path, args=args)
            self._set_after(lambda: self.var_phase2.set(f"Etap 2 (scraper): {'✅' if ok2 else '❌'}{msg2}"))

            if ok2:
                self._set_after(lambda: self.var_status.set(f"Zakończono: {label_cat} — woj. {region_label}"))
            else:
                self._set_after(lambda: self.var_status.set("Zakończono z błędami (etap 2)."))

        t = threading.Thread(target=worker, daemon=True)
        self._running_thread = t
        t.start()

    def _poll_progress(self):
        """Czytaj scraper_progress.txt i aktualizuj GUI"""
        progress_file = Path("scraper_progress.txt")
        if progress_file.exists():
            try:
                status = progress_file.read_text(encoding="utf-8").strip()
                if status:
                    self.var_phase2.set(f"Etap 2 (scraper): {status}")
            except Exception:
                pass
        if self._running_thread and self._running_thread.is_alive():
            self.after(3000, self._poll_progress)

    def _run_script_blocking(self, path: Path, args: list[str] | None = None) -> tuple[bool, str]:
        try:
            proc = subprocess.run([PYTHON, str(path)] + (args or []), capture_output=False, timeout=None)
            return (proc.returncode == 0, f" (kod wyjścia={proc.returncode})")
        except FileNotFoundError:
            return (False, " (brak pliku)")
        except Exception as e:
            return (False, f" ({e})")

    def _set_after(self, fn):
        try: self.after(0, fn)
        except Exception: pass

    def _safe_close(self):
        if self._running_thread and self._running_thread.is_alive():
            if not messagebox.askyesno("Wciąż trwa", "Proces działa. Na pewno zamknąć okno?"):
                return
        self.destroy()

# =================== SCRAPER FUNKCJA ===================
def pobierz_dane_z_otodom(url: str) -> dict:
    """Pobiera dane z pojedynczego ogłoszenia Otodom"""
    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    driver = webdriver.Chrome(options=opts)

    try:
        driver.get(url)
        # czekamy tylko na <body>, które zawsze będzie
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        html = driver.page_source
    except Exception as e:
        print(f"❌ Błąd przy linku: {url} ({e})")
        with open("bledy.csv", "a", encoding="utf-8-sig", newline="") as f:
            csv.writer(f).writerow([url, str(e)])
        return {"link": url}
    finally:
        driver.quit()

    soup = BeautifulSoup(html, "html.parser")
    dane: dict[str, str] = {}

    # Cena
    cena_elem = soup.select_one("strong[data-testid='ad-price']")
    dane["cena"] = cena_elem.get_text(strip=True) if cena_elem else ""

    # Cena za m²
    cena_m2_elem = soup.select_one("div[data-testid='ad-price-per-m']")
    dane["cena_za_metr"] = cena_m2_elem.get_text(strip=True) if cena_m2_elem else ""

    # Powierzchnia
    metry_elem = soup.find("div", string=re.compile("Powierzchnia"))
    dane["metry"] = metry_elem.find_next("div").get_text(strip=True) if metry_elem else ""

    # Liczba pokoi
    pokoje_elem = soup.find("div", string=re.compile("Liczba pokoi"))
    dane["liczba_pokoi"] = pokoje_elem.find_next("div").get_text(strip=True) if pokoje_elem else ""

    # Piętro
    pietro_elem = soup.find("div", string=re.compile("Piętro"))
    dane["pietro"] = pietro_elem.find_next("div").get_text(strip=True) if pietro_elem else ""

    # Rynek
    rynek_elem = soup.find("div", string=re.compile("Rynek"))
    dane["rynek"] = rynek_elem.find_next("div").get_text(strip=True) if rynek_elem else ""

    # Rok budowy
    rok_elem = soup.find("div", string=re.compile("Rok budowy"))
    dane["rok_budowy"] = rok_elem.find_next("div").get_text(strip=True) if rok_elem else ""

    # Materiał budynku
    material_elem = soup.find("div", string=re.compile("Materiał"))
    dane["material"] = material_elem.find_next("div").get_text(strip=True) if material_elem else ""

    # Lokalizacja
    breadcrumbs = soup.select("a[data-testid='breadcrumb-link']")
    for i, crumb in enumerate(breadcrumbs):
        txt = crumb.get_text(strip=True)
        if i == 1: dane["wojewodztwo"] = txt
        elif i == 2: dane["powiat"] = txt
        elif i == 3: dane["gmina"] = txt; dane["miejscowosc"] = txt
        elif i == 4: dane["dzielnica"] = txt
        elif i == 5: dane["ulica"] = txt

    dane["link"] = url
    return dane


if __name__ == "__main__":
    App().mainloop()
