#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import subprocess
import sys
import os
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from pathlib import Path

# lista dostępnych województw
WOJEWODZTWA = [
    "Dolnośląskie", "Kujawsko-Pomorskie", "Lubelskie", "Lubuskie",
    "Łódzkie", "Małopolskie", "Mazowieckie", "Opolskie",
    "Podkarpackie", "Podlaskie", "Pomorskie", "Śląskie",
    "Świętokrzyskie", "Warmińsko-Mazurskie", "Wielkopolskie", "Zachodniopomorskie"
]

# baza folderów na Pulpicie
BASE_DIR = os.path.join(os.path.expanduser("~"), "Desktop", "baza danych")
LINKI_DIR = os.path.join(BASE_DIR, "linki")
WOJ_DIR = os.path.join(BASE_DIR, "województwa")

# upewniamy się, że katalogi istnieją
os.makedirs(LINKI_DIR, exist_ok=True)
os.makedirs(WOJ_DIR, exist_ok=True)


# === Pomocnicze do uruchamiania innych skryptów ===
def _resolve_script(name: str) -> Path | None:
    base = Path(__file__).resolve().parent
    p = (base / name).resolve()
    return p if p.exists() else None


def _run_script(name: str, args: list[str] | None = None, exit_after: bool = False):
    """
    Uruchamia skrypt name (jeśli istnieje). Jeśli exit_after=True, zamyka to okno i proces.
    """
    p = _resolve_script(name)
    if p is None:
        messagebox.showerror("Nie znaleziono skryptu",
                             f"Nie znaleziono pliku: {name}\nSprawdź, czy znajduje się w folderze:\n{Path(__file__).resolve().parent}")
        return

    cmd = [sys.executable, str(p)]
    if args:
        cmd += args

    try:
        subprocess.Popen(cmd)
    except Exception as e:
        messagebox.showerror("Błąd uruchamiania", f"Nie udało się uruchomić '{name}':\n{e}")
        return

    if exit_after:
        try:
            root.destroy()
        finally:
            sys.exit(0)


# funkcja uruchamiająca scraper linków
def uruchom_scraper(region):
    log_text.insert(tk.END, f"[INFO] Startuję proces pobierania linków dla {region}\n")
    log_text.see(tk.END)

    output_file = os.path.join(LINKI_DIR, f"{region}.csv")

    try:
        subprocess.run([
            sys.executable, "linki_mieszkania.py",
            "--region", region,
            "--output", output_file
        ], check=True)
        log_text.insert(tk.END, f"[OK] Linki zapisane do {output_file}\n")
    except subprocess.CalledProcessError as e:
        log_text.insert(tk.END, f"[BŁĄD] Proces linków zakończył się błędem: {e}\n")
        log_text.see(tk.END)

    # uruchom scraper ofert
    input_file = output_file
    output_woj = os.path.join(WOJ_DIR, f"{region}.csv")

    try:
        subprocess.run([
            sys.executable, "scraper_otodom.py",
            "--region", region,
            "--input", input_file,
            "--output", output_woj
        ], check=True)
        log_text.insert(tk.END, f"[OK] Dane zapisane do {output_woj}\n")
    except subprocess.CalledProcessError as e:
        log_text.insert(tk.END, f"[BŁĄD] Scraper ofert zakończył się błędem: {e}\n")
        log_text.see(tk.END)


# obsługa przycisku START
def start_process():
    region = region_var.get()
    uruchom_scraper(region)


# NOWE: obsługa przycisku "Powrót"
def powrot_do_dalej():
    # wyłącz ten skrypt i włącz dalej.py
    _run_script("dalej.py", exit_after=True)


# NOWE: obsługa przycisku "Scal"
def uruchom_scalanie():
    log_text.insert(tk.END, "[INFO] Uruchamiam scalanie (scalanie.py)...\n")
    log_text.see(tk.END)
    _run_script("scalanie.py", exit_after=False)


# GUI
root = tk.Tk()
root.title("PriceBot - Baza Danych")

frame = tk.Frame(root)
frame.pack(pady=10, padx=10, fill="x")

# wybór województwa
tk.Label(frame, text="Wybierz województwo:").grid(row=0, column=0, sticky="w")
region_var = tk.StringVar(value=WOJEWODZTWA[0])
region_menu = ttk.Combobox(frame, textvariable=region_var, values=WOJEWODZTWA, state="readonly")
region_menu.grid(row=0, column=1, padx=5, sticky="ew")

# przyciski: Start, Scal, Powrót
start_btn = tk.Button(frame, text="Start", command=start_process)
start_btn.grid(row=1, column=1, pady=10, padx=5, sticky="ew")

scal_btn = tk.Button(frame, text="Scal", command=uruchom_scalanie)
scal_btn.grid(row=1, column=0, pady=10, padx=5, sticky="ew")

powrot_btn = tk.Button(frame, text="⟵ Powrót", command=powrot_do_dalej)
powrot_btn.grid(row=1, column=2, pady=10, padx=5, sticky="ew")

# układ kolumn
frame.grid_columnconfigure(0, weight=1)
frame.grid_columnconfigure(1, weight=1)
frame.grid_columnconfigure(2, weight=1)

# logi
log_text = scrolledtext.ScrolledText(root, width=70, height=20)
log_text.pack(padx=10, pady=10, fill="both", expand=True)

root.mainloop()
