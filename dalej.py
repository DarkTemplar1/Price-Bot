#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import messagebox

APP_TITLE = "Operacje plikowe"

# Kandydaci nazw skryptów (pierwszy istniejący zostanie uruchomiony)
SCRIPT_BAZA = ["bazadanych.py", "bazydanych.py"]   # obsługa obu pisowni
SCRIPT_WYNIK = ["wyniki.py", "wynik.py"]           # obsługa starej nazwy
SCRIPT_AUTOMAT = ["automat.py"]                    # nowy przycisk "Automat"

# Dozwolone rozszerzenia plików Excela (informacyjne – nie wymuszamy wyboru tutaj)
DOZWOLONE_ROZSZERZENIA = {".xlsx", ".xls", ".xlsm", ".xlsb"}


def _resolve_script(base_dir: Path, candidates: list[str]) -> Path | None:
    """
    Zwraca pierwszą istniejącą ścieżkę do skryptu w tym samym folderze co dalej.py.
    """
    for name in candidates:
        p = (base_dir / name).resolve()
        if p.exists():
            return p
    return None


def _launch_and_exit(script_candidates: list[str], file_arg: Path | None, root: tk.Tk):
    """
    Uruchamia wskazany skrypt (z opcjonalnym argumentem ścieżki pliku) i zamyka obecne okno/proces.
    """
    base = Path(__file__).resolve().parent
    script_path = _resolve_script(base, script_candidates)
    if script_path is None:
        pretty = " lub ".join(script_candidates)
        messagebox.showerror(
            "Nie znaleziono skryptu",
            f"Nie znaleziono pliku: {pretty}\nSprawdź, czy znajduje się w folderze:\n{base}",
            parent=root,
        )
        return

    cmd = [sys.executable, str(script_path)]
    if file_arg is not None:
        cmd.append(str(file_arg))

    try:
        subprocess.Popen(cmd)
    except Exception as e:
        messagebox.showerror("Błąd uruchamiania", f"Nie udało się uruchomić '{script_path.name}':\n{e}", parent=root)
        return

    # zamknij to okno i zakończ proces dalej.py
    try:
        root.destroy()
    finally:
        sys.exit(0)


def _go_back(root: tk.Tk):
    """
    Zachowuj się jak w sortownia.py: po prostu zamknij to okno i wyjdź.
    (Nie uruchamiaj ponownie main.py – on już działa.)
    """
    try:
        root.destroy()
    finally:
        sys.exit(0)


def _build_ui(selected_path: Path | None):
    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("720x220")  # poszerzone, by wygodnie zmieścić 4 przyciski
    root.resizable(False, False)

    container = tk.Frame(root, padx=16, pady=16)
    container.pack(fill="both", expand=True)

    # Nagłówek
    hdr = tk.Label(container, text="Wybrany plik Excel:", font=("Segoe UI", 11, "bold"))
    hdr.pack(anchor="w")

    # Ścieżka (tylko do wglądu)
    txt = str(selected_path) if selected_path else "(nie przekazano)"
    path_label = tk.Label(container, text=txt, wraplength=680, justify="left", font=("Segoe UI", 10))
    path_label.pack(anchor="w", pady=(4, 12))

    # Podpowiedź o rozszerzeniach (informacyjnie)
    if selected_path and selected_path.suffix.lower() not in DOZWOLONE_ROZSZERZENIA:
        tip_ext = tk.Label(
            container,
            text="Uwaga: to nie wygląda na standardowy plik Excela.",
            fg="#a33",
            font=("Segoe UI", 9),
        )
        tip_ext.pack(anchor="w", pady=(0, 8))

    # Przyciski
    row = tk.Frame(container)
    row.pack(fill="x", pady=8)

    # ⟵ Wróć – tylko zamyka to okno (nie odpala main.py)
    btn_back = tk.Button(
        row,
        text="⟵ Wróć",
        width=20,
        command=lambda: _go_back(root),
    )
    btn_back.grid(row=0, column=0, padx=6, pady=6)

    btn_baza = tk.Button(
        row,
        text="Baza danych",
        width=20,
        command=lambda: _launch_and_exit(SCRIPT_BAZA, selected_path, root),
    )
    btn_baza.grid(row=0, column=1, padx=6, pady=6)

    btn_wynik = tk.Button(
        row,
        text="Wynik manualne",
        width=20,
        command=lambda: _launch_and_exit(SCRIPT_WYNIK, selected_path, root),
    )
    btn_wynik.grid(row=0, column=2, padx=6, pady=6)

    btn_automat = tk.Button(
        row,
        text="Automat",
        width=20,
        command=lambda: _launch_and_exit(SCRIPT_AUTOMAT, selected_path, root),
    )
    btn_automat.grid(row=0, column=3, padx=6, pady=6)

    # Podpowiedź
    tip = tk.Label(
        container,
        text=("Przyciski uruchamiają wskazany skrypt (z tą samą ścieżką pliku jako 1. argument), "
              "a bieżące okno zostaje zamknięte. 'Wróć' tylko zamyka to okno."),
        fg="#555",
        wraplength=680,
        justify="left",
        font=("Segoe UI", 9),
    )
    tip.pack(anchor="w", pady=(12, 0))

    # Zamknięcie okna krzyżykiem – po prostu zakończ ten skrypt
    root.protocol("WM_DELETE_WINDOW", root.destroy)

    root.mainloop()


def main():
    # Odczytaj opcjonalną ścieżkę do Excela z argv[1]
    selected_path: Path | None = None
    if len(sys.argv) > 1 and sys.argv[1].strip():
        try:
            p = Path(sys.argv[1]).expanduser().resolve()
            selected_path = p
        except Exception:
            selected_path = None

    _build_ui(selected_path)


if __name__ == "__main__":
    main()
