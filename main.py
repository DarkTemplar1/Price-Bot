#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
import subprocess
from importlib.util import spec_from_file_location, module_from_spec
from pathlib import Path
from typing import Optional

import tkinter as tk
from tkinter import filedialog, messagebox

import openpyxl  # noqa: F401

DOZWOLONE_ROZSZERZENIA = {".xlsx", ".xls", ".xlsm", ".xlsb"}

APP_TITLE = "Wybór pliku Excel"
SCRIPT_DALEJ = "dalej.py"
SCRIPT_KOLUMNY = "kolumny.py"  # skrypt przygotowujący arkusze/kolumny


def wybierz_plik_excel(pierwsze_okno: Optional[tk.Tk] = None) -> Optional[Path]:
    sciezka = filedialog.askopenfilename(
        parent=pierwsze_okno,
        title="Wybierz plik Excel",
        filetypes=[
            ("Pliki Excel", "*.xlsx *.xls *.xlsm *.xlsb"),
            ("Wszystkie pliki", "*.*"),
        ],
    )
    if not sciezka:
        return None

    p = Path(sciezka)
    if p.suffix.lower() not in DOZWOLONE_ROZSZERZENIA:
        messagebox.showerror(
            "Nieprawidłowy plik",
            "Wybierz plik Excela o rozszerzeniu: .xlsx, .xls, .xlsm lub .xlsb.",
            parent=pierwsze_okno,
        )
        return None

    if pierwsze_okno is not None:
        try:
            pierwsze_okno.destroy()
        except Exception:
            pass
    return p


def otworz_sortownie(sciezka_pliku: Optional[Path], root: tk.Tk):
    base = Path(__file__).resolve().parent
    sort_path = base / "sortownia.py"

    if not sort_path.exists():
        messagebox.showerror("Błąd importu", f"Nie znaleziono pliku: {sort_path}", parent=root)
        return

    try:
        spec = spec_from_file_location("sortownia_modal", str(sort_path))
        if spec is None or spec.loader is None:
            raise ImportError("Nie udało się utworzyć spec dla sortownia.py")
        mod = module_from_spec(spec)
        sys.modules["sortownia_modal"] = mod
        spec.loader.exec_module(mod)
    except Exception as e:
        messagebox.showerror("Błąd importu", f"Nie można wczytać sortownia.py:\n{e}", parent=root)
        return

    if not hasattr(mod, "run_modal"):
        messagebox.showerror(
            "Błąd",
            "Plik sortownia.py nie zawiera funkcji run_modal(file_path: str | None).",
            parent=root,
        )
        return

    root.withdraw()
    try:
        mod.run_modal(str(sciezka_pliku) if sciezka_pliku else None)
    except Exception as e:
        messagebox.showerror("Błąd sortowni", f"Wystąpił błąd w sortownia.py:\n{e}", parent=root)
    finally:
        root.deiconify()
        root.lift()
        try:
            root.focus_force()
        except Exception:
            pass


def uruchom_skrypt_subprocess(script_name: str, sciezka_pliku: Optional[Path]):
    baza = Path(__file__).resolve().parent
    script_path = (baza / script_name).resolve()
    if not script_path.exists():
        messagebox.showerror(
            "Nie znaleziono skryptu",
            f"Nie znaleziono pliku: {script_path}\nUpewnij się, że znajduje się w tym samym folderze.",
        )
        return

    cmd = [sys.executable, str(script_path)]
    if sciezka_pliku is not None:
        cmd.append(str(sciezka_pliku))

    try:
        subprocess.Popen(cmd)
    except Exception as e:
        messagebox.showerror("Błąd uruchamiania", f"Nie udało się uruchomić skryptu:\n{e}")


def uruchom_kolumny_sync(sciezka_pliku: Optional[Path]):
    """
    Uruchamia 'kolumny.py' SYNCHRONICZNIE (czeka na zakończenie),
    przekazując ścieżkę do wybranego pliku Excela.
    """
    baza = Path(__file__).resolve().parent
    kol_path = (baza / SCRIPT_KOLUMNY).resolve()
    if not kol_path.exists():
        messagebox.showerror(
            "Brak kolumny.py",
            f"Nie znaleziono pliku: {kol_path}\nPominięto automatyczne dodanie kolumn.",
        )
        return

    cmd = [sys.executable, str(kol_path)]
    if sciezka_pliku is not None:
        cmd.append(str(sciezka_pliku))

    try:
        subprocess.run(cmd, check=True)
        messagebox.showinfo("Gotowe", "Przygotowanie pliku zakończone pomyślnie.")
    except subprocess.CalledProcessError as e:
        messagebox.showerror(
            "Błąd kolumny.py",
            f"Skrypt kolumny.py zakończył się błędem (kod {e.returncode}).\nSzczegóły w konsoli.",
        )
    except Exception as e:
        messagebox.showerror("Błąd kolumny.py", f"Nie udało się uruchomić kolumny.py:\n{e}")


def zbuduj_okno_glowne(sciezka_pliku: Path):
    """Drugie menu: ścieżka + przyciski operacji (w tym 'przygotowanie pliku')."""
    root = tk.Tk()
    root.title("Panel pliku Excel")
    root.geometry("620x240")
    root.resizable(False, False)

    ramka = tk.Frame(root, padx=16, pady=16)
    ramka.pack(fill="both", expand=True)

    naglowek = tk.Label(ramka, text="Wybrany plik Excel:", font=("Segoe UI", 11, "bold"))
    naglowek.pack(anchor="w")

    etykieta_sciezka = tk.Label(
        ramka, text=str(sciezka_pliku), wraplength=580, justify="left", font=("Segoe UI", 10)
    )
    etykieta_sciezka.pack(anchor="w", pady=(4, 16))

    przyciski = tk.Frame(ramka)
    przyciski.pack(fill="x", pady=8)

    def zmien_plik():
        nowy = filedialog.askopenfilename(
            parent=root,
            title="Wybierz plik Excel",
            filetypes=[("Pliki Excel", "*.xlsx *.xls *.xlsm *.xlsb"), ("Wszystkie pliki", "*.*")],
        )
        if nowy:
            p = Path(nowy)
            if p.suffix.lower() not in DOZWOLONE_ROZSZERZENIA:
                messagebox.showerror(
                    "Nieprawidłowy plik",
                    "Wybierz plik Excela o rozszerzeniu: .xlsx, .xls, .xlsm lub .xlsb.",
                    parent=root,
                )
                return
            etykieta_sciezka.config(text=str(p))
            nonlocal_sciezka[0] = p

    nonlocal_sciezka = [sciezka_pliku]

    # — Nowy guzik: przygotowanie pliku (dodanie arkuszy/kolumn) —
    btn_kolumny = tk.Button(
        przyciski,
        text="przygotowanie pliku",
        width=20,
        command=lambda: uruchom_kolumny_sync(nonlocal_sciezka[0]),
    )
    btn_kolumny.grid(row=0, column=0, padx=6, pady=6)

    btn_zmien = tk.Button(przyciski, text="Wybierz inny plik", width=20, command=zmien_plik)
    btn_zmien.grid(row=0, column=1, padx=6, pady=6)

    btn_sort = tk.Button(
        przyciski,
        text="sortowanie danych",
        width=20,
        command=lambda: otworz_sortownie(nonlocal_sciezka[0], root),
    )
    btn_sort.grid(row=1, column=0, padx=6, pady=6)

    btn_dalej = tk.Button(
        przyciski,
        text="operacje plikowe",
        width=20,
        command=lambda: uruchom_skrypt_subprocess(SCRIPT_DALEJ, nonlocal_sciezka[0]),
    )
    btn_dalej.grid(row=1, column=1, padx=6, pady=6)

    tip = tk.Label(
        ramka,
        text=(
            "Wskazówka: zacznij od „przygotowanie pliku”, aby dodać wymagane arkusze i kolumny.\n"
            "Sortowanie i pozostałe operacje działają na aktualnie wskazanym pliku."
        ),
        fg="#555",
        wraplength=580,
        justify="left",
        font=("Segoe UI", 9),
    )
    tip.pack(anchor="w", pady=(12, 0))

    root.mainloop()


def start():
    """Pierwsze okno: wybór pliku → przejście do drugiego menu (bez auto-uruchamiania kolumn)."""
    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("600x250")
    root.resizable(False, False)

    ramka = tk.Frame(root, padx=16, pady=16)
    ramka.pack(fill="both", expand=True)

    opis = tk.Label(
        ramka,
        text="Wybierz plik Excela, aby przejść dalej.",
        font=("Segoe UI", 11),
    )
    opis.pack(pady=(0, 12))

    def wybierz_i_przejdz():
        p = wybierz_plik_excel(root)
        if p:
            # Bez automatycznego odpalania: od razu pokazujemy drugie menu,
            # gdzie jest guzik „przygotowanie pliku”.
            zbuduj_okno_glowne(p)

    btn = tk.Button(ramka, text="Wybierz plik Excel…", width=22, command=wybierz_i_przejdz)
    btn.pack(pady=6)

    root.mainloop()



if __name__ == "__main__":
    start()
