# -*- coding: utf-8 -*-
"""
automat_matma.py — pomocnicze funkcje obliczeniowe i formatowanie cen
(Wersja automatyczna – bez GUI)
"""

from __future__ import annotations
import pandas as pd

# Kolumny wymagane w bazie danych
REQUIRED_COLUMNS = [
    "cena","cena_za_metr","metry","liczba_pokoi","pietro","rynek","rok_budowy","material",
    "wojewodztwo","powiat","gmina","miejscowosc","dzielnica","ulica","link",
]

def _coerce_numeric(series: pd.Series) -> pd.Series:
    """Przekonwertuj serię na liczby (uwzględniając przecinki, spacje, kropki tysięcy)."""
    s = series.astype(str)
    s = s.str.replace("\u00a0", " ", regex=False)           # NBSP
    s = s.str.replace(",", ".", regex=False)                # przecinek -> kropka
    s = s.str.replace(r"[^\d\.\-]", "", regex=True)         # usuń znaki nie-numeryczne
    s = s.str.replace(r"(?<=\d)\.(?=.*\.)", "", regex=True) # usuń kropki tysięcy
    return pd.to_numeric(s, errors="coerce")

def remove_outliers_iqr(df: pd.DataFrame, col: str = "cena_za_metr") -> pd.DataFrame:
    """Usuń obserwacje odstające metodą IQR (1.5×IQR)."""
    s = _coerce_numeric(df[col]).dropna()
    if len(s) < 4:
        return df[_coerce_numeric(df[col]).notna()]
    q1, q3 = s.quantile(0.25), s.quantile(0.75)
    iqr = q3 - q1
    low, high = q1 - 1.5*iqr, q3 + 1.5*iqr
    mask = _coerce_numeric(df[col]).between(low, high)
    return df[mask]

def mean_numeric(series: pd.Series) -> float | None:
    """Średnia arytmetyczna z pominięciem NaN; None gdy pusto."""
    s = _coerce_numeric(series).dropna()
    if s.empty:
        return None
    return float(s.mean())

def format_currency(v: float | None) -> str:
    if v is None:
        return "—"
    return f"{v:,.0f} zł".replace(",", " ").replace(".", ",")

def format_price_per_m2(v: float | None) -> str:
    if v is None:
        return "—"
    return f"{v:,.0f} zł/m²".replace(",", " ").replace(".", ",")
