# -*- coding: utf-8 -*-
"""
wyniki_matma.py — pomocnicze funkcje obliczeniowe i formatowanie cen
"""

import pandas as pd

REQUIRED_COLUMNS = [
    "cena","cena_za_metr","metry","liczba_pokoi","pietro","rynek","rok_budowy","material",
    "wojewodztwo","powiat","gmina","miejscowosc","dzielnica","ulica","link",
]

def _coerce_numeric(series: pd.Series) -> pd.Series:
    s = series.astype(str)
    s = s.str.replace("\u00a0", " ", regex=False)
    s = s.str.replace(",", ".", regex=False)
    s = s.str.replace(r"[^\d\.\-]", "", regex=True)
    s = s.str.replace(r"(?<=\d)\.(?=.*\.)", "", regex=True)
    return pd.to_numeric(s, errors="coerce")

def remove_outliers_iqr(df: pd.DataFrame, col: str = "cena_za_metr") -> pd.DataFrame:
    s = _coerce_numeric(df[col]).dropna()
    if len(s) < 4:
        return df[_coerce_numeric(df[col]).notna()]
    q1, q3 = s.quantile(0.25), s.quantile(0.75)
    iqr = q3 - q1
    low, high = q1 - 1.5*iqr, q3 + 1.5*iqr
    mask = _coerce_numeric(df[col]).between(low, high)
    return df[mask]

def format_currency(v: float | None) -> str:
    if v is None: return "—"
    return f"{v:,.0f} zł".replace(",", " ").replace(".", ",")

def format_price_per_m2(v: float | None) -> str:
    if v is None: return "—"
    return f"{v:,.0f} zł/m²".replace(",", " ").replace(".", ",")
