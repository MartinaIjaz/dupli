"""Тестови за deduplicate.py"""

import pandas as pd
import pytest
from deduplicate import count_filled, deduplicate, find_column, auto_detect_name_cols


# ---------------------------------------------------------------------------
# Помошни функции
# ---------------------------------------------------------------------------

def make_df_single_col():
    """Табела со единствена комбинирана колона 'Ime'."""
    data = {
        "Ime":     ["Марко Петров", "Ана Иванова",  "Марко Петров", "Ана Иванова",  "Јован Јовиќ"],
        "Датум":   ["2023-03-15",   "2022-07-01",   "2023-03-15",   "2022-07-01",   "2024-01-20"],
        "Телефон": ["070 111 222",  None,            None,           "071 333 444",  "072 555 666"],
        "Адреса":  ["ул. Мир 1",   "ул. Сонце 5",  None,           None,           "ул. Роза 3"],
        "Е-пошта": [None,           None,            "marko@mk.com", "ana@mk.com",   "jovan@mk.com"],
    }
    return pd.DataFrame(data)


def make_df_two_cols():
    """Табела со одделни колони 'Ime' и 'Prezime'."""
    data = {
        "Ime":     ["Марко", "Ана",    "Марко", "Ана",    "Јован"],
        "Prezime": ["Петров","Иванова","Петров","Иванова","Јовиќ"],
        "Датум":   ["2023-03-15", "2022-07-01", "2023-03-15", "2022-07-01", "2024-01-20"],
        "Телефон": ["070 111 222", None,         None,         "071 333 444", "072 555 666"],
        "Адреса":  ["ул. Мир 1",  "ул. Сонце 5", None,        None,          "ул. Роза 3"],
        "Е-пошта": [None,          None,          "marko@mk.com", "ana@mk.com", "jovan@mk.com"],
    }
    return pd.DataFrame(data)


# ---------------------------------------------------------------------------
# Единечни тестови
# ---------------------------------------------------------------------------

def test_count_filled_basic():
    row = pd.Series({"a": "x", "b": None, "c": "y", "d": ""})
    assert count_filled(row) > 0


def test_find_column_case_insensitive():
    df = pd.DataFrame(columns=["ime", "Datum", "Telefon"])
    assert find_column(df, ["Ime", "Name"]) == "ime"
    assert find_column(df, ["datum", "date"]) == "Datum"
    assert find_column(df, ["missing"]) is None


# ---------------------------------------------------------------------------
# auto_detect_name_cols
# ---------------------------------------------------------------------------

def test_auto_detect_two_separate_cols():
    """Ако постојат и 'Ime' и 'Prezime', ги враќа двете."""
    df = make_df_two_cols()
    result = auto_detect_name_cols(df)
    assert result == ["Ime", "Prezime"]


def test_auto_detect_combined_col():
    """Ако постои само 'Ime i Prezime', ја враќа неа."""
    df = pd.DataFrame(columns=["Ime i Prezime", "Датум", "Телефон"])
    result = auto_detect_name_cols(df)
    assert result == ["Ime i Prezime"]


def test_auto_detect_fallback_single_ime():
    """Ако постои само 'Ime', ја враќа неа."""
    df = make_df_single_col()
    result = auto_detect_name_cols(df)
    assert result == ["Ime"]


def test_auto_detect_no_match():
    """Ако нема позната колона, враќа празна листа."""
    df = pd.DataFrame(columns=["Код", "Датум", "Вредност"])
    result = auto_detect_name_cols(df)
    assert result == []


# ---------------------------------------------------------------------------
# deduplicate - единствена колона (назадна компатибилност)
# ---------------------------------------------------------------------------

def test_single_col_keeps_more_data():
    df = make_df_single_col()
    result = deduplicate(df, name_cols=["Ime"], date_col="Датум")
    assert len(result) == 3
    assert set(result["Ime"]) == {"Марко Петров", "Ана Иванова", "Јован Јовиќ"}


def test_single_col_marko_keeps_phone():
    df = make_df_single_col()
    result = deduplicate(df, name_cols=["Ime"], date_col="Датум")
    marko = result[result["Ime"] == "Марко Петров"].iloc[0]
    assert marko["Телефон"] == "070 111 222"


def test_single_col_sorted_by_date():
    df = make_df_single_col()
    result = deduplicate(df, name_cols=["Ime"], date_col="Датум")
    dates = pd.to_datetime(result["Датум"]).tolist()
    assert dates == sorted(dates)


# ---------------------------------------------------------------------------
# deduplicate - две одделни колони Ime + Prezime
# ---------------------------------------------------------------------------

def test_two_cols_removes_duplicates():
    df = make_df_two_cols()
    result = deduplicate(df, name_cols=["Ime", "Prezime"], date_col="Датум")
    assert len(result) == 3


def test_two_cols_correct_names():
    df = make_df_two_cols()
    result = deduplicate(df, name_cols=["Ime", "Prezime"], date_col="Датум")
    pairs = set(zip(result["Ime"], result["Prezime"]))
    assert pairs == {("Марко", "Петров"), ("Ана", "Иванова"), ("Јован", "Јовиќ")}


def test_two_cols_marko_keeps_phone():
    """Марко Петров: ред 0 (Телефон + Адреса) > ред 2 (само Е-пошта)."""
    df = make_df_two_cols()
    result = deduplicate(df, name_cols=["Ime", "Prezime"], date_col="Датум")
    marko = result[(result["Ime"] == "Марко") & (result["Prezime"] == "Петров")].iloc[0]
    assert marko["Телефон"] == "070 111 222"


def test_two_cols_sorted_by_date():
    df = make_df_two_cols()
    result = deduplicate(df, name_cols=["Ime", "Prezime"], date_col="Датум")
    dates = pd.to_datetime(result["Датум"]).tolist()
    assert dates == sorted(dates)


def test_two_cols_different_surname_not_duplicate():
    """Исто Ime, различно Prezime — НЕ се дупликати."""
    df = pd.DataFrame({
        "Ime":     ["Марко", "Марко"],
        "Prezime": ["Петров", "Николов"],
        "Датум":   ["2023-01-01", "2023-06-01"],
    })
    result = deduplicate(df, name_cols=["Ime", "Prezime"], date_col="Датум")
    assert len(result) == 2


def test_no_duplicates_unchanged():
    df = pd.DataFrame({
        "Ime":     ["Лице", "Лице"],
        "Prezime": ["А",    "Б"],
        "Датум":   ["2020-01-01", "2021-06-15"],
    })
    result = deduplicate(df, name_cols=["Ime", "Prezime"], date_col="Датум")
    assert len(result) == 2
