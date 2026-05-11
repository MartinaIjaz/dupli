"""Тестови за deduplicate.py"""

import pandas as pd
import pytest
from deduplicate import merge_duplicates, find_column, auto_detect_name_cols


# ---------------------------------------------------------------------------
# Помошни функции
# ---------------------------------------------------------------------------

def make_df_single_col():
    """Табела со единствена комбинирана колона 'Ime' — дупликатите имаат различни податоци."""
    data = {
        "Ime":     ["Марко Петров", "Ана Иванова",  "Марко Петров", "Ана Иванова",  "Јован Јовиќ"],
        "Датум":   ["2023-03-15",   "2022-07-01",   "2023-03-15",   "2022-07-01",   "2024-01-20"],
        "Телефон": ["070 111 222",  None,            None,           "071 333 444",  "072 555 666"],
        "Адреса":  ["ул. Мир 1",   "ул. Сонце 5",  None,           None,           "ул. Роза 3"],
        "Е-пошта": [None,           None,            "marko@mk.com", "ana@mk.com",   "jovan@mk.com"],
    }
    return pd.DataFrame(data)


def make_df_two_cols():
    """Табела со одделни 'Ime' и 'Prezime' колони."""
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
# find_column
# ---------------------------------------------------------------------------

def test_find_column_case_insensitive():
    df = pd.DataFrame(columns=["ime", "Datum", "Telefon"])
    assert find_column(df, ["Ime", "Name"]) == "ime"
    assert find_column(df, ["datum", "date"]) == "Datum"
    assert find_column(df, ["missing"]) is None


# ---------------------------------------------------------------------------
# auto_detect_name_cols
# ---------------------------------------------------------------------------

def test_auto_detect_two_separate_cols():
    df = make_df_two_cols()
    assert auto_detect_name_cols(df) == ["Ime", "Prezime"]


def test_auto_detect_combined_col():
    df = pd.DataFrame(columns=["Ime i Prezime", "Датум", "Телефон"])
    assert auto_detect_name_cols(df) == ["Ime i Prezime"]


def test_auto_detect_fallback_single_ime():
    df = make_df_single_col()
    assert auto_detect_name_cols(df) == ["Ime"]


def test_auto_detect_no_match():
    df = pd.DataFrame(columns=["Код", "Датум", "Вредност"])
    assert auto_detect_name_cols(df) == []


# ---------------------------------------------------------------------------
# merge_duplicates — клучна логика: спојување на сите податоци
# ---------------------------------------------------------------------------

def test_merge_single_col_count():
    """Од 5 → 3 уникатни редови."""
    df = make_df_single_col()
    result = merge_duplicates(df, name_cols=["Ime"], date_col="Датум")
    assert len(result) == 3


def test_merge_marko_has_both_phone_and_email():
    """Марко: ред 0 → Телефон+Адреса, ред 2 → Е-пошта.
       По спојување мора да ги има и трите."""
    df = make_df_single_col()
    result = merge_duplicates(df, name_cols=["Ime"], date_col="Датум")
    marko = result[result["Ime"] == "Марко Петров"].iloc[0]
    assert marko["Телефон"] == "070 111 222"
    assert marko["Адреса"] == "ул. Мир 1"
    assert marko["Е-пошта"] == "marko@mk.com"


def test_merge_ana_has_both_address_and_email():
    """Ана: ред 1 → Адреса, ред 3 → Телефон+Е-пошта.
       По спојување мора да ги има сите."""
    df = make_df_single_col()
    result = merge_duplicates(df, name_cols=["Ime"], date_col="Датум")
    ana = result[result["Ime"] == "Ана Иванова"].iloc[0]
    assert ana["Адреса"] == "ул. Сонце 5"
    assert ana["Телефон"] == "071 333 444"
    assert ana["Е-пошта"] == "ana@mk.com"


def test_merge_sorted_by_date():
    df = make_df_single_col()
    result = merge_duplicates(df, name_cols=["Ime"], date_col="Датум")
    dates = pd.to_datetime(result["Датум"]).tolist()
    assert dates == sorted(dates)


def test_merge_two_cols_count():
    df = make_df_two_cols()
    result = merge_duplicates(df, name_cols=["Ime", "Prezime"], date_col="Датум")
    assert len(result) == 3


def test_merge_two_cols_marko_complete():
    """По спојување Марко Петров мора да има Телефон, Адреса И Е-пошта."""
    df = make_df_two_cols()
    result = merge_duplicates(df, name_cols=["Ime", "Prezime"], date_col="Датум")
    marko = result[(result["Ime"] == "Марко") & (result["Prezime"] == "Петров")].iloc[0]
    assert marko["Телефон"] == "070 111 222"
    assert marko["Адреса"] == "ул. Мир 1"
    assert marko["Е-пошта"] == "marko@mk.com"


def test_merge_two_cols_ana_complete():
    """По спојување Ана Иванова мора да има Адреса, Телефон И Е-пошта."""
    df = make_df_two_cols()
    result = merge_duplicates(df, name_cols=["Ime", "Prezime"], date_col="Датум")
    ana = result[(result["Ime"] == "Ана") & (result["Prezime"] == "Иванова")].iloc[0]
    assert ana["Адреса"] == "ул. Сонце 5"
    assert ana["Телефон"] == "071 333 444"
    assert ana["Е-пошта"] == "ana@mk.com"


def test_merge_two_cols_sorted_by_date():
    df = make_df_two_cols()
    result = merge_duplicates(df, name_cols=["Ime", "Prezime"], date_col="Датум")
    dates = pd.to_datetime(result["Датум"]).tolist()
    assert dates == sorted(dates)


def test_different_surname_not_duplicate():
    """Исто Ime, различно Prezime — НЕ се дупликати."""
    df = pd.DataFrame({
        "Ime":     ["Марко", "Марко"],
        "Prezime": ["Петров", "Николов"],
        "Датум":   ["2023-01-01", "2023-06-01"],
        "Е-пошта": ["p@mk.com",   "n@mk.com"],
    })
    result = merge_duplicates(df, name_cols=["Ime", "Prezime"], date_col="Датум")
    assert len(result) == 2


def test_no_duplicates_unchanged():
    df = pd.DataFrame({
        "Ime":     ["Лице", "Лице"],
        "Prezime": ["А",    "Б"],
        "Датум":   ["2020-01-01", "2021-06-15"],
    })
    result = merge_duplicates(df, name_cols=["Ime", "Prezime"], date_col="Датум")
    assert len(result) == 2


def test_column_order_preserved():
    """Редоследот на колоните мора да остане ист."""
    df = make_df_two_cols()
    result = merge_duplicates(df, name_cols=["Ime", "Prezime"], date_col="Датум")
    assert list(result.columns) == list(df.columns)
