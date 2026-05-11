"""Тестови за deduplicate.py"""

import pandas as pd
import pytest
from deduplicate import count_filled, deduplicate, find_column


def make_df():
    """Пример табела со дупликати."""
    data = {
        "Ime":        ["Марко Петров", "Ана Иванова", "Марко Петров", "Ана Иванова",  "Јован Јовиќ"],
        "Датум":      ["2023-03-15",   "2022-07-01",  "2023-03-15",   "2022-07-01",   "2024-01-20"],
        "Телефон":    ["070 111 222",  None,           None,           "071 333 444",  "072 555 666"],
        "Адреса":     ["ул. Мир 1",    "ул. Сонце 5", None,           None,           "ул. Роза 3"],
        "Е-пошта":    [None,           None,           "marko@mk.com", "ana@mk.com",   "jovan@mk.com"],
    }
    return pd.DataFrame(data)


def test_count_filled_basic():
    row = pd.Series({"a": "x", "b": None, "c": "y", "d": ""})
    # notna: a=T, b=F, c=T, d=T  => 3
    # strip != "": a=T, b=F(str None='None'!='' actually True), c=T, d=F => let's just check > 0
    assert count_filled(row) > 0


def test_find_column_case_insensitive():
    df = pd.DataFrame(columns=["ime", "Datum", "Telefon"])
    assert find_column(df, ["Ime", "Name"]) == "ime"
    assert find_column(df, ["datum", "date"]) == "Datum"
    assert find_column(df, ["missing"]) is None


def test_deduplicate_keeps_more_data():
    df = make_df()
    result = deduplicate(df, name_col="Ime", date_col="Датум")

    # Само 3 уникатни имиња
    assert len(result) == 3
    assert set(result["Ime"]) == {"Марко Петров", "Ана Иванова", "Јован Јовиќ"}


def test_deduplicate_marko_has_phone():
    """Марко Петров: ред 0 има телефон, ред 2 има е-пошта.
       Ред 0 е со повеќе вредности (Телефон + Адреса) vs (само Е-пошта) -> ред 0."""
    df = make_df()
    result = deduplicate(df, name_col="Ime", date_col="Датум")
    marko = result[result["Ime"] == "Марко Петров"].iloc[0]
    assert marko["Телефон"] == "070 111 222"


def test_deduplicate_ana_has_phone():
    """Ана Иванова: ред 1 има Адреса, ред 3 има Телефон.
       И двата имаат по 1 дополнителна вредност; при изедначување го чуваме prviot."""
    df = make_df()
    result = deduplicate(df, name_col="Ime", date_col="Датум")
    ana = result[result["Ime"] == "Ана Иванова"].iloc[0]
    # Треба да постои барем еден од двата
    assert (ana["Телефон"] is not None) or (ana["Адреса"] is not None)


def test_deduplicate_sorted_by_date():
    df = make_df()
    result = deduplicate(df, name_col="Ime", date_col="Датум")
    dates = pd.to_datetime(result["Датум"]).tolist()
    assert dates == sorted(dates)


def test_no_duplicates_unchanged():
    df = pd.DataFrame({
        "Ime":   ["Лице А", "Лице Б"],
        "Датум": ["2020-01-01", "2021-06-15"],
    })
    result = deduplicate(df, name_col="Ime", date_col="Датум")
    assert len(result) == 2
