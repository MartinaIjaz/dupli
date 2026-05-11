"""
Скрипта за отстранување дупликати од Excel табела.

- Ги наоѓа редовите со исто ime (или друга колона)
- Го задржува редот со повеќе пополнети (не-празни) колони
- Ги сортира записите по датум
- Резултатот го зачувува во нов Excel фајл

Употреба:
    python deduplicate.py --input влезен_фајл.xlsx --output излезен_фајл.xlsx

Опционални аргументи:
    --name-col     Ime na kolonata so iminjata (default: auto-detect)
    --date-col     Ime na kolonata so datumite (default: auto-detect)
    --sheet        Ime na sheet-ot (default: prv sheet)
"""

import argparse
import sys
import pandas as pd


def count_filled(row: pd.Series) -> int:
    """Broi neprazni vrednosti vo red."""
    return row.notna().sum() + (row.astype(str).str.strip() != "").sum()


def find_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Najdi kolona po lista mozni imijna (case-insensitive)."""
    lower_cols = {c.lower(): c for c in df.columns}
    for name in candidates:
        if name.lower() in lower_cols:
            return lower_cols[name.lower()]
    return None


def deduplicate(
    df: pd.DataFrame,
    name_col: str,
    date_col: str | None,
) -> pd.DataFrame:
    """
    Za sekoja grupa duplikati (isto ime) go chuva samo redot so
    najgolем broj pополнети kolonи.  Потоа sortira po datum.
    """
    # Брои пополнети вредности за секој ред
    df = df.copy()
    df["__filled__"] = df.apply(count_filled, axis=1)

    # Сортира по ime + пополнетост (опаѓачки) па го зема prviot во секоја група
    df_sorted = df.sort_values(
        [name_col, "__filled__"], ascending=[True, False]
    )
    df_dedup = df_sorted.drop_duplicates(subset=[name_col], keep="first")

    # Сортира по датум ако постои
    if date_col and date_col in df_dedup.columns:
        df_dedup[date_col] = pd.to_datetime(df_dedup[date_col], errors="coerce")
        df_dedup = df_dedup.sort_values(date_col, ascending=True, na_position="last")

    df_dedup = df_dedup.drop(columns=["__filled__"])
    df_dedup = df_dedup.reset_index(drop=True)
    return df_dedup


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", required=True, help="Патека до влезниот Excel фајл")
    parser.add_argument(
        "--output",
        default=None,
        help="Патека до излезниот Excel фајл (default: deduplicated_<input>)",
    )
    parser.add_argument(
        "--name-col",
        default=None,
        help="Колона со имиња (ако не е наведена, скриптата ја бара сама)",
    )
    parser.add_argument(
        "--date-col",
        default=None,
        help="Колона со датуми (ако не е наведена, скриптата ја бара сама)",
    )
    parser.add_argument("--sheet", default=0, help="Sheet (ime или број, default: 0)")
    args = parser.parse_args()

    # Читање
    try:
        sheet = int(args.sheet) if str(args.sheet).isdigit() else args.sheet
        df = pd.read_excel(args.input, sheet_name=sheet)
    except FileNotFoundError:
        sys.exit(f"Грешка: фајлот '{args.input}' не е пронајден.")
    except Exception as exc:
        sys.exit(f"Грешка при читање: {exc}")

    print(f"Вчитани {len(df)} редови, {len(df.columns)} колони.")
    print(f"Колони: {list(df.columns)}")

    # Откривање колона со имиња
    name_col = args.name_col or find_column(
        df,
        ["Ime", "Ime i Prezime", "Полно Ime", "Name", "Full Name",
         "Вработен", "Лице", "Prezime", "Субјект"],
    )
    if not name_col:
        # Ако не е пронајдена, корисникот треба рачно да ја наведе
        print("\nНе можев да ја пронајдам колоната со имиња.")
        print("Наведи ја со --name-col \"Ime na kolonata\"")
        print("Достапни колони:", list(df.columns))
        sys.exit(1)
    print(f"Колона со имиња: '{name_col}'")

    # Откривање колона со датуми
    date_col = args.date_col or find_column(
        df,
        ["Datum", "Date", "Датум", "Дата", "Вработен од",
         "Датум на вработување", "Роденден", "Created"],
    )
    if date_col:
        print(f"Колона со датуми: '{date_col}'")
    else:
        print("Колона со датуми: не е пронајдена (нема сортирање по датум).")

    # Бројање дупликати пред
    duplicates_before = df.duplicated(subset=[name_col], keep=False).sum()
    print(f"\nРедови со дупликати (пред): {duplicates_before}")

    # Дедупликација
    df_clean = deduplicate(df, name_col, date_col)

    removed = len(df) - len(df_clean)
    print(f"Отстранети редови: {removed}")
    print(f"Останати редови: {len(df_clean)}")

    # Запишување
    output_path = args.output or f"deduplicated_{args.input}"
    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df_clean.to_excel(writer, index=False, sheet_name="Резултат")
        print(f"\nФајлот е зачуван: {output_path}")
    except Exception as exc:
        sys.exit(f"Грешка при запишување: {exc}")


if __name__ == "__main__":
    main()
