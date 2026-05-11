"""
Скрипта за отстранување дупликати од Excel табела.

- Ги наоѓа редовите со исто Ime + Prezime (или друга комбинација колони)
- Ги СПОЈУВА дупликатите: зема е-маил од едниот, телефон од другиот итн.
- Резултатот е еден комплетен ред по лице со сите достапни податоци
- Ги сортира записите по датум
- Резултатот го зачувува во нов Excel фајл

Употреба:
    python3 deduplicate.py --input влезен_фајл.xlsx --output излезен_фајл.xlsx

Опционални аргументи:
    --name-cols    Имиња на колоните кои заедно го идентификуваат лицето
                   (default: auto-detect, пр. "Ime" + "Prezime" или "Ime i Prezime")
    --date-col     Колона со датуми (default: auto-detect)
    --sheet        Ime na sheet-ot (default: прв sheet)

Примери:
    # Автоматско откривање - ги наоѓа Ime и Prezime сами
    python3 deduplicate.py --input lista.xlsx --output rezultat.xlsx

    # Рачно наведување на двете колони
    python3 deduplicate.py --input lista.xlsx --output rezultat.xlsx \\
        --name-cols "Ime" "Prezime"

    # Само една комбинирана колона
    python3 deduplicate.py --input lista.xlsx --output rezultat.xlsx \\
        --name-cols "Ime i Prezime"
"""

import argparse
import sys
import pandas as pd


def find_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Најди колона по листа можни имиња (case-insensitive)."""
    lower_cols = {c.lower(): c for c in df.columns}
    for name in candidates:
        if name.lower() in lower_cols:
            return lower_cols[name.lower()]
    return None


def auto_detect_name_cols(df: pd.DataFrame) -> list[str]:
    """
    Автоматски ги открива колоните со имиња.
    Прво проверува дали постојат одделни Ime и Prezime колони.
    Ако не, бара комбинирана колона.
    """
    lower_cols = {c.lower(): c for c in df.columns}

    ime_candidates  = ["ime", "first name", "firstname", "name", "вработен", "лице", "субјект"]
    prez_candidates = ["prezime", "last name", "lastname", "surname"]
    full_candidates = ["ime i prezime", "full name", "fullname", "полно ime",
                       "ime prezime", "name surname"]

    ime_col  = next((lower_cols[c] for c in ime_candidates  if c in lower_cols), None)
    prez_col = next((lower_cols[c] for c in prez_candidates if c in lower_cols), None)
    full_col = next((lower_cols[c] for c in full_candidates if c in lower_cols), None)

    if ime_col and prez_col:
        return [ime_col, prez_col]
    if full_col:
        return [full_col]
    if ime_col:
        return [ime_col]
    return []


def _first_non_empty(series: pd.Series):
    """Врати ја првата непразна вредност во серијата, или NaN ако нема."""
    for val in series:
        if pd.notna(val) and str(val).strip() != "":
            return val
    return pd.NA


def merge_duplicates(
    df: pd.DataFrame,
    name_cols: list[str],
    date_col: str | None,
) -> pd.DataFrame:
    """
    За секоја група дупликати (исто Ime + Prezime):
      - За секоја колона зема ја ПРВАТА непразна вредност од сите дупликати
      - Така се спојуваат сите податоци во еден комплетен ред
    Потоа сортира по датум.
    """
    all_cols = list(df.columns)
    rows = []

    for _key, group in df.groupby(name_cols, sort=False, dropna=False):
        merged = {col: _first_non_empty(group[col]) for col in all_cols}
        rows.append(merged)

    df_merged = pd.DataFrame(rows, columns=all_cols)

    if date_col and date_col in df_merged.columns:
        df_merged[date_col] = pd.to_datetime(df_merged[date_col], errors="coerce")
        df_merged = df_merged.sort_values(date_col, ascending=True, na_position="last")

    df_merged = df_merged.reset_index(drop=True)
    return df_merged


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--input", required=True, help="Патека до влезниот Excel фајл")
    parser.add_argument(
        "--output",
        default=None,
        help="Патека до излезниот Excel фајл (default: deduplicated_<input>)",
    )
    parser.add_argument(
        "--name-cols",
        nargs="+",
        default=None,
        metavar="КОЛОНА",
        help='Колони кои го идентификуваат лицето, пр. --name-cols "Ime" "Prezime"',
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

    # Откривање колони со имиња
    if args.name_cols:
        name_cols = args.name_cols
        missing = [c for c in name_cols if c not in df.columns]
        if missing:
            sys.exit(f"Грешка: колоните {missing} не постојат во фајлот.\n"
                     f"Достапни колони: {list(df.columns)}")
    else:
        name_cols = auto_detect_name_cols(df)
        if not name_cols:
            print("\nНе можев да ги пронајдам колоните со имиња.")
            print('Наведи ги рачно со --name-cols "Ime" "Prezime"')
            print("Достапни колони:", list(df.columns))
            sys.exit(1)

    print(f"Клучни колони (за дупликати): {name_cols}")

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
    duplicates_before = df.duplicated(subset=name_cols, keep=False).sum()
    print(f"\nРедови со дупликати (пред): {duplicates_before}")

    # Спојување дупликати
    df_clean = merge_duplicates(df, name_cols, date_col)

    removed = len(df) - len(df_clean)
    print(f"Отстранети редови: {removed}")
    print(f"Останати редови:   {len(df_clean)}")

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
