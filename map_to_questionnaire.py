"""
Ги пренесува податоците од постоечки Excel фајл во структурата на
"Draft 1 Mapping Questionnaire - Entrepreneurs" (REDI NGO).

Употреба:
    python3 map_to_questionnaire.py --input vashiot_fajl.xlsx --output questionnaire.xlsx

Скриптата автоматски ги препознава колоните (Ime, Prezime, Телефон, Е-пошта итн.)
и ги пренесува во соодветното поле на прашалникот.
Полиња за кои нема податоци остануваат празни — за REDI staff да ги пополни.
"""

import argparse
import sys
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Структура на прашалникот — редослед и имиња на колоните
# ---------------------------------------------------------------------------

QUESTIONNAIRE_COLUMNS = [
    "REDI Staff Member Name",
    "Date of Mapping",
    "Name Surname",
    "Contact Email",
    "Contact Phone Number",
    "Country",
    "Place of Residence (Municipality / City)",
    "Address",
    "Gender",
    "Year of Birth",
    "Level of Education",
    "Do you identify as Roma?",
    "If NO Roma – Connection to Roma Community",
    "Business Name",
    "Address of the Business",
    "Short Description of the Business",
    "Business Sector",
    "Is the Business Registered?",
    "Date of Business Registration",
    "Number of Employees",
    "Was Business Supported in Phase I?",
    "Year of Initial Support",
    "Type of Service Needed",
    "Interested in Business Club Membership?",
    "Previously had a Business Loan?",
    "Informed about Measures / Help for Business",
    "[REDI Staff] Assessment of Client Needs",
    "[REDI Staff] Action Recommended",
]

# ---------------------------------------------------------------------------
# Авто-детекција: мап од можни имиња на влезни колони → поле на прашалник
# ---------------------------------------------------------------------------

COLUMN_MAP: dict[str, str] = {
    # Ime / Name
    "ime i prezime":            "Name Surname",
    "name surname":             "Name Surname",
    "full name":                "Name Surname",
    "полно ime":                "Name Surname",
    # Презиме се спојува подоцна ако се одделни
    "ime":                      "__ime__",
    "prezime":                  "__prezime__",
    "first name":               "__ime__",
    "last name":                "__prezime__",
    "name":                     "__ime__",
    "surname":                  "__prezime__",
    # Е-пошта
    "е-пошта":                  "Contact Email",
    "email":                    "Contact Email",
    "e-mail":                   "Contact Email",
    "e mail":                   "Contact Email",
    "контакт email":            "Contact Email",
    # Телефон
    "телефон":                  "Contact Phone Number",
    "phone":                    "Contact Phone Number",
    "phone number":             "Contact Phone Number",
    "контакт телефон":          "Contact Phone Number",
    "mobile":                   "Contact Phone Number",
    # Земја
    "земја":                    "Country",
    "country":                  "Country",
    "држава":                   "Country",
    # Место / Општина
    "општина":                  "Place of Residence (Municipality / City)",
    "municipality":             "Place of Residence (Municipality / City)",
    "место":                    "Place of Residence (Municipality / City)",
    "град":                     "Place of Residence (Municipality / City)",
    "city":                     "Place of Residence (Municipality / City)",
    "place of residence":       "Place of Residence (Municipality / City)",
    # Адреса
    "адреса":                   "Address",
    "address":                  "Address",
    "улица":                    "Address",
    # Пол
    "пол":                      "Gender",
    "gender":                   "Gender",
    "sex":                      "Gender",
    # Година на раѓање
    "година на раѓање":         "Year of Birth",
    "year of birth":            "Year of Birth",
    "роденден":                 "Year of Birth",
    "birthdate":                "Year of Birth",
    "date of birth":            "Year of Birth",
    # Образование
    "образование":              "Level of Education",
    "level of education":       "Level of Education",
    "education":                "Level of Education",
    # Рома
    "рома":                     "Do you identify as Roma?",
    "roma":                     "Do you identify as Roma?",
    "is roma":                  "Do you identify as Roma?",
    # Врска со ромска заедница
    "врска со ромска заедница": "If NO Roma – Connection to Roma Community",
    "connection to roma":       "If NO Roma – Connection to Roma Community",
    # Бизнис
    "бизнис":                   "Business Name",
    "business name":            "Business Name",
    "назив на фирма":           "Business Name",
    "фирма":                    "Business Name",
    # Адреса на бизнис
    "адреса на бизнис":         "Address of the Business",
    "business address":         "Address of the Business",
    "address of the business":  "Address of the Business",
    # Опис на бизнис
    "опис на бизнис":           "Short Description of the Business",
    "business description":     "Short Description of the Business",
    "description":              "Short Description of the Business",
    # Сектор
    "сектор":                   "Business Sector",
    "sector":                   "Business Sector",
    "business sector":          "Business Sector",
    # Регистрација
    "регистриран":              "Is the Business Registered?",
    "registered":               "Is the Business Registered?",
    "is registered":            "Is the Business Registered?",
    # Датум на регистрација
    "датум на регистрација":    "Date of Business Registration",
    "date of registration":     "Date of Business Registration",
    "registration date":        "Date of Business Registration",
    # Вработени
    "вработени":                "Number of Employees",
    "employees":                "Number of Employees",
    "number of employees":      "Number of Employees",
    "broj vraboteni":           "Number of Employees",
    # Фаза I
    "фаза i":                   "Was Business Supported in Phase I?",
    "phase i":                  "Was Business Supported in Phase I?",
    "phase 1":                  "Was Business Supported in Phase I?",
    # Година на поддршка
    "година на поддршка":       "Year of Initial Support",
    "year of support":          "Year of Initial Support",
    # Тип на услуга
    "тип на услуга":            "Type of Service Needed",
    "service needed":           "Type of Service Needed",
    "service":                  "Type of Service Needed",
    # Бизнис клуб
    "бизнис клуб":              "Interested in Business Club Membership?",
    "business club":            "Interested in Business Club Membership?",
    # Кредит
    "кредит":                   "Previously had a Business Loan?",
    "loan":                     "Previously had a Business Loan?",
    "business loan":            "Previously had a Business Loan?",
    # Информиран
    "информиран":               "Informed about Measures / Help for Business",
    "informed":                 "Informed about Measures / Help for Business",
    # REDI staff
    "процена":                  "[REDI Staff] Assessment of Client Needs",
    "assessment":               "[REDI Staff] Assessment of Client Needs",
    "потребни акции":           "[REDI Staff] Action Recommended",
    "action":                   "[REDI Staff] Action Recommended",
    "action recommended":       "[REDI Staff] Action Recommended",
    # Датум на мапирање
    "датум":                    "Date of Mapping",
    "date":                     "Date of Mapping",
    "date of mapping":          "Date of Mapping",
}


def detect_mapping(df: pd.DataFrame) -> dict[str, str]:
    """
    За секоја колона во df, пронајди го соодветното поле на прашалникот.
    Враќа речник: {col_in_df: questionnaire_field}.
    """
    mapping = {}
    lower_cols = {c.lower().strip(): c for c in df.columns}

    ime_col = prez_col = None

    for lower, original in lower_cols.items():
        target = COLUMN_MAP.get(lower)
        if target == "__ime__":
            ime_col = original
        elif target == "__prezime__":
            prez_col = original
        elif target:
            mapping[original] = target

    return mapping, ime_col, prez_col


def build_questionnaire(
    df: pd.DataFrame,
    mapping: dict[str, str],
    ime_col: str | None,
    prez_col: str | None,
) -> pd.DataFrame:
    """Гради нова табела со структурата на прашалникот."""
    rows = []
    for _, row in df.iterrows():
        qrow: dict[str, object] = {col: None for col in QUESTIONNAIRE_COLUMNS}

        # Спои Ime + Prezime во Name Surname
        if ime_col and prez_col:
            ime  = str(row[ime_col]).strip()  if pd.notna(row[ime_col])  else ""
            prez = str(row[prez_col]).strip() if pd.notna(row[prez_col]) else ""
            full = f"{ime} {prez}".strip()
            if full:
                qrow["Name Surname"] = full
        elif ime_col:
            qrow["Name Surname"] = row[ime_col] if pd.notna(row[ime_col]) else None

        # Пренеси ги останатите колони
        for src_col, dst_col in mapping.items():
            val = row[src_col]
            if pd.notna(val) and str(val).strip() != "":
                # Не пребришувај ако веќе пополнето (Name Surname)
                if qrow[dst_col] is None:
                    qrow[dst_col] = val

        rows.append(qrow)

    return pd.DataFrame(rows, columns=QUESTIONNAIRE_COLUMNS)


# ---------------------------------------------------------------------------
# Excel форматирање
# ---------------------------------------------------------------------------

def _col_group(col: str) -> str:
    if col.startswith("[REDI"):
        return "redi"
    if col in ("REDI Staff Member Name", "Date of Mapping"):
        return "meta"
    if col in ("Name Surname", "Contact Email", "Contact Phone Number",
               "Country", "Place of Residence (Municipality / City)",
               "Address", "Gender", "Year of Birth", "Level of Education",
               "Do you identify as Roma?", "If NO Roma – Connection to Roma Community"):
        return "personal"
    return "business"


GROUP_COLORS = {
    "meta":     ("1F4E79", "D6E4F0"),   # темно сино  / светло сино
    "personal": ("375623", "E2EFDA"),   # темно зелено / светло зелено
    "business": ("7B3F00", "FCE4D6"),   # темно кафеа / светло праскова
    "redi":     ("4B0082", "EDE7F6"),   # виолетова   / светло виолетова
}


def format_questionnaire(ws, df: pd.DataFrame) -> None:
    thin  = Side(style="thin", color="BFBFBF")
    bdr   = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap  = Alignment(horizontal="left", vertical="center", wrap_text=True)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # --- Заглавија ---
    for col_idx, col_name in enumerate(df.columns, start=1):
        group = _col_group(col_name)
        hdr_color, _ = GROUP_COLORS[group]
        cell = ws.cell(row=1, column=col_idx)
        cell.value     = col_name
        cell.font      = Font(bold=True, color="FFFFFF", size=10)
        cell.fill      = PatternFill("solid", fgColor=hdr_color)
        cell.alignment = center
        cell.border    = bdr

    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 42

    # --- Редови со податоци ---
    for row_idx in range(2, ws.max_row + 1):
        group_of_row = None
        for col_idx, col_name in enumerate(df.columns, start=1):
            group = _col_group(col_name)
            _, row_color = GROUP_COLORS[group]
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font      = Font(size=10)
            cell.border    = bdr
            cell.alignment = wrap
            if row_idx % 2 == 0:
                cell.fill = PatternFill("solid", fgColor=row_color)
        ws.row_dimensions[row_idx].height = 18

    # --- Ширина на колони ---
    for col_idx, col_name in enumerate(df.columns, start=1):
        col_letter = get_column_letter(col_idx)
        # Заглавие + максимална вредност
        max_len = max(
            len(col_name),
            max(
                (len(str(ws.cell(row=r, column=col_idx).value or ""))
                 for r in range(2, ws.max_row + 1)),
                default=0,
            ),
        )
        ws.column_dimensions[col_letter].width = min(max_len + 3, 35)


def add_legend_sheet(wb) -> None:
    """Додај sheet со легенда за боите."""
    ws = wb.create_sheet("Легенда")
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 40

    thin = Side(style="thin", color="BFBFBF")
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)

    rows = [
        ("Боја", "Значење"),
        ("Темно сино / светло сино", "Мета-информации (REDI Staff, Датум)"),
        ("Темно зелено / светло зелено", "Лични податоци на клиентот"),
        ("Темно кафеаво / светло праскова", "Податоци за бизнис"),
        ("Виолетова / светло виолетова", "Проценка и препорака на REDI Staff"),
    ]
    colors = [
        ("1F4E79", "D6E4F0"),
        ("375623", "E2EFDA"),
        ("7B3F00", "FCE4D6"),
        ("4B0082", "EDE7F6"),
    ]

    header = ws[1]
    for col_idx, val in enumerate(rows[0], start=1):
        cell = ws.cell(row=1, column=col_idx, value=val)
        cell.font   = Font(bold=True, color="FFFFFF")
        cell.fill   = PatternFill("solid", fgColor="333333")
        cell.border = bdr
        cell.alignment = Alignment(horizontal="center")

    for i, (label, desc) in enumerate(rows[1:], start=2):
        hdr_c, row_c = colors[i - 2]
        ca = ws.cell(row=i, column=1, value=label)
        ca.fill   = PatternFill("solid", fgColor=hdr_c)
        ca.font   = Font(color="FFFFFF", bold=True)
        ca.border = bdr
        ca.alignment = Alignment(horizontal="left", vertical="center")

        cb = ws.cell(row=i, column=2, value=desc)
        cb.fill   = PatternFill("solid", fgColor=row_c)
        cb.border = bdr
        cb.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[i].height = 20


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--input",  required=True, help="Влезен Excel фајл со податоци")
    parser.add_argument("--output", default="questionnaire_output.xlsx",
                        help="Излезен Excel фајл (default: questionnaire_output.xlsx)")
    parser.add_argument("--sheet",  default=0, help="Sheet (број или ime, default: 0)")
    args = parser.parse_args()

    try:
        sheet = int(args.sheet) if str(args.sheet).isdigit() else args.sheet
        df = pd.read_excel(args.input, sheet_name=sheet)
    except FileNotFoundError:
        sys.exit(f"Грешка: фајлот '{args.input}' не е пронајден.")
    except Exception as exc:
        sys.exit(f"Грешка при читање: {exc}")

    print(f"Вчитани {len(df)} редови.")
    print(f"Влезни колони: {list(df.columns)}")

    mapping, ime_col, prez_col = detect_mapping(df)

    print(f"\nПронајдени пресликувања:")
    if ime_col and prez_col:
        print(f"  '{ime_col}' + '{prez_col}'  →  Name Surname")
    elif ime_col:
        print(f"  '{ime_col}'  →  Name Surname")
    for src, dst in mapping.items():
        print(f"  '{src}'  →  '{dst}'")

    df_q = build_questionnaire(df, mapping, ime_col, prez_col)

    filled = df_q.notna().sum().sum()
    total  = len(df_q) * len(QUESTIONNAIRE_COLUMNS)
    print(f"\nПополнети ќелии: {filled} / {total} "
          f"({100 * filled // total}%)")
    empty_cols = [c for c in QUESTIONNAIRE_COLUMNS if df_q[c].isna().all()]
    if empty_cols:
        print(f"Празни колони (нема извор): {empty_cols}")

    # Запишување
    try:
        with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
            df_q.to_excel(writer, index=False, sheet_name="Прашалник")
            format_questionnaire(writer.sheets["Прашалник"], df_q)
        # Додај легенда
        from openpyxl import load_workbook
        wb = load_workbook(args.output)
        add_legend_sheet(wb)
        wb.save(args.output)
        print(f"\nФајлот е зачуван: {args.output}")
    except Exception as exc:
        sys.exit(f"Грешка при запишување: {exc}")


if __name__ == "__main__":
    main()
