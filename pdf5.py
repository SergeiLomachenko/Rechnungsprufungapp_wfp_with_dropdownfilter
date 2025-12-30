import os 
import pandas as pd 
from openpyxl import load_workbook
import json
import math
import httpx

# Require the uploaded file path from app.py via environment.
filename = os.environ.get("INPUT_EXCEL_PATH")
if not filename:
    raise RuntimeError("INPUT_EXCEL_PATH is required for pdf5.py")

filepath = filename if os.path.isabs(filename) else os.path.join(os.getcwd(), filename)

try: 
    df = pd.read_excel(filepath, skiprows=4)
    df['Betrag'] = df['Betrag'].str.replace(',', '.').astype(float)
    df['Erledigt'] = pd.to_datetime(df['Erledigt'], format='%d.%m.%Y', errors='coerce')

    print("Die Datei 'Ausweise und schilde' ist gefunnden und hochgeladen", filepath)
    print(df.info())
    print(df.head())
    print(df.columns.tolist())

except FileNotFoundError:
    print("Die Datei 'Ausweise und schilde' ist nicht gefunnden und hochgeladen", filepath)

new_df = df[['Hersteller', 'Fahrgestellnummer', 'Fahrzeugtyp', 'Erledigt', 'Betrag']].copy()
new_df.rename(columns={
    'Fahrgestellnummer': 'VIN',
    'Fahrzeugtyp': 'Fahrzeug'
}, inplace=True)

new_df = new_df[new_df['Fahrzeug'].notna()]
new_df = new_df[new_df['Fahrzeug'].str.strip() != ""]

# output_file = "output.xlsx"
# with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
#     new_df.to_excel(writer, index=False, sheet_name="Daten")
#     worksheet = writer.sheets["Daten"]

#     worksheet.column_dimensions['B'].width = 25 
#     worksheet.column_dimensions['C'].width = 25 
#     worksheet.column_dimensions['D'].width = 25 
# print("Neue Datei erstellt und gespeichert in output.xlsx")


url_ausweise_schilder = os.getenv("ca3_ausweise_schilder")

if not url_ausweise_schilder:
    config_path = os.path.join(os.getcwd(), "config.json")
    if os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
        url_ausweise_schilder = config.get("ca3_ausweise_schilder", "")


df_metabase = pd.read_json(url_ausweise_schilder)
print(df_metabase.head())

# Merge new_df and df_metabase using VIN
merged_df = pd.merge(
    new_df,
    df_metabase,
    left_on="VIN",
    right_on="vin",
    how="left"
)

def is_empty(v):
    if v is None:
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    return False

def compare_null_logic(val1, val2):
    if is_empty(val1) and is_empty(val2):
        return "OK"
    elif (not is_empty(val1)) and (not is_empty(val2)):
        return "OK"
    else:
        return "NOK"

# VIN_vergleich
merged_df["VIN_vergleich"] = merged_df.apply(
    lambda row: compare_null_logic(row.get("VIN"), row.get("vin")),
    axis=1
)

# Schild_vergleich
merged_df["Schild_vergleich"] = merged_df.apply(
    lambda row: (
        "OK" if row.get("VIN_vergleich") == "OK"
                and not is_empty(row.get("Schild"))
                and not is_empty(row.get("Betrag"))
        else "NOK"
    ),
    axis=1
)
# Ausweis_vergleich
merged_df["Ausweis_vergleich"] = merged_df.apply(
    lambda row: (
        "OK" if row.get("VIN_vergleich") == "OK"
                and not is_empty(row.get("Ausweis"))
                and not is_empty(row.get("Betrag"))
        else "NOK"
    ),
    axis=1
)

def bemerkungen_logic(row):
    if row.get("VIN_vergleich") == "NOK":
        return "Transportauftrag nicht gefunden. Bitte WFPs 2010, 9010, 9020, 3300 und 3020 manuell prüfen"
    elif row.get("VIN_vergleich") == "OK" and row.get("Schild_vergleich") == "NOK" and row.get("Ausweis_vergleich") == "NOK":
        return "WFPs 3020 und 3300 nicht aktiviert, bitte prüfen"
    elif row.get("VIN_vergleich") == "OK" and row.get("Schild_vergleich") == "NOK":
        return "WFP 3300 nicht aktiviert, bitte prüfen"
    elif row.get("VIN_vergleich") == "OK" and row.get("Ausweis_vergleich") == "NOK":
        return "WFP 3020 nicht aktiviert, bitte prüfen"
    else:
        return ""

merged_df["Bemerkungen"] = merged_df.apply(bemerkungen_logic, axis=1)

# output_file = "vergleich.xlsx"
# with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
#     merged_df.to_excel(writer, index=False, sheet_name="Vergleich")
#     worksheet = writer.sheets["Vergleich"]

#     worksheet.column_dimensions['B'].width = 22
#     worksheet.column_dimensions['C'].width = 22
#     worksheet.column_dimensions['D'].width = 20
#     worksheet.column_dimensions['F'].width = 12
#     worksheet.column_dimensions['G'].width = 8
#     worksheet.column_dimensions['H'].width = 22
#     worksheet.column_dimensions['I'].width = 25
#     worksheet.column_dimensions['L'].width = 15
#     worksheet.column_dimensions['M'].width = 15
#     worksheet.column_dimensions['N'].width = 18
#     worksheet.column_dimensions['O'].width = 35

# print("Ergebnisse gespeichert in der Datei vergleich.xlsx")

compare_results = []

for i, row in merged_df.iterrows():
    if row["Bemerkungen"] and str(row["Bemerkungen"]).strip() != "":
        fehler_row = {
            "Hersteller": row["Hersteller"],
            "VIN": row["VIN"], 
            "Fahrzeug": row["Fahrzeug"],
            "Erledigt": row["Erledigt"],
            "Betrag": row["Betrag"],
            "Auftraggeber": row.get("Auftraggeber"),
            "WFP_3300_preis": row.get("Schild"),
            "WFP_3020_preis": row.get("Ausweis"),
            "Bemerkungen": row["Bemerkungen"]
        }
        compare_results.append(fehler_row)

df_final = pd.DataFrame(compare_results)
BASE_DIR = os.environ.get("BASE_DIR")
if not BASE_DIR:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
final_file = os.path.join(BASE_DIR, "Fehlerreport.xlsx")

with pd.ExcelWriter(final_file, engine="openpyxl") as writer:
    df_final.to_excel(writer, index=False, sheet_name="Fehlerreport")
    worksheet = writer.sheets["Fehlerreport"]

    worksheet.column_dimensions['B'].width = 22
    worksheet.column_dimensions['C'].width = 22
    worksheet.column_dimensions['D'].width = 20
    worksheet.column_dimensions['F'].width = 12
    worksheet.column_dimensions['G'].width = 18
    worksheet.column_dimensions['H'].width = 18
    worksheet.column_dimensions['I'].width = 40

