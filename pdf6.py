import os 
import pandas as pd 
from openpyxl import load_workbook
import json
import math
import httpx

# Require the uploaded file path from app.py via environment.
filename = os.environ.get("INPUT_EXCEL_PATH")
if not filename:
    raise RuntimeError("INPUT_EXCEL_PATH is required for pdf6.py")

filepath = filename if os.path.isabs(filename) else os.path.join(os.getcwd(), filename)

# Read Excel file - pandas automatically closes the file after reading
df = pd.read_excel(filepath, skiprows = 4, engine='openpyxl')
df['Betrag'] = df['Betrag'].str.replace(',', '.').astype(float)
# df['Erledigt'] = pd.to_datetime(df['Erledigt'], format='%d.%m.%Y', errors='coerce')

# print(df.head())

desc_col = None 
for candidate in ['Text', 'Beschreibung']: 
    if candidate in df.columns: 
        desc_col = candidate 
        break 

if desc_col is None: 
    raise ValueError("In der Datei 'Text' und 'Beschreibung' nicht gefunfen")


new_df = df[['Hersteller', 'Fahrgestellnummer', 'Fahrzeugtyp', 'Dienstleistung', desc_col, 'Erledigt', 'Betrag']].copy()
new_df.rename(columns={
    'Fahrgestellnummer': 'VIN',
    'Fahrzeugtyp': 'Fahrzeug',
    'Dienstleistung': 'Code',
    desc_col: 'Dienstleistung'
}, inplace=True)

new_df = new_df[new_df['Fahrzeug'].notna()]
new_df = new_df[new_df['Fahrzeug'].str.strip() != ""]

print(new_df.head())

url_service_leistungen = os.getenv("ca3_service_leistungen")
url_service_leistungen_rrm = os.getenv("rrm_service_leistungen")

if not url_service_leistungen or not url_service_leistungen_rrm:
    config_path = os.path.join(os.getcwd(), "config.json")
    if os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
        url_service_leistungen = config.get("ca3_service_leistungen", "")
        url_service_leistungen_rrm = config.get("rrm_service_leistungen", "")


df_metabase = pd.read_json(url_service_leistungen)
print(df_metabase.head())

df_metabase_rrm = pd.read_json(url_service_leistungen_rrm)
print(df_metabase_rrm.head())

# Объединяем оба датафрейма из Metabase (CA3 и RRM)
df_metabase_combined = pd.concat([df_metabase, df_metabase_rrm], ignore_index=True)

# Merge new_df с объединенным df_metabase_combined using VIN
merged_df = pd.merge(
    new_df,
    df_metabase_combined,
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

def check_code_logic_4200(row):
    # check VIN_vergleich
    if row.get("VIN_vergleich") == "OK":
        # wenn Code == 0010.0
        if row.get("Code") == 0010.0:
            if not is_empty(row.get("4200")):
                return "OK"
            else:
                return "NOK"
    # wenn Code != 0010.0 or VIN_vergleich != "OK" -> nothing to display
    return None

# 4200Vergleich
merged_df["4200vergleich"] = merged_df.apply(check_code_logic_4200, axis=1)

def check_code_logic_4415(row):
    # check VIN_vergleich
    if row.get("VIN_vergleich") == "OK":
        # wenn Code == 0010.0
        if row.get("Code") == 1524.0:
            if not is_empty(row.get("4415")):
                return "OK"
            else:
                return "NOK"
    # wenn Code != 1524.0 or VIN_vergleich != "OK" -> nothing to display
    return None

# 4415Vergleich 
merged_df["4415vergleich"] = merged_df.apply(check_code_logic_4415, axis=1)

# 4410Vergleich
def check_code_logic_4410(row):
    # check VIN_vergleich
    if row.get("VIN_vergleich") == "OK":
        # wenn Code == 0010.0
        if row.get("Code") == 2355.0:
            if not is_empty(row.get("4410")):
                return "OK"
            else:
                return "NOK"
    # wenn Code != 2355.0 or VIN_vergleich != "OK" -> nothing to display
    return None

# 4410Vergleich
merged_df["4410vergleich"] = merged_df.apply(check_code_logic_4410, axis=1)

def check_code_logic_4500(row):
    # check VIN_vergleich 
    if row.get("VIN_vergleich") == "OK":
        if row.get("Code") == 3140.0:
            if not is_empty(row.get("4500")):
                return "OK"
            else:
                return "NOK"
    # wenn Code != 3140.0 or VIN_vergleich != "OK" -> nothing to display 
    return None 

merged_df["4500vergleich"] = merged_df.apply(check_code_logic_4500, axis = 1)

def check_code_logic_3010(row):
    # check VIN_vergleich
    if row.get("VIN_vergleich") == "OK":
        if row.get("Code") == 66674.0:
            if not is_empty(row.get("3010")):
                return "OK"
            else:
                return "NOK"
    # wenn Code != 66674.0 or VIN_vergleich != "OK" -> nothing to display 
    return None

merged_df["3010vergleich"] = merged_df.apply(check_code_logic_3010, axis = 1)

def check_code_logic_4040(row):
    # check VIN_vergleich
    if row.get("VIN_vergleich") == "OK":
        if row.get("Code") == 68000.0 or row.get("Code") == 68001.0:
            if not is_empty(row.get("4040")):
                return "OK"
            else:
                return "NOK"
    # wenn Code != 68000.0 or VIN_vergleich != "OK" -> nothing to display 
    return None

merged_df["4040vergleich"] = merged_df.apply(check_code_logic_4040, axis = 1) 

def check_code_logic_4030(row):
    # check VIN_vergleich
    if row.get("VIN_vergleich") == "OK":
        if row.get("Code") == 68002.0 or row.get("Code") == 68003.0:
            if not is_empty(row.get("4030")):
                return "OK"
            else:
                return "NOK"
    # wenn Code != 68002.0 or VIN_vergleich != "OK" -> nothing to display 
    return None

merged_df["4030vergleich"] = merged_df.apply(check_code_logic_4030, axis = 1)

def check_code_logic_4020(row):
    # check VIN_vergleich
    if row.get("VIN_vergleich") == "OK":
        if row.get("Code") == 68004.0:
            if not is_empty(row.get("4020")):
                return "OK"
            else:
                return "NOK"
    # wenn Code != 68004.0 or VIN_vergleich != "OK" -> nothing to display 
    return None

merged_df["4020vergleich"] = merged_df.apply(check_code_logic_4020, axis = 1)

def check_code_logic_4900ent(row):
    # check VIN_vergleich
    if row.get("VIN_vergleich") == "OK":
        if row.get("Code") == 68006.0 or row.get("Code") == 68005.0 or row.get("Code") == 68008.0 or row.get("Code") == 66662.0 or row.get("Code") == 68090.0:
            if not is_empty(row.get("4900ent")):
                return "OK"
            else:
                return "NOK"
    # wenn Code != 4900.0 or VIN_vergleich != "OK" -> nothing to display 
    return None

merged_df["4900entvergleich"] = merged_df.apply(check_code_logic_4900ent, axis = 1)

def check_code_logic_20202040(row):
    # check VIN_vergleich
    if row.get("VIN_vergleich") == "OK":
        if row.get("Code") == 9806.0:
            if not is_empty(row.get("20202040")):
                return "OK"
            else:
                return "NOK"
    # wenn Code != 2020.0 or VIN_vergleich != "OK" -> nothing to display 
    return None

merged_df["20202040vergleich"] = merged_df.apply(check_code_logic_20202040, axis = 1)

def check_code_logic_not90109020(row):
    # check VIN_vergleich
    if row.get("VIN_vergleich") == "OK":
        if row.get("Code") == 8200.0:
            if not is_empty(row.get("9010")) or not is_empty(row.get("not90109020")):
                return "OK"
            else:
                return "NOK"
    return None

merged_df["not90109020vergleich"] = merged_df.apply(check_code_logic_not90109020, axis = 1)

def bemerkungen_logic(row):
    # Erst prüfen, ob der VIN-Match fehlt
    if row.get("VIN_vergleich") == "NOK":
        return "Transportauftrag nicht gefunden. Bitte alles manuell überprüfen."

    # Einzelne WFP-Prüfungen je Code
    if row.get("4200vergleich") == "NOK":
        return "WFP 4200 nicht aktiviert (für Code 0010)"
    if row.get("4415vergleich") == "NOK":
        return "WFP 4415 nicht aktiviert (für Code 1524)"
    if row.get("4410vergleich") == "NOK":
        return "WFP 4410 nicht aktiviert (für Code 2355)"
    if row.get("4500vergleich") == "NOK":
        return "WFP 4500 nicht aktiviert (für Code 3140)"
    if row.get("3010vergleich") == "NOK":
        return "WFP 3010 nicht aktiviert (für Code 66674)"
    if row.get("4040vergleich") == "NOK":
        return "WFP 4040 nicht aktiviert (für Code 68001 or 68000)"
    if row.get("4030vergleich") == "NOK":
        return "WFP 4030 nicht aktiviert (für Code 68003 or 68002)"
    if row.get("4020vergleich") == "NOK":
        return "WFP 4020 nicht aktiviert (für Code 68004)"
    if row.get("4900entvergleich") == "NOK":
        return "WFP 4900 nicht aktiviert oder kein Entschriftung-Comment im WFP 4900 (für Code 68006, 68008, 68005, 66662, 68090)"
    if row.get("20202040vergleich") == "NOK":
        return "WFP 2020 or WFP 2040 nicht aktiviert (für Code 9806)"
    if row.get("not90109020vergleich") == "NOK":
        return "WFP 9010 or WFP 9020 ist aktiv (für Code 8200)"
    return ""

merged_df["Bemerkungen"] = merged_df.apply(bemerkungen_logic, axis=1)

# output_file = "vergleich.xlsx"
# with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
#     merged_df.to_excel(writer, index=False, sheet_name="Vergleich")
#     worksheet = writer.sheets["Vergleich"]

#     worksheet.column_dimensions['B'].width = 22
#     worksheet.column_dimensions['C'].width = 22
#     worksheet.column_dimensions['D'].width = 10
#     worksheet.column_dimensions['E'].width = 25
#     worksheet.column_dimensions['F'].width = 12
#     worksheet.column_dimensions['G'].width = 8
#     worksheet.column_dimensions['H'].width = 22
#     worksheet.column_dimensions['L'].width = 15
#     worksheet.column_dimensions['M'].width = 15
#     worksheet.column_dimensions['N'].width = 18
#     worksheet.column_dimensions['R'].width = 22
#     worksheet.column_dimensions['U'].width = 20
#     worksheet.column_dimensions['AG'].width = 45

# print("Ergebnisse gespeichert in der Datei vergleich.xlsx")

compare_results = []

for i, row in merged_df.iterrows():
    if row["Bemerkungen"] and str(row["Bemerkungen"]).strip() != "":
        fehler_row = {
            "Hersteller": row["Hersteller"],
            "VIN": row["VIN"], 
            "Fahrzeug": row["Fahrzeug"],
            "Code": row["Code"],
            "Dienstleistung": row["Dienstleistung"],
            "Erledigt": row["Erledigt"],
            "Betrag": row["Betrag"],
            "Auftraggeber": row.get("Auftraggeber"),
            "Bemerkungen": row["Bemerkungen"]
        }
        compare_results.append(fehler_row)

# Ensure file is created in the same directory as the script (BASE_DIR from app.py)
BASE_DIR = os.environ.get("BASE_DIR")
if not BASE_DIR:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
final_file = os.path.join(BASE_DIR, "Fehlerreport.xlsx")


# Always remove old Fehlerreport.xlsx before creating new one to avoid caching issues
if os.path.exists(final_file):
    try:
        os.remove(final_file)
    except (PermissionError, OSError) as e:
        print(f"Could not remove old {final_file}: {e}")

# Create Fehlerreport.xlsx - always create file, even if no errors
if len(compare_results) == 0:
    # No errors found - create file with message
    df_final = pd.DataFrame([{"Status": "Datei bearbeitet. Keine Fehlermeldungen."}])
else:
    # Errors found - create file with error data
    df_final = pd.DataFrame(compare_results)

with pd.ExcelWriter(final_file, engine="openpyxl", mode='w') as writer:
    df_final.to_excel(writer, index=False, sheet_name="Fehlerreport")
    worksheet = writer.sheets["Fehlerreport"]
    
    if len(compare_results) > 0:
        # Set column widths for error data
        worksheet.column_dimensions['B'].width = 22
        worksheet.column_dimensions['C'].width = 22
        worksheet.column_dimensions['D'].width = 10
        worksheet.column_dimensions['E'].width = 30
        worksheet.column_dimensions['F'].width = 18
        worksheet.column_dimensions['G'].width = 10
        worksheet.column_dimensions['F'].width = 15
        worksheet.column_dimensions['H'].width = 12
        worksheet.column_dimensions['I'].width = 55
    else:
        # Set width for status message column
        worksheet.column_dimensions['A'].width = 50

# Verify file was created
if os.path.exists(final_file):
    file_size = os.path.getsize(final_file)
    print(f"Fehlerreport.xlsx successfully created at {final_file} with {len(df_final)} row(s), size: {file_size} bytes")
else:
    print(f"ERROR: Fehlerreport.xlsx was not created at {final_file}!")
