import os
import pandas as pd
from openpyxl.utils import get_column_letter
import json
import math

# 1. Load input file
filename = os.environ.get("INPUT_EXCEL_PATH")
if not filename:
    raise RuntimeError("INPUT_EXCEL_PATH is required for pdf10.py")

filepath = filename if os.path.isabs(filename) else os.path.join(os.getcwd(), filename)

try:
    # Read Excel, skipping the first 4 rows of the header
    df = pd.read_excel(filepath, skiprows=4)
    df['Betrag'] = df['Betrag'].astype(str).str.replace(',', '.').astype(float)
    df['Eingang'] = pd.to_datetime(df['Eingang'], format='%d.%m.%Y', errors='coerce')
    df['Ausgang'] = pd.to_datetime(df['Ausgang'], format='%d.%m.%Y', errors='coerce')
    
    # Keep Order-Nr as a clean string
    df['Order-Nr.'] = df['Order-Nr.'].astype(str).str.strip()
    
    print(f"File loaded: {filepath}")
except FileNotFoundError:
    raise RuntimeError(f"File not found: {filepath}")

# Data preparation
new_df = df[['Hersteller', 'Fahrgestellnummer', 'Order-Nr.', 'Fahrzeugtyp', 'Eingang', 'Ausgang', 'Betrag']].copy()
new_df.rename(columns={
    'Fahrgestellnummer': 'VIN',
    'Fahrzeugtyp':       'Fahrzeug',
    'Order-Nr.':         'OrderNr'
}, inplace=True)

new_df = new_df[new_df['Fahrzeug'].notna()]
new_df = new_df[new_df['Fahrzeug'].str.strip() != ""]

# VIN NORMALIZATION: Convert to UPPERCASE and remove leading/trailing spaces
new_df['VIN'] = new_df['VIN'].astype(str).str.strip().str.upper()

# 2. Load Database (Metabase)
ca3_url = os.getenv("CA3_Lagerhandling")
rrm_url = os.getenv("RRM_Lagerhandling")

if not ca3_url or not rrm_url:
    config_path = os.path.join(os.getcwd(), "config.json")
    if os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
        ca3_url = ca3_url or config.get("CA3_Lagerhandling", "")
        rrm_url = rrm_url or config.get("RRM_Lagerhandling", "")

try:
    df_ca3 = pd.read_json(ca3_url)
    df_rrm = pd.read_json(rrm_url)
    
    # Normalize column names to lowercase
    df_ca3.columns = [col.lower() for col in df_ca3.columns]
    df_rrm.columns = [col.lower() for col in df_rrm.columns]

    df_ca3["Auftraggeber"] = "CA3"
    df_rrm["Auftraggeber"] = "RRM"
    df_metabase = pd.concat([df_ca3, df_rrm], ignore_index=True)
    
    # VIN NORMALIZATION in DB: Ensure matching format
    df_metabase['vin'] = df_metabase['vin'].astype(str).str.strip().str.upper()
    
    print(f"DB loaded: {len(df_metabase)} entries")
except Exception as e:
    print(f"Error loading database: {e}")
    df_metabase = pd.DataFrame(columns=['vin', 'invoice', 'Auftraggeber', 'wfp4500'])

# 3. Search Logic (Simplified: VIN matching only)
def get_best_db_row(vin):
    if not vin or vin == 'NAN':
        return None
    try:
        # Match by normalized VIN
        candidates = df_metabase[df_metabase["vin"] == vin]
        if candidates.empty:
            return None
        # If multiple entries found, pick the latest one by invoice number
        return candidates.sort_values("invoice", ascending=False).iloc[0]
    except Exception:
        return None

# Apply search
db_rows = new_df['VIN'].apply(get_best_db_row)
db_df = pd.DataFrame([r.to_dict() if r is not None else {} for r in db_rows])

# Ensure required columns exist in the search result
for col in ["vin", "invoice", "Auftraggeber", "wfp4500"]:
    if col not in db_df.columns:
        db_df[col] = None

# Merge results
merged_df = pd.concat([new_df.reset_index(drop=True), db_df.reset_index(drop=True)], axis=1)

# 4. Validation Logic
merged_df["VIN_vergleich"] = merged_df["vin"].apply(
    lambda v: "NOK" if (v is None or str(v).strip() == "" or str(v) == 'nan') else "OK"
)

merged_df["WFP4500_vergleich"] = merged_df.apply(
    lambda row: "OK" if row["VIN_vergleich"] == "OK" and pd.notna(row.get("wfp4500")) else "NOK",
    axis=1
)

dup_vins = set(new_df[new_df.duplicated(subset=["VIN"], keep=False)]["VIN"])

def bemerkungen_logic(row):
    parts = []
    if row["VIN"] in dup_vins:
        parts.append("Duplicate in file")
    
    if row["VIN_vergleich"] == "NOK":
        parts.append("Transport order not found. Please check manually")
    else:
        if row["WFP4500_vergleich"] == "NOK":
            parts.append("WFP 4500 not activated")
        else:
            parts.append("OK, processed")
    return " | ".join(parts)

merged_df["Bemerkungen"] = merged_df.apply(bemerkungen_logic, axis=1)

# 5. Save final report
df_final = merged_df[[
    "Hersteller", "VIN", "OrderNr", "Fahrzeug", "Eingang", "Ausgang", "Betrag",
    "Auftraggeber", "wfp4500", "Bemerkungen"
]].rename(columns={
    "OrderNr": "Order-Nr",
    "wfp4500": "WFP_4500_Price"
}).copy()

final_file = os.path.join(os.environ.get("BASE_DIR", os.getcwd()), "Fehlerreport.xlsx")

with pd.ExcelWriter(final_file, engine="openpyxl") as writer:
    df_final.to_excel(writer, index=False, sheet_name="Fehlerreport")
    ws = writer.sheets["Fehlerreport"]
    for i in range(1, len(df_final.columns) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 25

print(f"Success! Report saved to: {final_file}")