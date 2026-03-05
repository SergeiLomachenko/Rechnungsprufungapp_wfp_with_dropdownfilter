import os
import pandas as pd
from openpyxl.utils import get_column_letter
import json
import math

# Input
filename = os.environ.get("INPUT_EXCEL_PATH")
if not filename:
    raise RuntimeError("INPUT_EXCEL_PATH is required for pdf8.py")

filepath = filename if os.path.isabs(filename) else os.path.join(os.getcwd(), filename)

try:
    df = pd.read_excel(filepath, skiprows=4)
    df['Betrag'] = df['Betrag'].astype(str).str.replace(',', '.').astype(float)
    df['Erledigt'] = pd.to_datetime(df['Erledigt'], format='%d.%m.%Y', errors='coerce')
    print("Datei gefunden und geladen:", filepath)
    print(df.columns.tolist())
except FileNotFoundError:
    raise RuntimeError(f"Datei nicht gefunden: {filepath}")

# Relevant columns
new_df = df[['Hersteller', 'Fahrgestellnummer', 'Fahrzeugtyp', 'Erledigt', 'Betrag']].copy()
new_df.rename(columns={
    'Fahrgestellnummer': 'VIN',
    'Fahrzeugtyp':       'Fahrzeug'
}, inplace=True)

new_df = new_df[new_df['Fahrzeug'].notna()]
new_df = new_df[new_df['Fahrzeug'].str.strip() != ""]

# Load DB
ca3_url = os.getenv("CA3_Batterie")
rrm_url = os.getenv("RRM_Batterie")

if not ca3_url or not rrm_url:
    config_path = os.path.join(os.getcwd(), "config.json")
    if os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
        ca3_url = ca3_url or config.get("CA3_Batterie", "")
        rrm_url = rrm_url or config.get("RRM_Batterie", "")

df_ca3 = pd.read_json(ca3_url)
df_rrm = pd.read_json(rrm_url)

df_ca3["Auftraggeber"] = "CA3"
df_rrm["Auftraggeber"] = "RRM"
df_metabase = pd.concat([df_ca3, df_rrm], ignore_index=True)

print("DB geladen:", len(df_metabase), "Einträge")

# Helpers
def is_empty(v):
    if v is None:
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    return False

# Merge by VIN
# If multiple DB rows per VIN, keep the one with the highest invoice value
df_metabase_dedup = (
    df_metabase
    .sort_values("invoice", ascending=False)
    .drop_duplicates(subset=["vin"], keep="first")
)

merged_df = pd.merge(
    new_df,
    df_metabase_dedup,
    left_on="VIN",
    right_on="vin",
    how="left"
)

# Vergleich columns
# VIN_vergleich: VIN found in DB → OK, not found → NOK
merged_df["VIN_vergleich"] = merged_df["vin"].apply(
    lambda v: "NOK" if is_empty(v) else "OK"
)

# HV_vergleich: VIN found AND Batterie column filled → OK (weiterverrechnet)
merged_df["HV_vergleich"] = merged_df.apply(
    lambda row: (
        "OK" if row["VIN_vergleich"] == "OK" and not is_empty(row.get("Hochvoltbatterie"))
        else "NOK"
    ),
    axis=1
)

# Duplikate by VIN
dup_vins = set(new_df[new_df.duplicated(subset=["VIN"], keep=False)]["VIN"])

# Bemerkungen
def bemerkungen_logic(row):
    parts = []
    if row["VIN"] in dup_vins:
        parts.append("Dublikat, prüfen")
    if row["VIN_vergleich"] == "NOK":
        parts.append("Transportauftrag nicht gefunden. Bitte WFPs 2010, 9010, 9020 manuell prüfen")
    elif row["HV_vergleich"] == "NOK":
        parts.append("WFP Hochvoltbatterie nicht aktiviert, bitte prüfen")
    return " | ".join(parts)

merged_df["Bemerkungen"] = merged_df.apply(bemerkungen_logic, axis=1)

# Build output
compare_results = []

for _, row in merged_df.iterrows():
    if row["Bemerkungen"] and str(row["Bemerkungen"]).strip() != "":
        compare_results.append({
            "Hersteller":       row["Hersteller"],
            "VIN":              row["VIN"],
            "Fahrzeug":         row["Fahrzeug"],
            "Erledigt":         row["Erledigt"],
            "Betrag":           row["Betrag"],
            "Auftraggeber":     row.get("Auftraggeber"),
            "WFP_HV_Preis": row.get("Hochvoltbatterie"),
            "Bemerkungen":      row["Bemerkungen"]
        })

df_final = pd.DataFrame(compare_results)

BASE_DIR = os.environ.get("BASE_DIR")
if not BASE_DIR:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

final_file = os.path.join(BASE_DIR, "Fehlerreport.xlsx")

with pd.ExcelWriter(final_file, engine="openpyxl") as writer:
    df_final.to_excel(writer, index=False, sheet_name="Fehlerreport")
    ws = writer.sheets["Fehlerreport"]
    for i in range(1, len(df_final.columns) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 25
    ws.column_dimensions['H'].width = 45  # Bemerkungen

print(f"Fehlerreport gespeichert: {final_file}")
print(f"Fehler gesamt: {len(df_final)}")