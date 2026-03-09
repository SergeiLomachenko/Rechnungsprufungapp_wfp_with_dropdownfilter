import os
import pandas as pd
from openpyxl.utils import get_column_letter
import json
import math

# ── Input ─────────────────────────────────────────────────────────────────────
filename = os.environ.get("INPUT_EXCEL_PATH")
if not filename:
    raise RuntimeError("INPUT_EXCEL_PATH is required for pdf10.py")

filepath = filename if os.path.isabs(filename) else os.path.join(os.getcwd(), filename)

try:
    df = pd.read_excel(filepath, skiprows=4)
    df['Betrag'] = df['Betrag'].astype(str).str.replace(',', '.').astype(float)
    df['Eingang'] = pd.to_datetime(df['Eingang'], format='%d.%m.%Y', errors='coerce')
    df['Ausgang'] = pd.to_datetime(df['Ausgang'], format='%d.%m.%Y', errors='coerce')
    # Order-Nr as string (can be numeric or alphanumeric)
    def parse_order_nr(v):
        if not pd.notna(v) or str(v).strip() == '':
            return ''
        s = str(v).strip()
        try:
            return str(int(float(s)))
        except (ValueError, OverflowError):
            return s
    df['Order-Nr.'] = df['Order-Nr.'].apply(parse_order_nr)
    print("Datei geladen:", filepath)
except FileNotFoundError:
    raise RuntimeError(f"Datei nicht gefunden: {filepath}")

new_df = df[['Hersteller', 'Fahrgestellnummer', 'Order-Nr.', 'Fahrzeugtyp', 'Eingang', 'Ausgang', 'Betrag']].copy()
new_df.rename(columns={
    'Fahrgestellnummer': 'VIN',
    'Fahrzeugtyp':       'Fahrzeug',
    'Order-Nr.':         'OrderNr'
}, inplace=True)

new_df = new_df[new_df['Fahrzeug'].notna()]
new_df = new_df[new_df['Fahrzeug'].str.strip() != ""]

# ── Load DB ───────────────────────────────────────────────────────────────────
ca3_url = os.getenv("CA3_Lagerhandling")
rrm_url = os.getenv("RRM_Lagerhandling")

if not ca3_url or not rrm_url:
    config_path = os.path.join(os.getcwd(), "config.json")
    if os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
        ca3_url = ca3_url or config.get("CA3_Lagerhandling", "")
        rrm_url = rrm_url or config.get("RRM_Lagerhandling", "")

df_ca3 = pd.read_json(ca3_url)
df_rrm = pd.read_json(rrm_url)

# Normalisation: lowercase column names for easier handling
df_ca3.columns = [col.lower() for col in df_ca3.columns]
df_rrm.columns = [col.lower() for col in df_rrm.columns]

df_ca3["Auftraggeber"] = "CA3"
df_rrm["Auftraggeber"] = "RRM"
df_metabase = pd.concat([df_ca3, df_rrm], ignore_index=True)

# Normalize invoice column to string for comparison
df_metabase["invoice_str"] = df_metabase["invoice"].apply(
    lambda v: str(int(float(v))) if pd.notna(v) and str(v).strip() != '' else ''
)

print("DB geladen:", len(df_metabase), "Einträge")

# ── Helpers ───────────────────────────────────────────────────────────────────
def is_empty(v):
    if v is None:
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    return False

# ── Duplikate ─────────────────────────────────────────────────────────────────
dup_vins = set(new_df[new_df.duplicated(subset=["VIN"], keep=False)]["VIN"])

# ── Merge by VIN, use Order-Nr as tiebreaker if available ────────────────────
def get_best_db_row(vin, order_nr):
    try:
        candidates = df_metabase[df_metabase["vin"] == vin]
        if candidates.empty:
            return None
        if order_nr and str(order_nr).strip():
            match = candidates[candidates["invoice_str"] == str(order_nr).strip()]
            if not match.empty:
                row = match.iloc[0]
                return row if isinstance(row, pd.Series) else None
        return candidates.sort_values("invoice", ascending=False).iloc[0]
    except Exception as e:
        print(f"⚠️ Error: {e}")
        return None

db_rows = new_df.apply(
    lambda row: get_best_db_row(row["VIN"], row["OrderNr"]), axis=1
)

db_rows = [r if (isinstance(r, pd.Series) or r is None) else None for r in db_rows]

db_df = pd.DataFrame([r.to_dict() if r is not None else {} for r in db_rows])

# Ensure required DB columns exist even when no matches found
for col in ["vin", "invoice", "Auftraggeber", "WFP4500"]:
    if col not in db_df.columns:
        db_df[col] = None

merged_df = pd.concat([new_df.reset_index(drop=True), db_df.reset_index(drop=True)], axis=1)

# ── Vergleich ─────────────────────────────────────────────────────────────────
merged_df["VIN_vergleich"] = merged_df["vin"].apply(
    lambda v: "NOK" if is_empty(v) else "OK"
)

merged_df["WFP4500_vergleich"] = merged_df.apply(
    lambda row: (
        "OK" if row["VIN_vergleich"] == "OK" and not is_empty(row.get("WFP4500"))
        else "NOK"
    ),
    axis=1
)

# ── Bemerkungen ───────────────────────────────────────────────────────────────
def bemerkungen_logic(row):
    parts = []
    if row["VIN"] in dup_vins:
        parts.append("Dublikat, prüfen")
    if row["VIN_vergleich"] == "NOK":
        parts.append("Transportauftrag nicht gefunden. Bitte manuell prüfen")
    else:
        if row["WFP4500_vergleich"] == "NOK":
            parts.append("WFP 4500 nicht aktiviert, bitte prüfen")
        elif row["WFP4500_vergleich"] == "OK":
            parts.append("OK, wurde weiterverrechnet")
    return " | ".join(parts)

merged_df["Bemerkungen"] = merged_df.apply(bemerkungen_logic, axis=1)

# ── Output: ALL rows ──────────────────────────────────────────────────────────
df_final = merged_df[[
    "Hersteller", "VIN", "OrderNr", "Fahrzeug", "Eingang", "Ausgang", "Betrag",
    "Auftraggeber", "WFP4500", "Bemerkungen"
]].rename(columns={
    "OrderNr": "Order-Nr",
    "WFP4500": "WFP_4500_Preis"
}).copy()

BASE_DIR = os.environ.get("BASE_DIR")
if not BASE_DIR:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

final_file = os.path.join(BASE_DIR, "Fehlerreport.xlsx")

with pd.ExcelWriter(final_file, engine="openpyxl") as writer:
    df_final.to_excel(writer, index=False, sheet_name="Fehlerreport")
    ws = writer.sheets["Fehlerreport"]
    for i in range(1, len(df_final.columns) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 22
    ws.column_dimensions[get_column_letter(len(df_final.columns))].width = 50  # Bemerkungen

print(f"Fehlerreport gespeichert: {final_file}")
print(f"Zeilen gesamt: {len(df_final)}")