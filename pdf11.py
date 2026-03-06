import re
import pdfplumber
import pandas as pd
import numpy as np
from openpyxl.utils import get_column_letter
import os
import json
import math

# PDF Parsing (from pdf4.py)
pdf_path = "invoice.pdf"

with pdfplumber.open(pdf_path) as pdf:
    pages = [page.extract_text() for page in pdf.pages]

# Alle Detailseiten zu einem Text zusammenführen (ohne Deckblatt und Zahlungsseite)
# Seiten-Header-Zeilen herausfiltern damit keine false Treffer beim Ziffer-Check entstehen
def strip_page_header(text):
    """Entfernt die ersten 3 Zeilen (Seitenkopf) jeder Folgeseite."""
    lines = text.split("\n")
    # Header-Zeilen sind typischerweise: Datum-Block oben, dann leer, dann Vorgangsnummer
    # Wir behalten alles ab der ersten Zeile die mit einem Datum (z.B. "7.11.25") anfängt
    return "\n".join(lines)

combined_text = "\n".join(pages[2:])  # Seite 1=Deckblatt, Seite 2=Zahlteil überspringen

ZIFFERN = ["BM", "VW", "DC", "FI", "NI", "ME", "VF", "VX", "GF", "PS",
           "EF", "TX", "AR", "RN", "WBA", "TMB", "VSS", "EU", "XX", "DA", "FO", "HY", "PO", "MA", 
"W0L", "W0V", "W1K", "W1N", "W1V", "WDD", "WDC", "WDB", "WDF",
"WME", "WMW", "WP0", "WP1", "WUA", "WVG", "WVW", "WV1", "WV2", "WV3",
"YV1", "YSM", "ZFA", "ZN6", "ZAR", "LPS", "LRW", "5YJ", "SB1", "JF1",
"JMB", "JTD", "JMZ", "KMH", "KMT", "KNA", "UU1", "VF1", "VF3", "VF7",
"VR3", "SAL", "TMA", "1C4", "WBAC", "WBAT", "WBS", "WBY"]

cars_data = []

for page_text in pages[1:]:
    lines = page_text.split("\n")

    for i, line in enumerate(lines):
        if not any(keyword in line for keyword in ZIFFERN):
            continue

        line_parts = line.split()
        if len(line_parts) < 5:
            continue
        if line_parts[0].strip() not in ZIFFERN:
            continue

        ziffer        = line_parts[0]
        vin           = line_parts[1]
        date          = line_parts[2]
        model         = " ".join(line_parts[3:-2])
        total         = line_parts[-1].replace("CHF", "").strip()
        raw_invoice   = lines[i + 1].split()[0] if (i + 1) < len(lines) else ""
        invoice_nr    = raw_invoice.split("&")[0].split("/")[0]

        # Location / Faktor line — remember position for Nebenkosten search
        loadingcity = delivercity = faktor = ansatz = ""
        faktor_j = i  # fallback
        for j in range(i + 1, len(lines)):
            candidate = lines[j]
            if candidate.startswith("CH") and "Faktor" in candidate and "Ansatz" in candidate:
                pattern = r'CH\s+\S+\s+(.*?)\s+CH\s+\S+\s+(.*?)\s+Faktor\s+(\S+)\s+Ansatz\s+(\S+)'
                match = re.search(pattern, candidate)
                if match:
                    loadingcity = match.group(1)
                    delivercity = match.group(2)
                    faktor      = match.group(3)
                    ansatz      = match.group(4)
                faktor_j = j
                break

        # Nebenkosten fields — search from Faktor line position
        def find_amount(keyword, window=8):
            for j in range(faktor_j + 1, min(faktor_j + window, len(lines))):
                ln = lines[j]
                parts = ln.split()
                if parts and parts[0] in ZIFFERN:
                    break
                m = re.search(re.escape(keyword) + r'(?:\s+CHF)?\s+([\d,.]+)', ln)
                if m:
                    return m.group(1)
            return ""

        # LEERFAHRT: appears in InvoiceNr line or nearby before Faktor line
        leerfahrt = ""
        for j in range(i + 1, min(faktor_j + 2, len(lines))):
            if "/ LEERFAHRT" in lines[j]:
                leerfahrt = "OK"
                break

        auktion_protokoll = find_amount("Car Auktion Protokoll")
        terminverein      = find_amount("Terminverein. Absender CarAukt")
        seilwinde_zuschlag = find_amount("Seilwinde-Zuschlag")
        terminzuschlag    = find_amount("Terminzuschlag")
        EFahrzeug         = find_amount("E-Fahrzeug")

        # Auftraggeber
        inv_len = len(invoice_nr)
        if inv_len == 6:
            auftraggeber = "CA3"
        elif inv_len == 5:
            auftraggeber = "RRM"
        else:
            auftraggeber = "Fehler"

        cars_data.append({
            "InvoiceNr":                    invoice_nr,
            "Invoiceshort":                 invoice_nr,
            "Auftraggeber":                 auftraggeber,
            "VIN":                          vin,
            "Model":                        model,
            "Faktor":                       faktor,
            "Ansatz":                       ansatz,
            "Total":                        total,
            "Loadingcity":                  loadingcity,
            "Delivercity":                  delivercity,
            "LEERFAHRT":                    leerfahrt,
            "Car Auktion Protokoll":        auktion_protokoll,
            "Terminverein. Absender CarAukt": terminverein,
            "Seilwinde":                    seilwinde_zuschlag,
            "Terminzuschlag":               terminzuschlag,
            "E-Fahrzeug":                   EFahrzeug,
        })

df2 = pd.DataFrame(cars_data)

# Numeric cleanup
def to_float(v):
    try:
        if pd.isna(v) or str(v).strip() == "":
            return None
        s = str(v).replace("\xa0", "").strip()
        if ',' in s and '.' not in s:
            s = s.replace(',', '.')
        elif '.' in s and ',' in s:
            s = s.replace('.', '').replace(',', '.')
        return round(float(s), 2)
    except Exception:
        return None

for col in ["Faktor", "Ansatz", "Total"]:
    df2[col] = df2[col].apply(to_float)

# Load DB
ca3_url = os.getenv("CA3_URL")
rrm_url = os.getenv("RRM_URL")

if not ca3_url or not rrm_url:
    config_path = os.path.join(os.getcwd(), "config.json")
    if os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
        ca3_url = ca3_url or config.get("CA3_URL", "")
        rrm_url = rrm_url or config.get("RRM_URL", "")

db_ca3 = pd.read_json(ca3_url)
db_rrm = pd.read_json(rrm_url)

# Helper functions (same as pdf3)
def is_empty(v):
    if v is None:
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    try:
        if isinstance(v, float) and math.isnan(v):
            return True
    except Exception:
        pass
    return False

def compare_text_values(v1, v2):
    def clean(val):
        if pd.isna(val):
            return ""
        s = str(val).replace("\xa0", " ").strip()
        if s.endswith(".0"):
            s = s[:-2]
        return s.lower()
    return "OK" if clean(v1) == clean(v2) else "NOK"

def longest_common_substring(s1, s2):
    m, n = len(s1), len(s2)
    dp = [[0] * (n + 1) for _ in range(m + 1)]
    longest = 0
    for i in range(1, m + 1):
        for j in range(1, n + 1):
            if s1[i-1] == s2[j-1]:
                dp[i][j] = dp[i-1][j-1] + 1
                longest = max(longest, dp[i][j])
            else:
                dp[i][j] = 0
    return longest

def compare_city(val1, val2):
    s1 = "" if pd.isna(val1) else str(val1).replace("\xa0", "").strip().lower()
    s2 = "" if pd.isna(val2) else str(val2).replace("\xa0", "").strip().lower()
    if s1 == "" and s2 == "":
        return "OK"
    if (s1 == "nebikon" and s2 == "altishofen") or (s1 == "altishofen" and s2 == "nebikon"):
        return "OK"
    if (longest_common_substring(s1, s2) >= 3
            or (longest_common_substring(s1, s2) >= 2 and s1 == "au")
            or (longest_common_substring(s1, s2) >= 2 and s2 == "au")):
        return "OK"
    return "NOK"

def compare_null_logic(val1, val2):
    if is_empty(val1) and is_empty(val2):
        return "OK"
    elif not is_empty(val1) and not is_empty(val2):
        return "OK"
    return "NOK"

def get_ref(invoice_short):
    s = str(invoice_short).strip()
    if len(s) == 5:
        df_ref = db_rrm
    elif len(s) == 6:
        df_ref = db_ca3
    else:
        return None
    rows = df_ref[df_ref["invoice"].astype(str).str.strip() == s]
    return rows.iloc[0] if len(rows) > 0 else None

# Comparison loop → df_vergleich
compare_results = []

for _, row in df2.iterrows():
    ref = get_ref(row["Invoiceshort"])

    if ref is not None:
        auftraggeber_vergleich  = compare_text_values(row["Auftraggeber"], ref["Auftraggeber"])
        vin_vergleich           = compare_text_values(row["VIN"], ref["vin"])
        loadingcity_vergleich   = compare_city(row["Loadingcity"], ref["loadingcity"])
        delivercity_vergleich   = compare_city(row["Delivercity"], ref["delivercity"])
        faktor_vergleich        = compare_null_logic(row["Faktor"], ref.get("Faktor"))

        telavis_vergleich = "OK" if is_empty(row["Terminverein. Absender CarAukt"]) \
            else compare_null_logic(row["Terminverein. Absender CarAukt"], ref.get("Terminzuschlag"))
        seilwinde_vergleich = "OK" if is_empty(row["Seilwinde"]) \
            else compare_null_logic(row["Seilwinde"], ref.get("Seilwinde"))
        seilwinde_transport_vergleich = "OK" if is_empty(row["Seilwinde"]) \
            else compare_null_logic(row["Seilwinde"], ref.get("Seilwindeintransport"))
        terminzuschlag_vergleich = "OK" if is_empty(row["Terminzuschlag"]) \
            else compare_null_logic(row["Terminzuschlag"], ref.get("Terminzuschlag"))
        efahrzeug_vergleich = "OK" if is_empty(row["Car Auktion Protokoll"]) \
            else compare_null_logic(row["Car Auktion Protokoll"], ref.get("EÜbernahme"))
        leerfahrt_vergleich = "OK" if is_empty(row["LEERFAHRT"]) \
            else compare_null_logic(row["LEERFAHRT"], ref.get("Leerfahrt"))
    else:
        auftraggeber_vergleich = vin_vergleich = loadingcity_vergleich = delivercity_vergleich = "NOK"
        faktor_vergleich = telavis_vergleich = seilwinde_vergleich = "NOK"
        seilwinde_transport_vergleich = terminzuschlag_vergleich = "NOK"
        efahrzeug_vergleich = leerfahrt_vergleich = "NOK"

    # Bemerkungen
    bemerkungen = []
    if auftraggeber_vergleich == "NOK":
        bemerkungen.append("Unterschiedliche Auftraggeber")
    if vin_vergleich == "NOK":
        bemerkungen.append("VIN stimmt nicht überein")
    if loadingcity_vergleich == "NOK":
        bemerkungen.append("Unterschiedliche Ladeorte")
    if delivercity_vergleich == "NOK":
        bemerkungen.append("Unterschiedliche Lieferorte")
    if faktor_vergleich == "NOK":
        bemerkungen.append("Etwas stimmt nicht mit dem Transportauftrag und WFPs 9010, 9020, 2010. Bitte prüfen")
    if telavis_vergleich == "NOK":
        tv = "" if pd.isna(row["Terminverein. Absender CarAukt"]) else str(row["Terminverein. Absender CarAukt"]).strip()
        rv = "" if ref is None or pd.isna(ref.get("Terminzuschlag", "")) else str(ref.get("Terminzuschlag")).strip()
        if tv == "" and rv != "":
            bemerkungen.append("Terminvereinbaren ist vorhanden nur auf CA3 (RRM)")
        elif tv != "" and rv == "":
            bemerkungen.append("Terminvereinbaren ist vorhanden nur in der Gallikerrechnung")
        else:
            bemerkungen.append("Terminvereinbarung weicht ab")
    if seilwinde_vergleich == "NOK":
        sv = "" if pd.isna(row["Seilwinde"]) else str(row["Seilwinde"]).strip()
        rv = "" if ref is None or pd.isna(ref.get("Seilwinde", "")) else str(ref.get("Seilwinde")).strip()
        if seilwinde_transport_vergleich == "OK":
            if sv != "" and rv == "":
                bemerkungen.append("Seilwinde ist vorhanden nur in der Gallikerrechnung. Mit hoher Wahrscheinlichkeit ist der Seilwindepreis schon im Transportpreis berücksichtigt")
            elif sv == "" and rv != "":
                bemerkungen.append("Seilwinde ist vorhanden nur auf CA3 (RRM)")
            else:
                bemerkungen.append("Seilwinde unterschiedlich")
        else:
            if sv == "" and rv != "":
                bemerkungen.append("Seilwinde ist vorhanden nur auf CA3 (RRM)")
            elif sv != "" and rv == "":
                bemerkungen.append("Seilwinde ist vorhanden nur in der Gallikerrechnung")
            else:
                bemerkungen.append("Seilwinde unterschiedlich")
    if terminzuschlag_vergleich == "NOK":
        tv = "" if pd.isna(row["Terminzuschlag"]) else str(row["Terminzuschlag"]).strip()
        rv = "" if ref is None or pd.isna(ref.get("Terminzuschlag", "")) else str(ref.get("Terminzuschlag")).strip()
        if tv == "" and rv != "":
            bemerkungen.append("Terminzuschlag ist vorhanden nur auf CA3 (RRM)")
        elif tv != "" and rv == "":
            bemerkungen.append("Terminzuschlag ist vorhanden nur in der Gallikerrechnung")
        else:
            bemerkungen.append("Terminzuschlag weicht ab")
    if efahrzeug_vergleich == "NOK":
        ev = "" if pd.isna(row["Car Auktion Protokoll"]) else str(row["Car Auktion Protokoll"]).strip()
        rv = "" if ref is None or pd.isna(ref.get("EÜbernahme", "")) else str(ref.get("EÜbernahme")).strip()
        if ev == "" and rv != "":
            bemerkungen.append("E-Fahrzeug ist vorhanden nur auf CA3 (RRM)")
        elif ev != "" and rv == "":
            bemerkungen.append("E-Fahrzeug ist vorhanden nur in der Gallikerrechnung")
        else:
            bemerkungen.append("E-Fahrzeug unterschiedlich")
    if leerfahrt_vergleich == "NOK":
        lv = "" if pd.isna(row["LEERFAHRT"]) else str(row["LEERFAHRT"]).strip()
        rv = "" if ref is None or pd.isna(ref.get("Leerfahrt", "")) else str(ref.get("Leerfahrt")).strip()
        if lv == "" and rv != "":
            bemerkungen.append("Leerfahrt ist vorhanden nur auf CA3 (RRM)")
        elif lv != "" and rv == "":
            bemerkungen.append("Leerfahrt ist vorhanden nur in der Gallikerrechnung. Bitte den 'Comment' in WFP 2900 prüfen")
        else:
            bemerkungen.append("Leerfahrt bitte prüfen")

    compare_results.append({
        "InvoiceNr":                 row["InvoiceNr"],
        "Invoiceshort":              row["Invoiceshort"],
        "Auftraggeber":              row["Auftraggeber"],
        "VIN":                       row["VIN"],
        "Model":                     row["Model"],
        "Faktor":                    row["Faktor"],
        "Total":                     row["Total"],
        "Loadingcity":               row["Loadingcity"],
        "Delivercity":               row["Delivercity"],
        "AuftraggeberVergleich":     auftraggeber_vergleich,
        "VINVergleich":              vin_vergleich,
        "LoadingcityVergleich":      loadingcity_vergleich,
        "DelivercityVergleich":      delivercity_vergleich,
        "TransportWFP_Vergleich":    faktor_vergleich,
        "Bemerkungen":               ", ".join(bemerkungen),
    })

df_vergleich = pd.DataFrame(compare_results)

# STEP 1: Anzahl TA
ca3_count   = df_vergleich[(df_vergleich["Auftraggeber"] == "CA3") & (df_vergleich["AuftraggeberVergleich"] == "OK")].shape[0]
rrm_count   = df_vergleich[(df_vergleich["Auftraggeber"] == "RRM") & (df_vergleich["AuftraggeberVergleich"] == "OK")].shape[0]
fehler_count = df_vergleich[df_vergleich["AuftraggeberVergleich"] == "NOK"].shape[0]

step1_df = pd.DataFrame({"CA3": [ca3_count], "RRM": [rrm_count], "Fehler": [fehler_count]})

cols_step1 = ["Invoiceshort", "VIN", "Model", "Faktor", "Total", "Loadingcity", "Delivercity", "Auftraggeber", "AuftraggeberVergleich"]
dv = df_vergleich[cols_step1]

step_1_1_df = dv[dv["AuftraggeberVergleich"] == "NOK"]
step_1_2_df = dv[(dv["Auftraggeber"] == "CA3") & (dv["AuftraggeberVergleich"] == "OK")]
step_1_3_df = dv[(dv["Auftraggeber"] == "RRM") & (dv["AuftraggeberVergleich"] == "OK")]

step_1_4_df = df_vergleich[
    (df_vergleich["AuftraggeberVergleich"] == "OK") &
    (
        (df_vergleich["VINVergleich"]           == "NOK") |
        (df_vergleich["LoadingcityVergleich"]   == "NOK") |
        (df_vergleich["DelivercityVergleich"]   == "NOK") |
        (df_vergleich["TransportWFP_Vergleich"] == "NOK")
    )
]

# STEP 2: Falschbeträge
valid_totals = [126, 156.8, 189, 207.75, 235.20, 294.35,
                311.65, 311.6, 313.60, 378, 392, 415.5,
                127.25, 158.35, 190.90, 190.85, 209.85, 237.55, 237.5,
                314.80, 314.75, 395.90, 395.85, 419.70, 297.30, 381.80,
                385.95, 390.70, 416.95]

df2_check = df2.copy()
df2_check["Betrag okey"] = df2_check["Total"].round(2).isin(
    [round(v, 2) for v in valid_totals]
).map({True: "OK", False: "NOK"})

step2_df = df2_check[df2_check["Betrag okey"] == "NOK"][[
    "InvoiceNr", "Invoiceshort", "Auftraggeber", "VIN", "Model", "Faktor", "Total", "Betrag okey"
]]

# ── STEP 3: Nebenkosten summary ───────────────────────────────────────────────
def count_nonempty(series):
    return series.apply(lambda v: not is_empty(v)).sum()

leerfahrt_count   = (df2["LEERFAHRT"] == "OK").sum()
seilwinde_count   = count_nonempty(df2["Seilwinde"])
terminzuschlag_count = count_nonempty(df2["Terminzuschlag"])
eubernahme_count  = count_nonempty(df2["Car Auktion Protokoll"])
terminverein_count = count_nonempty(df2["Terminverein. Absender CarAukt"])
efahrzeug_count   = count_nonempty(df2["E-Fahrzeug"])

step3_df = pd.DataFrame({
    "Kategorie": ["Leerfahrt", "Seilwinde-Zuschlag", "Terminzuschlag", "eÜbernahme (Car Auktion Protokoll)", "Terminverein. Absender CarAukt", "E-Fahrzeug"],
    "Anzahl":    [leerfahrt_count, seilwinde_count, terminzuschlag_count, eubernahme_count, terminverein_count, efahrzeug_count]
})

# ── STEP 4: Weiterverrechnet ──────────────────────────────────────────────────
neben_rows = []
NEBEN_MAP = [
    ("LEERFAHRT",                      "Leerfahrt"),
    ("Car Auktion Protokoll",          "eÜbernahme"),
    ("Terminverein. Absender CarAukt", "Terminverein"),
    ("Seilwinde",                      "Seilwinde-Zuschlag"),
    ("Terminzuschlag",                 "Terminzuschlag"),
    ("E-Fahrzeug",                     "E-Fahrzeug"),
]

for _, row in df2.iterrows():
    for col, label in NEBEN_MAP:
        val = row[col]
        if is_empty(val) or val == "":
            continue
        ref = get_ref(row["Invoiceshort"])
        
        weiterverrechnet = "Nein, bitte manuell prüfen"
        if ref is not None:
            if label == "Leerfahrt":
                if not is_empty(ref.get("Leerfahrt")):
                    weiterverrechnet = "Ja"
            elif label == "Seilwinde-Zuschlag":
                if not is_empty(ref.get("Seilwinde")) or not is_empty(ref.get("Seilwindeintransport")):
                    weiterverrechnet = "Ja"
            elif label == "Terminzuschlag" or label == "Terminverein":
                if not is_empty(ref.get("Terminzuschlag")):
                    weiterverrechnet = "Ja"
            elif label in ("eÜbernahme", "E-Fahrzeug"):
                if not is_empty(ref.get("EÜbernahme")):
                    weiterverrechnet = "Ja"

        neben_rows.append({
            "InvoiceNr":       row["InvoiceNr"],
            "Invoiceshort":    row["Invoiceshort"],
            "Auftraggeber":    row["Auftraggeber"],
            "VIN":             row["VIN"],
            "Nebenkosten":     label,
            "Betrag":          val,
            "Weiterverrechnet": weiterverrechnet
        })

step4_df = pd.DataFrame(neben_rows)

# ── STEP 5: PW / SUV / LNF ───────────────────────────────────────────────────
pw_count  = len(df2[df2["Faktor"] == 1.0])
suv_count = len(df2[df2["Faktor"] == 1.5])
lnf_count = len(df2[df2["Faktor"].isin([2.0, 2.5])])
total_veh = pw_count + suv_count + lnf_count

step5_df = pd.DataFrame({
    "PW": [pw_count], "SUV": [suv_count], "LNF": [lnf_count], "Total": [total_veh]
})
step5_df["Check"] = "OK" if total_veh == len(df2) else "NOK"

# ── STEP 6: Duplikate ─────────────────────────────────────────────────────────
dup_mask = df2.duplicated(subset=["VIN", "Faktor"], keep=False)
if dup_mask.any():
    step6_df = df2[dup_mask][["InvoiceNr", "Invoiceshort", "Auftraggeber", "VIN", "Model", "Faktor", "Total"]].copy()
else:
    step6_df = pd.DataFrame({"Meldung": ["Keine Doppeleinträge gefunden"]})

# ── STEP 7: Gesamtbetrag ──────────────────────────────────────────────────────
# Extract total from last PDF page
last_page_text = pages[-1]
summe_ohne_mwst = ""
for line in last_page_text.split("\n"):
    if "Summe ohne Mwst" in line:
        m = re.search(r"Summe ohne Mwst.*?CHF\s*([\d.,]+)", line)
        if m:
            summe_ohne_mwst = m.group(1).strip()
        break

def convert_german_number(value):
    value = value.replace(".", "").replace(",", ".").replace(" ", "")
    return float(value)

try:
    summe_pdf = round(convert_german_number(summe_ohne_mwst), 2)
except Exception:
    summe_pdf = 0.0

def to_float_neben(v):
    """Konvertiert Nebenkosten-Strings (z.B. '100,00') in float."""
    try:
        if pd.isna(v) or str(v).strip() in ('', 'OK'):
            return 0.0
        return float(str(v).replace(',', '.').replace(' ', ''))
    except Exception:
        return 0.0

# Haupttransporte + alle Nebenkosten summieren
neben_cols = ["Car Auktion Protokoll", "Terminverein. Absender CarAukt",
              "Seilwinde", "Terminzuschlag", "E-Fahrzeug"]
neben_sum = sum(df2[col].apply(to_float_neben).sum() for col in neben_cols)
kalk_sum = round(df2["Total"].sum() + neben_sum, 2)

step7_df = pd.DataFrame({
    "Summe ohne MwSt (PDF)":      [summe_pdf],
    "Kalkulatorische Summe":       [kalk_sum],
    "Stimmt der Gesamtbetrag":     ["Ja" if abs(summe_pdf - kalk_sum) < 0.01 else "Nein"]
})

# ── STEP 8: Komische Transporte ───────────────────────────────────────────────
mask_gallik = (
    df2["Loadingcity"].notna() & df2["Loadingcity"].str.contains("Gallik", case=False, na=False) &
    df2["Delivercity"].notna() & df2["Delivercity"].str.contains("Gallik", case=False, na=False)
)
if mask_gallik.any():
    step8_df = df2[mask_gallik][["InvoiceNr", "Invoiceshort", "Auftraggeber", "VIN", "Model", "Faktor", "Total", "Loadingcity", "Delivercity"]].copy()
else:
    step8_df = pd.DataFrame({"Meldung": ["Keine komischen Transporte gefunden"]})

# ── Write file4.xlsx ──────────────────────────────────────────────────────────
def write_sheet(writer, df, sheet_name, col_width=25):
    df.to_excel(writer, index=False, sheet_name=sheet_name)
    ws = writer.sheets[sheet_name]
    for i in range(1, len(df.columns) + 1):
        ws.column_dimensions[get_column_letter(i)].width = col_width

with pd.ExcelWriter("file4.xlsx", engine="openpyxl") as writer:
    write_sheet(writer, step1_df,    "Step_1 Anzahl TA")
    write_sheet(writer, step_1_1_df, "Step_1_1 Fehler")
    write_sheet(writer, step_1_2_df, "Step_1_2 CA3")
    write_sheet(writer, step_1_3_df, "Step_1_3 RRM")
    write_sheet(writer, step_1_4_df, "Step_1_4 Zusätzlich")
    write_sheet(writer, step2_df,    "Step_2_Falschbeträge")
    write_sheet(writer, step3_df,    "Step_3_Nebenkosten")
    write_sheet(writer, step4_df,    "Step_4_Weiteverrechnet")
    write_sheet(writer, step5_df,    "Step_5_PW_SUV_LNF")
    write_sheet(writer, step6_df,    "Step_6_Dublikate")
    write_sheet(writer, step7_df,    "Step_7_Gesamtbetrag", col_width=35)
    write_sheet(writer, step8_df,    "Step_8_Komische_Transporte")

print("file4.xlsx erfolgreich gespeichert mit 8 Steps.")