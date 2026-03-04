import pandas as pd
import numpy as np
import openpyxl
from openpyxl.utils import get_column_letter
import os
import json
import re

# ── Input ─────────────────────────────────────────────────────────────────────
df_raw = pd.read_excel("Cotra_invoice.xlsx", header=0)

# Drop empty/summary rows
df_raw = df_raw.dropna(subset=["LS"])
df_raw["LS"] = df_raw["LS"].astype(int)

# Mark Leerfahrt: Standard rows where WGR == 'LEER' → treat as Nebenkosten
leer_mask = (df_raw["WGR"].astype(str).str.strip().str.upper() == "LEER") & (df_raw["BEZEICHNUNG"] == "Standard")
df_raw.loc[leer_mask, "BEZEICHNUNG"] = "Leerfahrt"

# ── Load DB ───────────────────────────────────────────────────────────────────
ca3_url = os.getenv("CA3_URL_Excel")
rrm_url = os.getenv("RRM_URL_Excel")

if not ca3_url or not rrm_url:
    config_path = os.path.join(os.getcwd(), "config.json")
    if os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
        ca3_url = ca3_url or config.get("CA3_URL_Excel", "")
        rrm_url = rrm_url or config.get("RRM_URL_Excel", "")

db_ca3 = pd.read_json(ca3_url)
db_rrm = pd.read_json(rrm_url)

# ── Helpers ───────────────────────────────────────────────────────────────────
def norm(v):
    if pd.isna(v):
        return ""
    return str(v).replace("\xa0", " ").strip()

def norm_lower(v):
    return norm(v).lower()

def compare_text(v1, v2):
    def clean(v):
        s = norm(v).lower()
        if s.endswith(".0"):
            s = s[:-2]
        return s
    return "OK" if clean(v1) == clean(v2) else "NOK"

def longest_common_substring(s1, s2):
    m, n = len(s1), len(s2)
    dp = [[0] * (n + 1) for _ in range(m + 1)]
    best = 0
    for i in range(1, m + 1):
        for j in range(1, n + 1):
            if s1[i-1] == s2[j-1]:
                dp[i][j] = dp[i-1][j-1] + 1
                best = max(best, dp[i][j])
    return best

def compare_city(v1, v2):
    s1 = norm_lower(v1)
    s2 = norm_lower(v2)
    if s1 == "" and s2 == "":
        return "OK"
    if {s1, s2} == {"nebikon", "altishofen"}:
        return "OK"
    if (longest_common_substring(s1, s2) >= 3
            or (longest_common_substring(s1, s2) >= 2 and s1 == "au")
            or (longest_common_substring(s1, s2) >= 2 and s2 == "au")):
        return "OK"
    return "NOK"

def null_ok(v1, v2):
    import math
    def empty(v):
        if v is None: return True
        if isinstance(v, str) and v.strip() == "": return True
        try:
            if isinstance(v, float) and math.isnan(v): return True
        except: pass
        return False
    if empty(v1) and empty(v2): return "OK"
    if not empty(v1) and not empty(v2): return "OK"
    return "NOK"

def extract_city(address_str):
    s = norm(address_str)
    m = re.search(r'CH-\d{4}\s+(.+)$', s)
    return m.group(1).strip() if m else s

def extract_plz(address_str):
    s = norm(address_str)
    m = re.search(r'CH-(\d{4})', s)
    return m.group(1) if m else ""

def lookup_vin_all(vin):
    """Returns (DataFrame of all matching rows, source label) or (empty DF, None)."""
    v = norm(vin).upper()
    if not v:
        return pd.DataFrame(), None
    for db, label in [(db_ca3, "CA3"), (db_rrm, "RRM")]:
        match = db[db["vin"].astype(str).str.strip().str.upper() == v]
        if len(match) > 0:
            return match, label
    return pd.DataFrame(), None

def get_invoiceshort(vin, von_str, nach_str):
    """
    Find Invoiceshort from DB by VIN + VON/NACH matching:
    - If NACH contains 'Cotra' → rows where DB Faktor == 2010
    - Otherwise               → rows where DB Faktor in [9010, 9020]
    - Among candidates, pick the row whose loadingcity is in VON
      and delivercity is in NACH (best match).
    - Fallback: first candidate row.
    """
    matches, _ = lookup_vin_all(vin)
    if matches.empty:
        return ""

    nach_lower = norm_lower(nach_str)
    von_lower  = norm_lower(von_str)
    is_cotra_nach = "cotra" in nach_lower

    # Step 1: filter by Faktor code based on NACH contains Cotra
    if is_cotra_nach:
        candidates = matches[matches["Faktor"] == 2010]
    else:
        candidates = matches[matches["Faktor"].isin([9010, 9020])]

    if candidates.empty:
        candidates = matches

    # Step 2: if more than one candidate, narrow down by loadingcity in VON and delivercity in NACH
    if len(candidates) > 1:
        for _, r in candidates.iterrows():
            lc = norm_lower(r.get("loadingcity", ""))
            dc = norm_lower(r.get("delivercity", ""))
            if lc and dc and lc in von_lower and dc in nach_lower:
                val = r.get("invoice")
                return "" if pd.isna(val) else str(val).strip()

    # Fallback: first candidate
    val = candidates.iloc[0]["invoice"]
    return "" if pd.isna(val) else str(val).strip()

# ── Build one row per Standard line (= one transport) ────────────────────────
def get_nebenkosten(group):
    extra = group[~group["BEZEICHNUNG"].isin(["Standard", "Treibstoffzuschlag"])]["BEZEICHNUNG"]
    return list(extra.dropna().unique())

records = []
for ls_id, grp in df_raw.groupby("LS"):
    neben = get_nebenkosten(grp)
    standard_rows = grp[grp["BEZEICHNUNG"] == "Standard"]

    for _, std_row in standard_rows.iterrows():
        vin  = norm(std_row.get("KREF"))
        nach = norm(std_row.get("NACH", ""))
        von  = norm(std_row.get("VON", ""))
        records.append({
            "LS":             ls_id,
            "Invoiceshort":   get_invoiceshort(vin, von, nach),
            "POS_DATUM":      std_row.get("POS_DATUM"),
            "DG":             std_row.get("DG"),
            "TOUR":           norm(std_row.get("TOUR")),
            "VIN":            vin,
            "Model":          norm(std_row.get("SOLL_MEE")),
            "Faktor":         std_row.get("SOLL_M3"),
            "SOLL_PL":        std_row.get("SOLL_PL"),
            "SOLL_TO":        std_row.get("SOLL_TO"),
            "SOLL_M3":        std_row.get("SOLL_M3"),
            "Total":          std_row.get("POSITIONSBETRAG"),
            "VON":            von,
            "NACH":           nach,
            "Loadingcity":    extract_city(von),
            "Loadingzipcode": extract_plz(von),
            "Delivercity":    norm(std_row.get("EMPF_ORT")),
            "Deliverzipcode": str(int(float(norm(std_row.get("EMPF_PLZ"))))) if norm(std_row.get("EMPF_PLZ")) else "",
            "Bezeichnung":    norm(std_row.get("BEZEICHNUNG")),
            "Nebenkosten":    ", ".join(neben) if neben else np.nan,
            "Hat_Seilwinde":  "Zuschlag Seilwinde" in neben,
            "Hat_Wartezeit":  any("Wartezeit" in n for n in neben),
            "Hat_Engadin":    "Zuschlag Engadin" in neben,
            "Hat_Admin":      "Administrationsaufwand" in neben,
        })

new_df = pd.DataFrame(records)
print(f"Transporte gesamt (Standard-Zeilen): {len(new_df)}")

# ── Build full comparison table ───────────────────────────────────────────────
compare_results = []

for _, row in new_df.iterrows():
    matches, source = lookup_vin_all(row["VIN"])

    if not matches.empty:
        # Pick the right DB row based on NACH: Cotra→2010, else→9010/9020
        is_cotra_nach = "cotra" in norm_lower(row["NACH"])
        faktor_col = matches["Faktor"]
        if is_cotra_nach:
            ref_rows = matches[faktor_col == 2010]
        else:
            ref_rows = matches[faktor_col.isin([9010, 9020])]
        ref = ref_rows.iloc[0] if not ref_rows.empty else matches.iloc[0]

        auftraggeber     = source
        auftraggeber_ver = "OK"
        vin_ver          = "OK"
        db_loadingcity   = norm_lower(ref.get("loadingcity", ""))
        db_delivercity   = norm_lower(ref.get("delivercity", ""))
        loadingcity_ver  = "OK" if (db_loadingcity and db_loadingcity in norm_lower(row["VON"])) else "NOK"
        delivercity_ver  = "OK" if (db_delivercity and db_delivercity in norm_lower(row["NACH"])) else "NOK"
        deliver_plz_ver  = compare_text(row["Deliverzipcode"], ref.get("deliverzipcode", ""))
        faktor_ver       = null_ok(row["Faktor"],              ref.get("Faktor", np.nan))
    else:
        ref = None
        auftraggeber     = "Fehler"
        auftraggeber_ver = vin_ver = loadingcity_ver = delivercity_ver = "NOK"
        deliver_plz_ver = faktor_ver = "NOK"

    compare_results.append({
        "LS":                              row["LS"],
        "Invoiceshort":                    row["Invoiceshort"],
        "Auftraggeber":                    auftraggeber,
        "VIN":                             row["VIN"],
        "Model":                           row["Model"],
        "Faktor":                          row["Faktor"],
        "Total":                           row["Total"],
        "SOLL_PL":                         row["SOLL_PL"],
        "SOLL_TO":                         row["SOLL_TO"],
        "SOLL_M3":                         row["SOLL_M3"],
        "Bezeichnung":                     row["Bezeichnung"],
        "Nebenkosten":                     row["Nebenkosten"],
        "Loadingcity":                     row["Loadingcity"],
        "Delivercity":                     row["Delivercity"],
        "Loadingzipcode":                  row["Loadingzipcode"],
        "Deliverzipcode":                  row["Deliverzipcode"],
        "VON":                             row["VON"],
        "NACH":                            row["NACH"],
        "AuftraggeberVergleich":           auftraggeber_ver,
        "VINVergleich":                    vin_ver,

        "DeliverzipVergleich":             deliver_plz_ver,
        "LoadingcityVergleich":            loadingcity_ver,
        "DelivercityVergleich":            delivercity_ver,
        "TransportWFP_aktiviert_Vergleich": faktor_ver,
        "Hat_Seilwinde":                   row["Hat_Seilwinde"],
        "Hat_Wartezeit":                   row["Hat_Wartezeit"],
        "Hat_Engadin":                     row["Hat_Engadin"],
        "Hat_Admin":                       row["Hat_Admin"],
    })

df_vergleich = pd.DataFrame(compare_results)

# ── STEP 1: Anzahl TA ─────────────────────────────────────────────────────────
ca3_count    = df_vergleich[(df_vergleich["Auftraggeber"] == "CA3") & (df_vergleich["AuftraggeberVergleich"] == "OK")].shape[0]
rrm_count    = df_vergleich[(df_vergleich["Auftraggeber"] == "RRM") & (df_vergleich["AuftraggeberVergleich"] == "OK")].shape[0]
fehler_count = df_vergleich[df_vergleich["AuftraggeberVergleich"] == "NOK"].shape[0]

step1_df = pd.DataFrame({"CA3": [ca3_count], "RRM": [rrm_count], "Fehler": [fehler_count]})

# ── STEP 1_1 / 1_2 / 1_3 ─────────────────────────────────────────────────────
# LS and Invoiceshort always first two columns
COLS_S1 = ["LS", "Invoiceshort", "VIN", "Model", "Faktor", "Total",
           "Loadingcity", "Delivercity", "Auftraggeber", "AuftraggeberVergleich"]

step_1_1_df = df_vergleich[
    df_vergleich["AuftraggeberVergleich"] == "NOK"
][COLS_S1].copy()

step_1_2_df = df_vergleich[
    (df_vergleich["Auftraggeber"] == "CA3") & (df_vergleich["AuftraggeberVergleich"] == "OK")
][COLS_S1].copy()

step_1_3_df = df_vergleich[
    (df_vergleich["Auftraggeber"] == "RRM") & (df_vergleich["AuftraggeberVergleich"] == "OK")
][COLS_S1].copy()

# ── STEP 1_4: Zusätzlich ──────────────────────────────────────────────────────
COLS_S14 = [
    "LS", "Invoiceshort", "Auftraggeber", "VIN", "Model", "Faktor", "Total",
    "Loadingcity", "Delivercity", "Loadingzipcode", "Deliverzipcode",
    "AuftraggeberVergleich", "VINVergleich",
    "DeliverzipVergleich",
    "LoadingcityVergleich", "DelivercityVergleich",
    "TransportWFP_aktiviert_Vergleich"
]
df_zusatz = df_vergleich[
    (df_vergleich["AuftraggeberVergleich"] == "OK") &
    (
        (df_vergleich["VINVergleich"]                    == "NOK") |

        (df_vergleich["DeliverzipVergleich"]             == "NOK") |
        (df_vergleich["LoadingcityVergleich"]            == "NOK") |
        (df_vergleich["DelivercityVergleich"]            == "NOK") |
        (df_vergleich["TransportWFP_aktiviert_Vergleich"]== "NOK")
    )
][COLS_S14].copy()

# ── STEP 2: Falschbeträge ─────────────────────────────────────────────────────
# Standard row POSITIONSBETRAG is always just 154 or 214 - Nebenkosten are separate rows
VALID_BASE = {154, 214}

def check_betrag(row):
    total = row["Total"]
    if pd.isna(total):
        return "NOK"
    total = round(float(total), 2)
    return "OK" if total in {round(v, 2) for v in VALID_BASE} else "NOK"

df_vergleich["Betrag okey"] = df_vergleich.apply(check_betrag, axis=1)

# Additional check: SOLL_PL * SOLL_TO must equal SOLL_M3
def check_faktor(row):
    try:
        pl = float(row["SOLL_PL"])
        to = float(row["SOLL_TO"])
        m3 = float(row["SOLL_M3"])
        if round(m3, 4) > round(pl * to, 4):
            return "NOK, Faktor prüfen"
    except:
        pass
    return row["Betrag okey"]

df_vergleich["Betrag okey"] = df_vergleich.apply(
    lambda r: check_faktor(r) if r["Betrag okey"] == "OK" else r["Betrag okey"], axis=1
)

COLS_S2 = ["LS", "Invoiceshort", "Auftraggeber", "VIN", "Model",
           "SOLL_PL", "SOLL_TO", "SOLL_M3", "Faktor", "Bezeichnung",
           "Nebenkosten", "Total", "Loadingcity", "Delivercity",
           "Loadingzipcode", "Deliverzipcode", "VON", "NACH", "Betrag okey"]

step2_df_errors = df_vergleich[
    df_vergleich["Betrag okey"].isin(["NOK", "Manuell prüfen", "NOK, Faktor prüfen"])
][COLS_S2].copy()

# ── STEP 3: Nebenkosten Zusammenfassung ───────────────────────────────────────
# Count directly from df_raw Nebenkosten lines (not df_vergleich which may have duplicates)
df_neben_raw = df_raw[~df_raw["BEZEICHNUNG"].isin(["Standard", "Treibstoffzuschlag"])]

known_mask = (
    (df_neben_raw["BEZEICHNUNG"] == "Leerfahrt") |
    (df_neben_raw["BEZEICHNUNG"] == "Zuschlag Seilwinde") |
    (df_neben_raw["BEZEICHNUNG"].str.contains("Wartezeit", na=False)) |
    (df_neben_raw["BEZEICHNUNG"] == "Zuschlag Engadin") |
    (df_neben_raw["BEZEICHNUNG"] == "Administrationsaufwand")
)
andere_count = int((~known_mask).sum())

summary_nebenkosten = pd.DataFrame({
    "Kategorie": ["Leerfahrt", "Zuschlag Seilwinde", "Wartezeit", "Zuschlag Engadin", "Administrationsaufwand", "Andere"],
    "Anzahl":    [
        int((df_neben_raw["BEZEICHNUNG"] == "Leerfahrt").sum()),
        int((df_neben_raw["BEZEICHNUNG"] == "Zuschlag Seilwinde").sum()),
        int(df_neben_raw["BEZEICHNUNG"].str.contains("Wartezeit", na=False).sum()),
        int((df_neben_raw["BEZEICHNUNG"] == "Zuschlag Engadin").sum()),
        int((df_neben_raw["BEZEICHNUNG"] == "Administrationsaufwand").sum()),
        andere_count,
    ]
})

# ── STEP 4: Weiterverrechnet ──────────────────────────────────────────────────
# One row per Nebenkosten line (BEZEICHNUNG != Standard/Treibstoffzuschlag)
df_raw_neben = df_raw[~df_raw["BEZEICHNUNG"].isin(["Standard", "Treibstoffzuschlag"])].copy()

# Join Invoiceshort and Auftraggeber from df_vergleich by LS
# Use drop_duplicates to avoid duplicating Nebenkosten rows when LS has 2 Standard lines
ls_meta = df_vergleich[["LS", "Invoiceshort", "Auftraggeber", "Faktor"]].drop_duplicates(subset=["LS"]).copy()
df_raw_neben["LS"] = df_raw_neben["LS"].astype(int)
step_4_neben = df_raw_neben.merge(ls_meta, on="LS", how="left")

# VIN always taken directly from KREF (works even for Leerfahrt-only LS)
step_4_neben["VIN"] = step_4_neben["KREF"].apply(norm)

# Fill Invoiceshort for rows where merge found nothing (Leerfahrt-only LS)
leer_mask = step_4_neben["Invoiceshort"].isna()
if leer_mask.any():
    step_4_neben.loc[leer_mask, "Invoiceshort"] = step_4_neben.loc[leer_mask].apply(
        lambda r: get_invoiceshort(norm(r["KREF"]), norm(r["VON"]), norm(r["NACH"])), axis=1
    )
    # Auftraggeber for Leerfahrt-only LS: lookup from DB
    def get_auftraggeber(vin):
        _, source = lookup_vin_all(vin)
        return source if source else "Fehler"
    step_4_neben.loc[leer_mask, "Auftraggeber"] = step_4_neben.loc[leer_mask, "VIN"].apply(get_auftraggeber)

step_4_neben = step_4_neben[[
    "LS", "Invoiceshort", "Auftraggeber", "VIN", "Faktor",
    "BEZEICHNUNG", "POSITIONSBETRAG"
]].rename(columns={
    "BEZEICHNUNG":    "Nebenkosten",
    "POSITIONSBETRAG": "Nebenkosten_Betrag"
}).copy()

for col in ["Transport_WFP", "Seilwinde", "Terminzuschlag",
            "EÜbernahme", "Leerfahrt", "Seilwindeintransport"]:
    step_4_neben[col] = None

step_4_neben["Weiterverrechnet"] = ""

for idx, row in step_4_neben.iterrows():
    matches, _ = lookup_vin_all(row["VIN"])
    ref = matches.iloc[0] if not matches.empty else None

    if ref is not None:
        step_4_neben.at[idx, "Transport_WFP"]       = ref.get("Faktor",               np.nan)
        step_4_neben.at[idx, "Seilwinde"]            = ref.get("Seilwinde",            np.nan)
        step_4_neben.at[idx, "Terminzuschlag"]       = ref.get("Terminzuschlag",       np.nan)
        step_4_neben.at[idx, "EÜbernahme"]           = ref.get("EÜbernahme",           np.nan)
        step_4_neben.at[idx, "Leerfahrt"]            = ref.get("Leerfahrt",            np.nan)
        step_4_neben.at[idx, "Seilwindeintransport"] = ref.get("Seilwindeintransport", np.nan)

    neben   = str(row["Nebenkosten"]) if pd.notna(row["Nebenkosten"]) else ""
    betrag  = round(float(row["Nebenkosten_Betrag"]), 2) if pd.notna(row["Nebenkosten_Betrag"]) else None
    wv = "Nein, bitte manuell prüfen"

    if neben == "Leerfahrt":
        leer_val = ref.get("Leerfahrt") if ref is not None else None
        if betrag is not None and pd.notna(leer_val):
            wv = "Ja"
    elif "Seilwinde" in neben:
        seil_val = ref.get("Seilwinde")            if ref is not None else None
        seil_tr  = ref.get("Seilwindeintransport") if ref is not None else None
        if betrag is not None and (pd.notna(seil_val) or pd.notna(seil_tr)):
            wv = "Ja"
    elif "Wartezeit" in neben or "Administrationsaufwand" in neben:
        wv = "Manuell prüfen"

    step_4_neben.at[idx, "Weiterverrechnet"] = wv

# ── STEP 5: PW / SUV / LNF ───────────────────────────────────────────────────
pw_count    = len(df_vergleich[df_vergleich["Faktor"] <= 1.0])
lnf_count   = len(df_vergleich[df_vergleich["Faktor"] >= 2.0])
suv_count   = len(df_vergleich[(df_vergleich["Faktor"] > 1.0) & (df_vergleich["Faktor"] < 2.0)])
total_count = pw_count + suv_count + lnf_count
step1_total = ca3_count + rrm_count + fehler_count

df_5 = pd.DataFrame({
    "PW":    [pw_count],
    "SUV":   [suv_count],
    "LNF":   [lnf_count],
    "Total": [total_count],
    "Check": ["OK" if total_count == step1_total else "NOK"]
})

# ── STEP 6: Dublikate ─────────────────────────────────────────────────────────
# A real duplicate = same VIN (KREF) + same VON + same NACH + same BEZEICHNUNG
# across different LS entries (not just Nebenkosten rows of the same transport)
# Duplicates: same VIN + VON + NACH + BEZEICHNUNG across all df_raw rows
dup_mask = df_raw.duplicated(subset=["KREF", "VON", "NACH", "BEZEICHNUNG"], keep=False)
df_dups_raw = df_raw[dup_mask].copy()

if not df_dups_raw.empty:
    df_duplicates = df_dups_raw[[
        "LS", "KREF", "SOLL_MEE", "SOLL_PL", "POSITIONSBETRAG", "VON", "NACH", "BEZEICHNUNG"
    ]].rename(columns={
        "KREF":           "VIN",
        "SOLL_MEE":       "Model",
        "SOLL_PL":        "Faktor",
        "POSITIONSBETRAG": "Betrag"
    }).sort_values(["VIN", "VON", "NACH", "BEZEICHNUNG"])
    print(f"Gefunden: {len(df_duplicates)} Duplikat-Zeilen")
else:
    df_duplicates = pd.DataFrame({"Meldung": ["Keine Doppeleinträge gefunden"]})
    print("Keine Doppeleinträge gefunden")

# ── STEP 7: Gesamtbetrag ──────────────────────────────────────────────────────
kalk_sum = round(df_raw["POSITIONSBETRAG"].sum(), 2)

df_7 = pd.DataFrame({
    "Invoice":                      ["Cotra Transportrechnung"],
    "Total_invoice":                [np.nan],
    "Kalkulatorische_Betragssumme": [kalk_sum],
    "Stimmt der Gesamtbetrag":      ["Manuell prüfen – kein Gesamtbetrag in Datei"]
})

# ── STEP 8: Komische Transporte ───────────────────────────────────────────────
mask_von  = df_vergleich["VON"].str.lower().str.contains("cotra", na=False)
mask_nach = df_vergleich["NACH"].str.lower().str.contains("cotra", na=False)

if (mask_von & mask_nach).any():
    df_8 = df_vergleich[mask_von & mask_nach][[
        "LS", "Invoiceshort", "VIN", "Model", "Faktor",
        "Total", "VON", "NACH", "Auftraggeber"
    ]].copy()
else:
    df_8 = pd.DataFrame({"Meldung": ['Keine "komischen" Transporte gefunden']})

# ── Write output ──────────────────────────────────────────────────────────────
def write_sheet(writer, df_sheet, sheet_name, col_width=25):
    df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
    ws = writer.sheets[sheet_name]
    for i in range(1, len(df_sheet.columns) + 1):
        ws.column_dimensions[get_column_letter(i)].width = col_width

with pd.ExcelWriter("file4.xlsx", engine="openpyxl") as writer:
    write_sheet(writer, step1_df,            "Step_1 Anzahl TA",           col_width=25)
    write_sheet(writer, step_1_1_df,         "Step_1_1 Fehler",            col_width=25)
    write_sheet(writer, step_1_2_df,         "Step_1_2 CA3",               col_width=25)
    write_sheet(writer, step_1_3_df,         "Step_1_3 RRM",               col_width=25)
    write_sheet(writer, df_zusatz,           "Step_1_4 Zusätzlich",        col_width=25)
    write_sheet(writer, step2_df_errors,     "Step_2_Falschbeträge",       col_width=25)
    write_sheet(writer, summary_nebenkosten, "Step_3_Nebenkosten",         col_width=30)
    write_sheet(writer, step_4_neben,        "Step_4_Weiteverrechnet",     col_width=25)
    write_sheet(writer, df_5,                "Step_5_PW_SUV_LNF",          col_width=25)
    write_sheet(writer, df_duplicates,       "Step_6_Dublikate",           col_width=30)
    write_sheet(writer, df_7,                "Step_7_Gesamtbetrag",        col_width=35)
    write_sheet(writer, df_8,                "Step_8_Komische_Transporte", col_width=30)

print("file4.xlsx erfolgreich gespeichert.")