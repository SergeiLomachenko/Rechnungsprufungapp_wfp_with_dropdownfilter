import re
import pdfplumber
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
import os 
import json

# PDF-path
excel_path = "invoice.xlsx"

df = pd.read_excel("invoice.xlsx", skiprows=5)

df_temp = df.replace(r'^\s*$', pd.NA, regex=True)
empty_rows = df_temp.isna().all(axis=1) 
if empty_rows.any():
    first_empty_index = empty_rows.idxmax()
    df = df.loc[:first_empty_index - 1]
else:
    pass
df = df.reset_index(drop=True)
df = df.dropna(how='all')

# changes to PLZ
def clean_plz(series):
    numeric = pd.to_numeric(series, errors='coerce').fillna(0)
    # making to int, then to string
    return numeric.astype(int).astype(str)
df["Absender PLZ"] = clean_plz(df["Absender PLZ"])
df["Empfaenger PLZ"] = clean_plz(df["Empfaenger PLZ"])

for col in ["Ansatz", "Faktor", "Betrag"]:
    df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', '.'), errors='coerce')

invoice_nr = (
    df['Ordernummer']
    .fillna('')
    .astype(str)
    .str.extract(r'^\s*([^\s/]+)')  # extract all till the first backspase
    .squeeze()
    .str.strip()
)
lengths = invoice_nr.str.len()

# Auftraggeber kalkulieren
auftraggeber = lengths.map({
    5: "RRM",
    6: "CA3"
}).fillna("Fehler")

new_df = pd.DataFrame({
    "InvoiceNr": df["Ordernummer"],
    "Invoiceshort" : invoice_nr,
    'Auftraggeber': auftraggeber,
    'VIN': df['Fahrgestellnummer'], 
    'Model': df['Fahrzeugtyp'], 
    "Ansatz": df["Ansatz"],
    "Faktor": df["Faktor"],
    "Nebenkosten": df["Nebenkosten"],
    "Total": df["Betrag"],
    "Loadingcity": df["Absenderort"],
    "Delivercity": df["Empfaengerort"],
    "Absender" : df["Absender"],
    "Empfaenger": df["Empfaenger"],
    "Absender PLZ" : df["Absender PLZ"],
    "Empfaenger PLZ": df["Empfaenger PLZ"]
})

nebenkosten_df = new_df[new_df['Nebenkosten'].notna()]

# loading envs
ca3_daten = os.getenv("CA3_URL_Excel")
rrm_daten = os.getenv("RRM_URL_Excel")

if not ca3_daten:
    config_path = os.path.join(os.getcwd(), "config.json")
    if os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
        ca3_daten = config.get("CA3_URL_Excel", "")

if not rrm_daten:
    config_path = os.path.join(os.getcwd(), "config.json")
    if os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
        rrm_daten = config.get("RRM_URL_Excel", "")

df_ca3 = pd.read_json(ca3_daten)
df_rrm = pd.read_json(rrm_daten)

# Below are all the functions that were used to compare the invoice-data with metabase-database
def compare_text_values(v1, v2):
    """
    Vergleicht zwei Textwerte.
    Fehlt ein Wert (pd.isna), wird er als leerer String berücksichtigt.
    Zusätzlich werden unsichtbare Leerzeichen entfernt.
    Beide Werte werden getrimmt und in Kleinbuchstaben umgewandelt.
    """
    def clean(val):
        if pd.isna(val):
            return ""
        # Приводим к строке, убираем неразрывные пробелы и обычные пробелы по краям
        s = str(val).replace("\xa0", " ").strip()
        # Убираем возможные .0 у целых чисел (если были float)
        if s.endswith(".0"):
            s = s[:-2]
        return s.lower()

    s1 = clean(v1)
    s2 = clean(v2)
    return "OK" if s1 == s2 else "NOK"



# def compare_text_values(v1, v2):
#     """
#     Vergleicht zwei Textwerte.
#     Fehlt ein Wert (pd.isna), wird er als leerer String berücksichtigt.
#     Zusätzlich werden unsichtbare Leerzeichen entfernt.
#     Beide Werte werden getrimmt und in Kleinbuchstaben umgewandelt.
#     """
#     s1 = "" if pd.isna(v1) else str(v1).replace("\xa0", "").strip().lower()
#     s2 = "" if pd.isna(v2) else str(v2).replace("\xa0", "").strip().lower()
#     return "OK" if s1 == s2 else "NOK"

def longest_common_substring(s1, s2):
    m, n = len(s1), len(s2)
    dp = [[0]*(n+1) for _ in range(m+1)]
    longest = 0
    for i in range(1, m+1):
        for j in range(1, n+1):
            if s1[i-1] == s2[j-1]:
                dp[i][j] = dp[i-1][j-1] + 1
                if dp[i][j] > longest:
                    longest = dp[i][j]
            else:
                dp[i][j] = 0
    return longest

def compare_city(val1, val2):
    # Konvertiere die Eingabewerte in Kleinbuchstaben,
    # entferne unsichtbare Leerzeichen (z. B. non-breaking spaces) und trimme den Whitespace.
    s1 = "" if pd.isna(val1) else str(val1).replace("\xa0", "").strip().lower()
    s2 = "" if pd.isna(val2) else str(val2).replace("\xa0", "").strip().lower()
    
    # Falls beide Werte leer sind, werden sie als gleich angesehen.
    if s1 == "" and s2 == "":
        return "OK"
    
    # Sonderregel: Wenn ein Wert "nebikon" und der andere "altishofen" (oder umgekehrt) ist,
    # werden diese als gleich betrachtet.
    if (s1 == "nebikon" and s2 == "altishofen") or (s1 == "altishofen" and s2 == "nebikon"):
        return "OK"
    
    # Wenn die Länge der längsten gemeinsamen Teilzeichenkette mindestens 5 Zeichen beträgt,
    # werden die Werte als ähnlich betrachtet.
    if (
        longest_common_substring(s1, s2) >= 3 
        or (longest_common_substring(s1, s2) >= 2 and s1 == "au") 
        or (longest_common_substring(s1, s2) >= 2 and s2 == "au")
        ):
        return "OK"
    
    # Ansonsten gelten die Werte als unterschiedlich.
    return "NOK"
# End of all the functions that were used to compare the invoice-data with metabase-database

# Liste zur Speicherung der Vergleichsergebnisse erstellen
compare_results = []

for idx, row in new_df.iterrows():
    Invoiceshort = str(row["Invoiceshort"]).strip()
    
    # Auswahl der Referenzdatei anhand der Länge der InvoiceNr:
    # - 5-stellige InvoiceNr → rrm.xlsx
    # - 6-stellige InvoiceNr → ca3.xlsx
    if len(Invoiceshort) == 5:
        df_reference = df_rrm
    elif len(Invoiceshort) == 6:
        df_reference = df_ca3
    else:
        df_reference = None

    # Suche der passenden Zeile in der Referenzdatei anhand der Spalte "invoice"
    if df_reference is not None:
        matching_rows = df_reference[df_reference["invoice"].astype(str).str.strip() == Invoiceshort]
        if len(matching_rows) > 0:
            ref = matching_rows.iloc[0]
        else:
            ref = None
    else:
        ref = None

    # Falls eine Referenzzeile gefunden wurde, erfolgt der Vergleich.
    def compare_null_logic(val1, val2):
        # считаем пустыми: None, "", пробелы, NaN
        def is_empty(v):
            if v is None:
                return True
            if isinstance(v, str) and v.strip() == "":
                return True
            try:
                import math
                if isinstance(v, float) and math.isnan(v):
                    return True
            except:
                pass
            return False

        if is_empty(val1) and is_empty(val2):
            return "OK"
        elif not is_empty(val1) and not is_empty(val2):
            return "OK"
        else:
            return "NOK"
    

    if ref is not None:
        auftraggeber_vergleich = compare_text_values(row["Auftraggeber"], ref["Auftraggeber"]) 
        vin_vergleich = compare_text_values(row["VIN"], ref["vin"]) 
        loadingcity_vergleich = compare_city(row["Loadingcity"], ref["loadingcity"]) 
        delivercity_vergleich = compare_city(row["Delivercity"], ref["delivercity"])
        loadingzipcode_vergleich = compare_text_values(row["Absender PLZ"], ref["loadingzipcode"])
        deliverzipcode_vergleich = compare_text_values(row["Empfaenger PLZ"], ref["deliverzipcode"])
        faktor_vergleich = compare_null_logic(row["Faktor"], ref["Faktor"])
        transportrpeis_vergleich = compare_null_logic(row["Total"], ref["Faktor"])

        # telavis_vergleich
        # def is_emptytel(val):
        #     return val is None or (isinstance(val, str) and val.strip() == "")
        # if is_emptytel(row["Terminverein. Absender CarAukt"]):
        #     telavis_vergleich = "OK"
        # else:
        #     telavis_vergleich = compare_null_logic(new_df["Terminverein. Absender CarAukt"], ref["Terminzuschlag"])

        # seilwinde_vergleich
        # def is_emptyseil(val):
        #     return val is None or (isinstance(val, str) and val.strip() == "")
        # if is_emptyseil(row["Seilwinde"]):
        #     seilwinde_vergleich = "OK"
        # else:
        #     seilwinde_vergleich = compare_null_logic(new_df["Seilwinde"], ref["Seilwinde"])
        
        # compare Seilwindeintransport
        # def is_emptyseiltr(val):
        #     return val is None or (isinstance(val, str) and val.strip() == "")
        # if is_emptyseiltr(new_df["Seilwinde"]):
        #     seilwinde_transport_vergleich = "OK"
        # else:
        #     seilwinde_transport_vergleich = compare_null_logic(new_df["Seilwinde"], ref["Seilwindeintransport"])

        # terminzuschlag_vergleich
        # def is_emptytermin(val):
        #     return val is None or (isinstance(val, str) and val.strip() == "")
        # if is_emptytermin(new_df["Terminzuschlag"]):
        #     terminzuschlag_vergleich = "OK"
        # else:
        #     terminzuschlag_vergleich = compare_null_logic(new_df["Terminzuschlag"], ref["Terminzuschlag"])

        # efahrzeug_vergleich
        # below is new function for uebernahme comparison
        # def is_empty(val):
        #     return val is None or (isinstance(val, str) and val.strip() == "")
        # if is_empty(new_df["Car Auktion Protokoll"]):
        #     efahrzeug_vergleich = "OK"
        # else:
        #     efahrzeug_vergleich = compare_null_logic(new_df["Car Auktion Protokoll"], ref["EÜbernahme"])

        # leerfahrt_vergleich
        # def is_emptyleer(val):
        #     return val is None or (isinstance(val, str) and val.strip() == "")
        # if is_empty(new_df["LEERFAHRT"]):
        #     leerfahrt_vergleich = "OK"
        # else:
        #     leerfahrt_vergleich = compare_null_logic(new_df["LEERFAHRT"], ref["Leerfahrt"])

    else:
        auftraggeber_vergleich = vin_vergleich = loadingcity_vergleich = delivercity_vergleich = "NOK" 
        loadingzipcode_vergleich = deliverzipcode_vergleich = "NOK"
        faktor_vergleich = transportrpeis_vergleich = "NOK"
        # telavis_vergleich = seilwinde_vergleich = seilwinde_transport_vergleich = terminzuschlag_vergleich = "NOK"
        # efahrzeug_vergleich = leerfahrt_vergleich = "NOK"
        


# Zusammenstellung der neuen Zeile mit den Originaldaten aus file2 und den Vergleichsergebnissen
    new_row = {
        "InvoiceNr": row["InvoiceNr"],
        "Invoiceshort" : row["Invoiceshort"],
        "Auftraggeber": row["Auftraggeber"],  # Auftraggeber aus file2
        "VIN": row["VIN"],
        "Model": row["Model"],
        "Faktor": row["Faktor"],
        "Total": row["Total"],
        "Loadingcity": row["Loadingcity"],
        "Delivercity": row["Delivercity"],
        "Absender PLZ": row["Absender PLZ"],
        "Empfaenger PLZ": row["Empfaenger PLZ"], 
        # Vergleichsergebnisse
        "AuftraggeberVergleich": auftraggeber_vergleich,
        "VINVergleich": vin_vergleich,
        "Absender PLZ Vergleich": loadingzipcode_vergleich,
        "Empfaenger PLZ Vergleich": deliverzipcode_vergleich,
        "LoadingcityVergleich": loadingcity_vergleich,
        "DelivercityVergleich": delivercity_vergleich,
        "FaktorVergleich": faktor_vergleich,
        "TransportrpeisVergleich": transportrpeis_vergleich
        # "TelavisVergleich": telavis_vergleich,
        # "SeilwindeVergleich": seilwinde_vergleich,
        # "SeilwindeTransportVergleich": seilwinde_transport_vergleich,
        # "TerminzuschlagVergleich": terminzuschlag_vergleich,
        # "E-FahrzeugVergleich": efahrzeug_vergleich,
        # "LeerfahrtVergleich": leerfahrt_vergleich,
        # "Bemerkungen": bemerkungen_text
    }
    compare_results.append(new_row)

# Erstellen eines neuen DataFrames für file4 mit den Vergleichsergebnissen
df_vergleich = pd.DataFrame(compare_results)

# making sure the type of the columns is okey
df_vergleich = df_vergleich.astype({
    "InvoiceNr": "string",
    "Invoiceshort": "string",
    "Auftraggeber": "string",
    "VIN": "string",
    "Model": "string",
    "Faktor": "float",
    "Total": "float",
    "Absender PLZ": "string",
    "Empfaenger PLZ": "string", 
    "Loadingcity": "string",
    "Delivercity": "string",
    "AuftraggeberVergleich": "string",
    "VINVergleich": "string",
    "Absender PLZ Vergleich": "string",
    "Empfaenger PLZ Vergleich": "string",
    "LoadingcityVergleich": "string",
    "DelivercityVergleich": "string",
    "FaktorVergleich": "string",
    "TransportrpeisVergleich": "string"
})


# Conditions for step_1: Anzahl CA3, RRM, Fehler:
ca3_count = df_vergleich[
    (df_vergleich["Auftraggeber"] == "CA3") &
    (df_vergleich["Faktor"] != 0) &
    (df_vergleich["AuftraggeberVergleich"] == "OK")
].shape[0]

rrm_count = df_vergleich[
    (df_vergleich["Auftraggeber"] == "RRM") &
    (df_vergleich["Faktor"] != 0) &
    (df_vergleich["AuftraggeberVergleich"] == "OK")
].shape[0]

fehler_count = df_vergleich[
    (df_vergleich["Auftraggeber"] == "Fehler") &
    (df_vergleich["Faktor"] != 0) &
    (df_vergleich["AuftraggeberVergleich"] == "NOK")
].shape[0]

# creating new dataFrame for step_1
step1_df = pd.DataFrame({
    "CA3": [ca3_count],
    "RRM": [rrm_count],
    "Fehler": [fehler_count]
})

# cleen the df_vergleich to make less columns
columns_to_keep_step_1 = ["Invoiceshort", "VIN", "Model", "Faktor", "Total", "Loadingcity", "Delivercity", "Auftraggeber", "AuftraggeberVergleich"]
df_vergleich_cleen = df_vergleich[columns_to_keep_step_1]

# Conditions for step_1.1-1.3: CA3 RRM Fehler FZG
df_ca3 = df_vergleich_cleen[
    (df_vergleich_cleen["Auftraggeber"] == "CA3") &
    (df_vergleich_cleen["Faktor"] != 0) &
    (df_vergleich_cleen["AuftraggeberVergleich"] == "OK")
]

df_rrm = df_vergleich_cleen[
    (df_vergleich_cleen["Auftraggeber"] == "RRM") &
    (df_vergleich_cleen["Faktor"] != 0) &
    (df_vergleich_cleen["AuftraggeberVergleich"] == "OK")
]

df_fehler = df_vergleich_cleen[
    (df_vergleich_cleen["Auftraggeber"] == "Fehler") &
    (df_vergleich_cleen["Faktor"] != 0) &
    (df_vergleich_cleen["AuftraggeberVergleich"] == "NOK")
]

# creating new dataFrame for step_1_1
step_1_1_df = df_fehler
# creating new dataFrame for step_1_2
step_1_2_df = df_ca3
# creating new dataFrame for step_1_3
step_1_3_df = df_rrm

# opening excel file
with pd.ExcelWriter("file4.xlsx", engine="openpyxl") as writer:
    
    # writing Fehlerreport sheet
    new_df.to_excel(writer, sheet_name="Fehlerreport", index=False)
    worksheet = writer.sheets["Fehlerreport"]
    for col_num in range(1, len(new_df.columns) + 1):
        worksheet.column_dimensions[get_column_letter(col_num)].width = 25

    # writing CA3 sheet
    df_ca3.to_excel(writer, sheet_name="CA3", index=False)
    worksheet = writer.sheets["CA3"]
    for col_num in range(1, len(df_ca3.columns) + 1):
        worksheet.column_dimensions[get_column_letter(col_num)].width = 15

    # writing RRM sheet
    df_rrm.to_excel(writer, sheet_name="RRM", index=False)
    worksheet = writer.sheets["RRM"]
    for col_num in range(1, len(df_rrm.columns) + 1):
        worksheet.column_dimensions[get_column_letter(col_num)].width = 15
    
    # writing Vergleich sheet
    df_vergleich.to_excel(writer, sheet_name="Vergleich", index=False)
    worksheet = writer.sheets["Vergleich"]
    for col_num in range(1, len(df_vergleich.columns) + 1):
        worksheet.column_dimensions[get_column_letter(col_num)].width = 15
    
    # writing Step1 sheet (number CA3, RRM, Fehler)
    step1_df.to_excel(writer, sheet_name="Step_1 Anzahl der FZG", index=False)
    worksheet = writer.sheets["Step_1 Anzahl der FZG"]
    for col_num in range(1, len(step1_df.columns) + 1):
        worksheet.column_dimensions[get_column_letter(col_num)].width = 25

    # writing Step1_1 sheet (Fehler)
    step_1_1_df.to_excel(writer, sheet_name="Step_1_1 Fehler", index=False)
    worksheet = writer.sheets["Step_1_1 Fehler"]
    for col_num in range(1, len(step_1_1_df.columns) + 1):
        worksheet.column_dimensions[get_column_letter(col_num)].width = 25
    
    # writing Step1_2 sheet (CA3)
    step_1_2_df.to_excel(writer, sheet_name="Step_1_2 CA3", index=False)
    worksheet = writer.sheets["Step_1_2 CA3"]
    for col_num in range(1, len(step_1_2_df.columns) + 1):
        worksheet.column_dimensions[get_column_letter(col_num)].width = 25
    
    # writing Step1_3 sheet (CA3)
    step_1_3_df.to_excel(writer, sheet_name="Step_1_3 RRM", index=False)
    worksheet = writer.sheets["Step_1_3 RRM"]
    for col_num in range(1, len(step_1_3_df.columns) + 1):
        worksheet.column_dimensions[get_column_letter(col_num)].width = 25
