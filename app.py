import os
import requests 
import pandas as pd 
import subprocess
import sys
import uuid
import json
# import httpx
from flask import Flask, render_template, request, redirect, flash, url_for, send_file

app = Flask(__name__)
app.secret_key = 'dein_geheimer_schluessel' 

# Define the Base directiry of our project (where input and output files are stored)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DESIRED_ORDER = [
    "Auftraggeber", "invoice", "vin", "vihicle", "loadingcity", "delivercity",
    "Faktor", "Gallikerpreis", "Telavic", "Seilwinde", "Terminzuschlag",
    "EÜbernahme", "Leerfahrt", "Seilwindeintransport"
]

def reorder_df(df, desired):
    for col in desired:
        if col not in df.columns:
            df[col] = ""
    return df[desired]

@app.route('/health', methods=['GET', 'HEAD'])
def health_check():
    return "OK", 200

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # File from the form
        invoice_file = request.files.get('invoice')
        
        # Checking that the file is uploaded
        if not invoice_file:
            flash("Bitte laden die erforderliche Invoice-Datei hoch")
            return redirect(request.url)
        
        # Saving to temporary with unique name
        temp_name = f"invoice_{uuid.uuid4().hex}.pdf"
        temp_path = os.path.join(BASE_DIR, temp_name)
        invoice_file.save(temp_path)

        # Rerighting into invoice.pdf, which will be analysed by pdf4.py, saving to BASE_DIR
        invoice_path = os.path.join(BASE_DIR, 'invoice.pdf')
        if os.path.exists(invoice_path):
            os.remove(invoice_path)
        os.replace(temp_path, invoice_path)

        # Checking: output of size and time of modification
        statinfo = os.stat(invoice_path)
        print("Invoice gespeichert:", invoice_path,
              "Größe:", statinfo.st_size, "Bytes",
              "mtime:", statinfo.st_mtime)
        print("Invoice gespeichert:", invoice_path, os.path.getsize(invoice_path))
        
        # Saving to в BASE_DIR
        # invoice_path = os.path.join(BASE_DIR, 'invoice.pdf')
        # if os.path.exists(invoice_path):
        #     os.remove(invoice_path)
        # invoice_file.save(invoice_path)

        # For deploying on render
        CA3 = os.getenv("CA3_URL")
        RRM = os.getenv("RRM_URL")

        # For local deployment, we just read config.json
        if not CA3 or not RRM:
            config_path = os.path.join(BASE_DIR, "config.json")
            if os.path.exists(config_path):
                with open(config_path) as f:
                    config = json.load(f)
                CA3 = config.get("CA3_URL", "")
                RRM = config.get("RRM_URL", "")
        
        # CA3
        resp_ca3 = requests.get(CA3, verify=False)
        # resp_ca3 = httpx.get(CA3, timeout=120, follow_redirects=True)
        if resp_ca3.status_code == 200:
            data_ca3 = resp_ca3.json()
            df_public_ca3 = pd.json_normalize(data_ca3) 
            df_public_ca3 = reorder_df(df_public_ca3, DESIRED_ORDER)
            ca3_path = os.path.join(BASE_DIR, 'ca3.xlsx')
            df_public_ca3.to_excel(ca3_path, index=False, engine="openpyxl")
            print("Excel CA3 erfolgreich gespeichert!")
        else:
            print("Error while saving CA3 file:", resp_ca3.status_code, resp_ca3.text[:200])

        # RRM
        resp_rrm = requests.get(RRM, verify=False)
        # resp_rrm = httpx.get(RRM, timeout=120, follow_redirects=True)
        if resp_rrm.status_code == 200:
            data_rrm = resp_rrm.json()
            df_public_rrm = pd.json_normalize(data_rrm) 
            df_public_rrm = reorder_df(df_public_rrm, DESIRED_ORDER)
            rrm_path = os.path.join(BASE_DIR, 'rrm.xlsx')
            df_public_rrm.to_excel(rrm_path, index=False, engine="openpyxl")
            print("Excel RRM erfolgreich gespeichert!")
        else:
            print("Error while saving RRM file:", resp_rrm.status_code, resp_rrm.text[:200])

        # We call the analysis function, which will do the pdf4.py script and process the files
        run_analysis()
        
        # We delete the files, as we don not need them anymore 
        for f in ['invoice.pdf', 'ca3.xlsx', 'rrm.xlsx']:
            path = os.path.join(BASE_DIR, f)
            if os.path.exists(path):
                os.remove(path)
        
        # Rederecting the user to the download page 
        return redirect(url_for("download_page"))
    
    return render_template("index.html")


def run_analysis():
    """
    Функция run_analysis() выполняет скрипт pdf4.py, который ожидается создать файлы:
        file1.xlsx, file2.xlsx, file3.xlsx, file4.xlsx.
    После выполнения эти файлы переименовываются в:
        Gesamtinvoiceinfo.xlsx, Invoiceinfo.xlsx, Rechnungsprüfung.xlsx, Validierung.xlsx,
    и остаются в BASE_DIR для последующего скачивания.
    """
    # Names of the files thet were created by pdf4.py
    original_files = ['file1.xlsx', 'file2.xlsx', 'file3.xlsx', 'file4.xlsx']
    # New names of the files
    new_names = [
        'Gesamtinvoiceinfo.xlsx', 
        'Invoiceinfo.xlsx', 
        'Rechnungsprüfung.xlsx', 
        'Fehlerreport.xlsx'
    ]
    
    try:
        result = subprocess.run(
            [sys.executable, 'pdf4.py'],
            cwd=BASE_DIR,
            capture_output=True,
            text=True
        )
        print("stdout:\n", result.stdout)
        print("stderr:\n", result.stderr)
        print("Return code:", result.returncode)
    except Exception as ex:
        print("Fehler beim Ausführen von pdf4.py:", ex)
    
    # Renaming and moving the files
    for original, new_name in zip(original_files, new_names):
        src = os.path.join(BASE_DIR, original)
        dst = os.path.join(BASE_DIR, new_name)
        if os.path.exists(src):
            try:
                os.replace(src, dst)
                print(f"Datei {original} wurde in {new_name} umbenannt.")
            except Exception as e:
                print(f"Fehler beim Umbenennen der Datei {original}: {e}")
        else:
            print(f"Ausgabedatei {original} wurde nicht gefunden.")


@app.route("/downloads")
def download_page():
    # List of the files for downloading
    results = [
        'Gesamtinvoiceinfo.xlsx', 
        'Invoiceinfo.xlsx', 
        'Rechnungsprüfung.xlsx', 
        'Fehlerreport.xlsx'
    ]
    return render_template("download.html", results=results)


@app.route("/download/<filename>")
def download_file(filename):
    file_path = os.path.join(BASE_DIR, filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        return "Datei nicht gefunden", 404


if __name__ == '__main__':
    app.run(debug=True)
