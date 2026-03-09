# 🧾 Rechnungsprüfung App

**Live:** [rechnungsprufungapp-wfp-with.onrender.com](https://rechnungsprufungapp-wfp-with.onrender.com)

---

## Overview

A web application for automated validation of transportation invoices. It verifies that billed items belong to us, confirms that costs have been correctly re-invoiced to our clients, and calculates key financial metrics including quantified losses.

Data is fetched directly from our internal database via **Metabase**.


We receive various invoices from the transportation company. 
The purpose of this application is to validate these invoices — that is, to verify that the billed items are indeed ours, and that we have re-invoiced these items to our client. 
It also calculates certain key metrics and quantifies our losses.
The App takes the data directly from our DB (Metabase)

Techstack: python(flask pandas requests httpx sys re pdfplumber  pdopenpyxl uuid openpyxl.utils get_column_letter os json math httpx subprocess), Metabase, html 

---

## Features

- Upload invoices in PDF or Excel format
- Automatic validation against live Metabase database records
- Checks that billed transports match our orders (VIN, route, client)
- Verifies surcharges and additional costs are correctly re-invoiced
- Calculates totals, detects duplicates, and flags anomalies
- Exports a structured error report (`Fehlerreport.xlsx`) with categorized findings

---

## Supported Invoice Types

| Script | Invoice Type |
|--------|-------------|
| `pdf2.py` | Cotra — Hauptrechnung (Excel) |
| `pdf3.py` | Galliker — Hauptrechnung (Excel) |
| `pdf4.py` | Galliker — Hauptrechnung (PDF) |
| `pdf5.py` | Galliker — Schild und Ausweise (Excel) |
| `pdf6.py` | Galliker — Service Leistungen (Excel) |
| `pdf7.py` | Galliker — Batterie (Excel) |
| `pdf8.py` | Galliker — HV Batterie (Excel) |
| `pdf9.py` | Galliker — eÜbernahme (Excel) |
| `pdf10.py` | Galliker — LagerßHandling (Excel) |

---

## Output Structure (Fehlerreport.xlsx)
Each report contains the following sheets:

| Sheet | Description |
|-------|-------------|
| `Step_1 Anzahl TA` | Summary count: CA3 / RRM / Errors |
| `Step_1_1 Fehler` | Transports not found in database |
| `Step_1_2 CA3` | All CA3 transports |
| `Step_1_3 RRM` | All RRM transports |
| `Step_1_4 Zusätzlich` | Mismatches in route, city or WFP flag |
| `Step_2_Falschbeträge` | Incorrect invoice amounts |
| `Step_3_Nebenkosten` | Surcharge summary by category |
| `Step_4_Weiteverrechnet` | Surcharges — re-invoiced check |
| `Step_5_PW_SUV_LNF` | Vehicle type breakdown |
| `Step_6_Dublikate` | Duplicate entries |
| `Step_7_Gesamtbetrag` | Total amount reconciliation |
| `Step_8_Komische_Transporte` | Suspicious transports |


---

## Tech Stack

**Backend**
- Python, Flask
- pandas, openpyxl, pdfplumber
- requests, httpx, subprocess

**Frontend**
- HTML (Jinja2 templates)

**Data Source**
- Metabase (live DB queries via JSON API)

---
## Deployment
Deployed on **Render** using environment variables for all database URLs (no `config.json` needed in production)