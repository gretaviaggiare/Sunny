import pdfplumber
import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox
import sys
import os

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

TEMPLATE_PATH = resource_path("import prenotazioni_MASTER - ALL INCLUSIVE PRODUCT SUNNY CARS.xlsx")
AIRPORT_MAP = {
    "AHO": "ALGHERO AEROPORTO",
    "BDS": "BRINDISI AEROPORTO",
    "BGY": "BERGAMO AEROPORTO",
    "BOR": "BORGO VALSUGANA",
    "BRI": "BARI AEROPORTO",
    "CA": "CAGLIARI AEROPORTO",
    "CTA": "CATANIA AEROPORTO",
    "FCO": "ROMA FIUMICINO AEROPORTO",
    "OT": "OLBIA AEROPORTO",
    "PMO": "PALERMO AEROPORTO",
    "PSR": "PESCARA AEROPORTO",
    "RDG": "RODI GARGANICO",
    "SUF": "LAMEZIA TERME AEROPORTO",
    "VE": "VENEZIA AEROPORTO",
    "VI": "ROSSANO VENETO",
    "VRN": "VERONA AEROPORTO"
}

# keyword → stesso output
KEYWORD_MAP = {
    "ALGHERO": "ALGHERO AEROPORTO",
    "BRINDISI": "BRINDISI AEROPORTO",
    "BERGAMO": "BERGAMO AEROPORTO",
    "BARI": "BARI AEROPORTO",
    "CAGLIARI": "CAGLIARI AEROPORTO",
    "CATANIA": "CATANIA AEROPORTO",
    "ROMA": "ROMA FIUMICINO AEROPORTO",
    "FIUMICINO": "ROMA FIUMICINO AEROPORTO",
    "OLBIA": "OLBIA AEROPORTO",
    "PALERMO": "PALERMO AEROPORTO",
    "PESCARA": "PESCARA AEROPORTO",
    "LAMEZIA": "LAMEZIA TERME AEROPORTO",
    "VENEZIA": "VENEZIA AEROPORTO",
    "VERONA": "VERONA AEROPORTO",
    "ROSSANO": "ROSSANO VENETO"
}

# -------------------------
# ESTRAZIONE DATI DAL PDF
# -------------------------
def extract_reservation(text):

    def find(pattern):
        match = re.search(pattern, text)
        return match.group(1).strip() if match else ""

    reservation = find(r"RESERVATION REPORT #(\d+)")
    targetrent = find(r"Our Ref\.\:\s*(\d+)")

    # Keep only the part BEFORE "To:"
    # PICKUP
    pickup_match = re.search(r"Pickup:\s*(.*?)(?:\n|To:)", text)
    pickup_raw = pickup_match.group(1).strip() if pickup_match else ""
    pickup = map_airport(pickup_raw)

    # DROPOFF (one way support)
    dropoff_match = re.search(r"Drop off:\s*(.*?)(?:\n|Car Group:)", text)
    dropoff_raw = dropoff_match.group(1).strip() if dropoff_match else ""

    if dropoff_raw:
        dropoff = map_airport(dropoff_raw)
    else:
        dropoff = pickup  # fallback se non esiste

    # Date
    from_match = re.search(r"From:\s*(\d{2}\.\d{2}\.\d{4}) / (\d{2}:\d{2})", text)
    to_match = re.search(r"To:\s*(\d{2}\.\d{2}\.\d{4}) / (\d{2}:\d{2})", text)

    data_out = from_match.group(1) if from_match else ""
    hour_out = from_match.group(2) if from_match else ""

    data_in = to_match.group(1) if to_match else ""
    hour_in = to_match.group(2) if to_match else ""

    data_out = data_out.replace(".", "/")
    data_in = data_in.replace(".", "/")

    days = find(r"Days:\s*(\d+)")

    # GRP
    car_full = find(r"Car Group:\s*(.+)")
    grp = car_full.split(" ")[0] if car_full else ""

    # Driver
    driver_match = re.search(r"Driver:\s*(.*?)\s*Our Ref\.:", text)
    driver_raw = driver_match.group(1).strip() if driver_match else find(r"Driver:\s*(.+)")

    driver = re.sub(r"Mr\.|Mrs\.|Ms\.", "", driver_raw).replace(",", "").strip()
    # Observations
    flight = find(r"Flight No\.\:\s*(\S+)")
    comment = find(r"Comment:\s*(.+)")
    extras = find(r"Extras:\s*(.+)")
    parts = []

    if flight:
        parts.append(f"Flight: {flight}")

    if comment:
        parts.append(f" - {comment}")

    if extras:
        parts.append(f" - {extras}")

    observations = " ".join(parts)
    
    # Rate logic
    rate_match = re.search(r"Rate Code:\s*(.+)", text)
    rate_source = rate_match.group(1).strip() if rate_match else ""

    rate_source = rate_source.upper()

    if "EXP" in rate_source:
        fonte_fix = "_TO-SunnyCars ALL INCLUSIVE EXPRESS Prepaid"
    elif "ND" in rate_source:
        fonte_fix = "_TO-SunnyCars ALL INCLUSIVE NO DEPOSIT Prepaid"
    else:
        fonte_fix = "_TO-SunnyCars ALL INCLUSIVE Prepaid"

    return {
        "Reservation Status": "NEW",
        "Broker reservation number": targetrent,
        "Targetrent number": targetrent,
        "Station pick-up": pickup,
        "Data out": data_out,
        "Hour out": hour_out,
        "Station drop-off": dropoff,
        "Data in": data_in,
        "Hour in": hour_in,
        "Days": days,
        "GRP": grp,
        "Rate code": "VIAG",
        "Customer name": driver,
        "Observations": observations,
        "FONTE FIX": fonte_fix,
        "MEZZO": ""
    }

# -------------------------
# MAPPA AEREOPORTI
# -------------------------
def map_airport(pickup):

    if not pickup:
        return ""

    text = pickup.upper()

    # 🔹 1. prova con sigla (prima del -)
    code = pickup.split("-")[0].strip().upper()

    if code in AIRPORT_MAP:
        return AIRPORT_MAP[code]

    # 🔹 2. prova con keyword dentro la stringa
    for key, value in KEYWORD_MAP.items():
        if key in text:
            return value

    # 🔹 fallback
    return pickup

# -------------------------
# PROCESSO FILE
# -------------------------
def process_files(files):

    extracted_rows = []

    for file in files:
        with pdfplumber.open(file) as pdf:

            for page in pdf.pages:
                text = page.extract_text()

                if not text:
                    continue

                blocks = re.split(r"\bRESERVATION\b", text)

                for block in blocks[1:]:  # skip header

                    if "Our Ref." not in block:
                        continue

                    row = extract_reservation(block)

                    if row["Targetrent number"]:
                        extracted_rows.append(row)

    new_df = pd.DataFrame(extracted_rows)

    # -------------------------
    # CARICA TEMPLATE
    # -------------------------
    try:
        template_df = pd.read_excel(TEMPLATE_PATH)
    except Exception as e:
        messagebox.showerror("Errore", f"Template non trovato:\n{e}")
        return

    # -------------------------
    # ALLINEA COLONNE
    # -------------------------
    for col in template_df.columns:
        if col not in new_df.columns:
            new_df[col] = ""

    new_df = new_df[template_df.columns]

    # -------------------------
    # UNISCE DATI
    # -------------------------
    final_df = pd.concat([template_df, new_df], ignore_index=True)

    # -------------------------
    # RIMUOVE DUPLICATI
    # -------------------------
    final_df.columns = final_df.columns.str.strip()

    final_df = final_df.drop_duplicates(
        subset=["Broker reservation number", "Targetrent number"]
    )

    # -------------------------
    # SALVA FILE
    # -------------------------
    output_path = os.path.join(
        os.path.expanduser("~"),
        "Desktop",
        "prenotazioni_sunny.xlsx"
    )

    try:
        final_df.to_excel(output_path, index=False)
    except PermissionError:
        messagebox.showerror(
            "Errore",
            "Chiudi il file 'prenotazioni_sunny.xlsx' prima di salvare."
        )
        return

    messagebox.showinfo(
        "Operazione completata",
        "File salvato sul Desktop:\nprenotazioni_sunny.xlsx\n\n(Il file è stato aggiornato se già esistente)"
    )


# -------------------------
# GUI
# -------------------------
def open_files():
    files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if files:
        process_files(files)


root = tk.Tk()
root.title("PDF → Excel SunnyCars")

label = tk.Label(root, text="Seleziona i PDF da convertire", pady=20)
label.pack()

btn = tk.Button(root, text="Seleziona PDF", command=open_files)
btn.pack(pady=10)

root.mainloop()
