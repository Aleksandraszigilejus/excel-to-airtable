
from flask import Flask, request, jsonify
import pandas as pd
import openpyxl
import io
import requests
import os
import re

app = Flask(__name__)

AIRTABLE_BASE_ID = os.getenv("AIRTABLE_BASE_ID")
AIRTABLE_ORDERS_TABLE = os.getenv("AIRTABLE_ORDERS_TABLE")
AIRTABLE_ORDER_LINES_TABLE = os.getenv("AIRTABLE_ORDER_LINES_TABLE")
AIRTABLE_API_KEY = os.getenv("AIRTABLE_API_KEY")

@app.route("/", methods=["GET"])
def index():
    return "Serveris veikia", 200

@app.route("/process", methods=["POST"])
def process_excel():
    if "file" not in request.files:
        return jsonify({"error": "No file part"}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "No selected file"}), 400

    try:
        in_memory_file = io.BytesIO()
        file.save(in_memory_file)
        in_memory_file.seek(0)

        # Nuskaitymas be antraščių
        raw_df = pd.read_excel(in_memory_file, sheet_name="Užsakymas", header=None)
        in_memory_file.seek(0)
        df = pd.read_excel(in_memory_file, sheet_name="Užsakymas", header=8)
        df = df.fillna("")

        raw_text = raw_df.astype(str).values.flatten()
        content = "
".join(raw_text)

        def search(pattern):
            match = re.search(pattern, content, re.IGNORECASE)
            return match.group(1).strip() if match else ""

        klientas = search(r"Mok[eė]tojas\s*:\s*(.+)")
        miestas = search(r"Adresas\s*:\s*(.+)")
        komentaras = search(r"Data\s*:\s*(.+)")
        telefonas = search(r"(?:(?:\+370|8)\s?\d{3}[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2})")
        el_pastas = search(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")

        # Užsakymo duomenys
        order_payload = {
            "fields": {
                "Klientas": klientas,
                "Miestas": miestas,
                "Telefonas": telefonas,
                "El. paštas": el_pastas,
                "Komentaras": komentaras,
            }
        }

        order_response = requests.post(
            f"https://api.airtable.com/v0/{AIRTABLE_BASE_ID}/{AIRTABLE_ORDERS_TABLE}",
            headers={
                "Authorization": f"Bearer {AIRTABLE_API_KEY}",
                "Content-Type": "application/json"
            },
            json=order_payload
        )

        if order_response.status_code != 200:
            return jsonify({
                "error": "Nepavyko įrašyti užsakymo",
                "details": order_response.text,
                "payload": order_payload
            }), 500

        order_id = order_response.json()["id"]

        # Užsakymo eilutės
        df.columns = df.columns.str.replace("\\n", " ").str.strip()

        def find_col(possibilities):
            for col in df.columns:
                if any(p.lower() in col.lower() for p in possibilities):
                    return col
            return None

        gaminys_col = find_col(["Gaminio", "Prek"])
        kiekis_col = find_col(["Kiek"])
        plotis_col = find_col(["Plotis"])
        aukstis_col = find_col(["Aukšt"])
        pastabos_col = find_col(["Pastab"])

        lines_payload = {"records": []}
        for _, row in df.iterrows():
            gaminys = row.get(gaminys_col, "")
            kiekis = row.get(kiekis_col, "")
            plotis = row.get(plotis_col, "")
            aukstis = row.get(aukstis_col, "")
            pastabos = row.get(pastabos_col, "")
            matmenys = f"{plotis} x {aukstis}".strip(" x")

            record = {
                "fields": {
                    "Prekė": gaminys,
                    "Kiekis": kiekis,
                    "Matmenys": matmenys,
                    "Pastabos": pastabos,
                    "Užsakymas": [order_id]
                }
            }
            lines_payload["records"].append(record)

        lines_response = requests.post(
            f"https://api.airtable.com/v0/{AIRTABLE_BASE_ID}/{AIRTABLE_ORDER_LINES_TABLE}",
            headers={
                "Authorization": f"Bearer {AIRTABLE_API_KEY}",
                "Content-Type": "application/json"
            },
            json=lines_payload
        )

        if lines_response.status_code != 200:
            return jsonify({
                "error": "Nepavyko įrašyti užsakymo eilučių",
                "details": lines_response.text,
                "payload": lines_payload
            }), 500

        return jsonify({
            "status": "ok",
            "order_id": order_id,
            "eilutes": len(lines_payload["records"])
        }), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500
