
from flask import Flask, request, jsonify
import pandas as pd
import openpyxl
import io
import requests
import os

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
        df = pd.read_excel(in_memory_file, sheet_name="Užsakymas")
        df = df.fillna("")

        # Užsakymo duomenys (pirma eilutė)
        order_data = df.iloc[0].to_dict()

        order_payload = {
            "fields": {
                "Klientas": order_data.get("Klientas"),
                "Miestas": order_data.get("Miestas"),
                "Telefonas": order_data.get("Telefonas"),
                "El. paštas": order_data.get("El. paštas"),
                "Komentaras": order_data.get("Komentaras")
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
            return jsonify({"error": "Nepavyko įrašyti užsakymo", "details": order_response.text}), 500

        order_id = order_response.json()["id"]

        # Užsakymo eilutės
        lines_payload = {
            "records": []
        }

        for i in range(1, len(df)):
            line = df.iloc[i].to_dict()
            record = {
                "fields": {
                    "Prekė": line.get("Prekė"),
                    "Kiekis": line.get("Kiekis"),
                    "Matmenys": line.get("Matmenys"),
                    "Pastabos": line.get("Pastabos"),
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
            return jsonify({"error": "Nepavyko įrašyti užsakymo eilučių", "details": lines_response.text}), 500

        return jsonify({"status": "ok"}), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500
