
from flask import Flask, request, jsonify
import pandas as pd
import os
import tempfile
import requests

app = Flask(__name__)

AIRTABLE_BASE_ID = os.environ.get("AIRTABLE_BASE_ID")
AIRTABLE_API_KEY = os.environ.get("AIRTABLE_API_KEY")
AIRTABLE_ORDERS_TABLE = "Užsakymai"
AIRTABLE_LINES_TABLE = "Užsakymo eilutės"

@app.route("/")
def index():
    return "Serveris veikia", 200

@app.route("/process", methods=["POST"])
def process_excel():
    if 'file' not in request.files:
        return jsonify({"error": "Nerastas failas"}), 400

    file = request.files['file']
    filename = file.filename

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsm") as tmp:
        file.save(tmp.name)
        excel_path = tmp.name

    try:
        df = pd.read_excel(excel_path, sheet_name="Užsakymas", skiprows=8)
        df = df.dropna(how="all")  # Pašalinti tuščias eilutes

        # Užsakymo duomenys iš pirmųjų langelių (virš lentelės)
        metadata = pd.read_excel(excel_path, sheet_name="Užsakymas", nrows=8, header=None)
        customer = metadata.iloc[0, 1]
        order_number = metadata.iloc[1, 1]
        order_date = str(metadata.iloc[2, 1])[:10]

        # Įkelti į Airtable Užsakymai
        order_payload = {
            "fields": {
                "Užsakymo numeris": order_number,
                "Klientas": customer,
                "Data": order_date
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
            return jsonify({"error": "Nepavyko įkelti užsakymo į Airtable", "details": order_response.text}), 500

        order_id = order_response.json()["id"]

        # Įkelti kiekvieną eilutę į „Užsakymo eilutės“
        for _, row in df.iterrows():
            line_payload = {
                "fields": {
                    "Užsakymo ID": [order_id],
                    "Pozicija": str(row.get("Pozicija", "")),
                    "Gaminio tipas": str(row.get("Gaminio tipas", "")),
                    "Matmenys": str(row.get("Matmenys", "")),
                    "Spalva": str(row.get("Spalva", "")),
                    "Kiekis": int(row.get("Kiekis", 1))
                }
            }

            line_response = requests.post(
                f"https://api.airtable.com/v0/{AIRTABLE_BASE_ID}/{AIRTABLE_LINES_TABLE}",
                headers={
                    "Authorization": f"Bearer {AIRTABLE_API_KEY}",
                    "Content-Type": "application/json"
                },
                json=line_payload
            )

            if line_response.status_code != 200:
                return jsonify({"error": "Nepavyko įkelti užsakymo eilutės", "details": line_response.text}), 500

        return jsonify({"status": "ok", "message": f"Gautas failas: {filename}"}), 200

    finally:
        os.unlink(excel_path)

if __name__ == "__main__":
    app.run(debug=True)
