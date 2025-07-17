from flask import Flask, request, jsonify
import pandas as pd
from io import BytesIO
import os
import requests

app = Flask(__name__)

AIRTABLE_BASE_ID = os.getenv("AIRTABLE_BASE_ID")
AIRTABLE_API_KEY = os.getenv("AIRTABLE_API_KEY")
TABLE_ORDERS = "Užsakymai"
TABLE_LINES = "Užsakymo eilutės"

def send_to_airtable(table_name, records):
    url = f"https://api.airtable.com/v0/{AIRTABLE_BASE_ID}/{table_name}"
    headers = {
        "Authorization": f"Bearer {AIRTABLE_API_KEY}",
        "Content-Type": "application/json"
    }
    for record in records:
        data = {"fields": record}
        response = requests.post(url, json=data, headers=headers)
        response.raise_for_status()

@app.route("/", methods=["GET"])
def index():
    return "Serveris veikia. Naudokite POST į /process."

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})

@app.route("/process", methods=["POST"])
def process_file():
    if "file" not in request.files:
        return jsonify({"error": "No file part"}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "No selected file"}), 400

    try:
        in_memory_file = BytesIO()
        file.save(in_memory_file)
        in_memory_file.seek(0)

        xls = pd.ExcelFile(in_memory_file)
        df = pd.read_excel(xls, sheet_name="Užsakymas")

        uzsakymas = {
            "Klientas": df.loc[0, "Klientas"],
            "Data": str(df.loc[0, "Data"]),
            "Miestas": df.loc[0, "Miestas"],
            "Pastabos": df.loc[0, "Pastabos"]
        }

        send_to_airtable(TABLE_ORDERS, [uzsakymas])

        rows = []
        for i in range(1, len(df)):
            row = {
                "Prekė": df.loc[i, "Prekė"],
                "Kiekis": df.loc[i, "Kiekis"],
                "Kaina": df.loc[i, "Kaina"]
            }
            rows.append(row)

        send_to_airtable(TABLE_LINES, rows)

        return jsonify({"status": "success", "rows_uploaded": len(rows)})

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True)