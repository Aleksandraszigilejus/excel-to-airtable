
from flask import Flask, request, jsonify
import pandas as pd
import os
import tempfile
import requests

app = Flask(__name__)

# Airtable nustatymai (pakeisk savo reikšmėmis)
airtable_token = os.environ.get("AIRTABLE_TOKEN")
base_id = os.environ.get("AIRTABLE_BASE_ID")
orders_table_name = "Užsakymai"
lines_table_name = "Užsakymo eilutės"

headers = {
    "Authorization": f"Bearer {airtable_token}",
    "Content-Type": "application/json"
}

@app.route("/health", methods=["GET"])
def health_check():
    return jsonify({"status": "ok"})

@app.route("/process", methods=["POST"])
def process_excel():
    file = request.files.get("file")
    if not file:
        return jsonify({"error": "No file provided"}), 400

    # Išsaugom failą
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsm") as tmp:
        file.save(tmp.name)
        file_path = tmp.name

    print("Gautas failas:", file.filename)

    try:
        # Nuskaitom Excel lapą
        df = pd.read_excel(file_path, sheet_name="Užsakymas", header=None)

        # Užsakymo duomenys (viršutinė dalis)
        order_fields = {
            "Klientas": df.iloc[1, 1],
            "Užsakymo numeris": df.iloc[0, 1],
            "Data": df.iloc[2, 1].strftime("%Y-%m-%d") if pd.notnull(df.iloc[2, 1]) else None,
            "Komentaras": df.iloc[3, 1]
        }

        response1 = requests.post(
            f"https://api.airtable.com/v0/{base_id}/{orders_table_name}",
            headers=headers,
            json={"fields": order_fields}
        )
        print("Airtable response 1:", response1.status_code, response1.text)

        if response1.status_code != 200:
            return jsonify({"error": "Nepavyko sukurti užsakymo"}), 500

        order_id = response1.json()["id"]

        # Užsakymo eilutės (nuo 7 eilutės)
        for i in range(6, len(df)):
            if pd.isnull(df.iloc[i, 0]):
                continue  # praleidžiam tuščias eilutes
            line_fields = {
                "Užsakymo ID": [order_id],
                "Pozicija": str(df.iloc[i, 0]),
                "Gaminio tipas": df.iloc[i, 1],
                "Kiekis": df.iloc[i, 2],
                "Plotis": df.iloc[i, 3],
                "Aukštis": df.iloc[i, 4],
                "Spalva": df.iloc[i, 5]
            }

            response2 = requests.post(
                f"https://api.airtable.com/v0/{base_id}/{lines_table_name}",
                headers=headers,
                json={"fields": line_fields}
            )
            print(f"Airtable response 2 [{i}]:", response2.status_code, response2.text)

        return jsonify({"status": "success"})

    except Exception as e:
        print("Klaida:", str(e))
        return jsonify({"error": str(e)}), 500

    finally:
        os.remove(file_path)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
