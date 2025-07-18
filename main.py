
from flask import Flask, request, jsonify
import pandas as pd
import os
from openpyxl import load_workbook

app = Flask(__name__)

@app.route("/health", methods=["GET"])
def health_check():
    return jsonify(status="ok")

@app.route("/process", methods=["POST"])
def process_file():
    try:
        if 'file' not in request.files:
            return jsonify(error="Lauke 'file' nėra failo"), 400

        file = request.files['file']

        if file.filename == "":
            return jsonify(error="Failas neturi pavadinimo"), 400

        print(f"Gautas failas: {file.filename}")

        # Patikriname failo formatą
        if not file.filename.endswith((".xls", ".xlsx", ".xlsm")):
            return jsonify(error="Netinkamas failo formatas, turi būti Excel (.xls/.xlsx/.xlsm)"), 400

        # Įkeliame Excel failą
        wb = load_workbook(file, data_only=True)
        if "Užsakymas" not in wb.sheetnames:
            return jsonify(error="Nėra lapo 'Užsakymas'"), 400

        sheet = wb["Užsakymas"]

        rows = list(sheet.iter_rows(values_only=True))
        if not rows or len(rows) < 5:
            return jsonify(error="Lapas 'Užsakymas' yra tuščias arba per trumpas"), 400

        order_data = {
            "uzsakovas": rows[1][1],
            "miestas": rows[2][1],
            "el_pastas": rows[3][1],
            "tel": rows[4][1],
            "komentaras": rows[5][1] if len(rows) > 5 else ""
        }

        # Užsakymo eilutės (pradedant nuo 10-os eilutės)
        items = []
        for row in rows[9:]:
            if row[0] is None:
                continue
            item = {
                "poz": row[0],
                "kiekis": row[1],
                "aukstis": row[2],
                "plotis": row[3],
                "medis": row[4],
                "spalva": row[5],
                "mechanizmas": row[6]
            }
            items.append(item)

        return jsonify({
            "order": order_data,
            "items": items
        })

    except Exception as e:
        return jsonify(error=f"Klaida apdorojant failą: {str(e)}"), 500


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))
