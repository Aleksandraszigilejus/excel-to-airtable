
from flask import Flask, request, jsonify
import pandas as pd
import tempfile
import os

app = Flask(__name__)

@app.route('/process', methods=['POST'])
def process_excel():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part in request'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    try:
        # Išsaugom įkeltą failą laikinai
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            file.save(tmp.name)
            filepath = tmp.name

        # Nuskaitom tik lapą „Užsakymas“
        df = pd.read_excel(filepath, sheet_name="Užsakymas")

        # Pašalinam tuščias eilutes
        df.dropna(how='all', inplace=True)

        # Raskime, kur yra pirmoji reikšmė "Kliento vardas"
        order_info_row = df[df.iloc[:, 0] == "Kliento vardas"].index[0]

        # Užsakymo informacija
        order_info = df.iloc[order_info_row:order_info_row+5].set_index(df.columns[0]).T.to_dict('records')[0]

        # Užsakymo eilutės – prasideda po tuščios eilutės
        start_row = df[df.iloc[:, 0] == "Poz. Nr."].index[0]
        order_lines = df.iloc[start_row+1:].dropna(how='all')

        lines = []
        for _, row in order_lines.iterrows():
            lines.append({
                "pozicija": row[0],
                "tipas": row[1],
                "plotis": row[2],
                "aukstis": row[3],
                "kiekis": row[4]
            })

        # Išvalom failą
        os.unlink(filepath)

        return jsonify({
            "order": order_info,
            "lines": lines
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run()
