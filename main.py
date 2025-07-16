import pandas as pd
from flask import Flask, request, jsonify
import os

app = Flask(__name__)

@app.route('/')
def home():
    return 'Serveris veikia!'

@app.route('/process', methods=['POST'])
def process_excel():
    file = request.files['file']
    path = "/tmp/" + file.filename
    file.save(path)

    try:
        df_orders = pd.read_excel(path, sheet_name='Užsakymai')
        df_lines = pd.read_excel(path, sheet_name='Užsakymo eilutės')

        return jsonify({
            "uzsakymai": df_orders.to_dict(orient="records"),
            "eilutes": df_lines.to_dict(orient="records")
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500
