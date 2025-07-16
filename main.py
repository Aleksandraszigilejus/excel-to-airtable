
from flask import Flask, request, jsonify
import pandas as pd
from io import BytesIO

app = Flask(__name__)

@app.route("/process", methods=["POST"])
def process_excel():
    if 'file' not in request.files:
        return "No file part", 400

    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400

    try:
        df = pd.read_excel(file, sheet_name="Užsakymas")

        # Dummy response for test purpose
        return jsonify({
            "columns": df.columns.tolist(),
            "rows": df.to_dict(orient="records")
        })

    except Exception as e:
        return f"Error processing file: {str(e)}", 500

if __name__ == "__main__":
    app.run(debug=True)
