from flask import Flask, request, send_file, jsonify
import pandas as pd
import io
import base64
import json

app = Flask(__name__)

@app.route("/generate", methods=["POST"])
def generate():

    if not request.data:
        return jsonify({"error": "No file uploaded"}), 400

    try:

        file_bytes = None

        # Try parsing as JSON first
        try:
            body = json.loads(request.data)

            if isinstance(body, dict) and "$content" in body:
                file_bytes = base64.b64decode(body["$content"])

        except Exception:
            pass

        # If not JSON, assume raw binary
        if file_bytes is None:
            file_bytes = request.data

        file_stream = io.BytesIO(file_bytes)

        df = pd.read_csv(file_stream)

    except Exception as e:
        return jsonify({"error": str(e)}), 400

    # Example analysis
    df['Total'] = df['qty'] * df['cost']

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Analysis')

    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="analysis.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/")
def home():
    return "API running"
