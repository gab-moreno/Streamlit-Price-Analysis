from flask import Flask, request, send_file, jsonify
import pandas as pd
import io
import base64
import json

app = Flask(__name__)

@app.route("/generate", methods=["POST"])
def generate():
    raw_data = request.get_data(force=True)
    print(f"Content-Type: {request.content_type}")
    print(f"Raw data length: {len(raw_data)}")
    print(f"Raw data preview: {raw_data[:200]}")

    if not raw_data:
        return jsonify({"error": "No file uploaded"}), 400

    try:
        file_bytes = None
        try:
            body = json.loads(raw_data)
            if isinstance(body, dict) and "$content" in body:
                file_bytes = base64.b64decode(body["$content"])
        except Exception:
            pass

        if file_bytes is None:
            file_bytes = raw_data

        file_stream = io.BytesIO(file_bytes)
        df = pd.read_csv(file_stream)
    except Exception as e:
        return jsonify({"error": str(e)}), 400

    df['Total'] = df['qty'] * df['cost']

    return send_file(
        output,
        as_attachment=True,
        download_name="analysis.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/")
def home():
    return "API running"
