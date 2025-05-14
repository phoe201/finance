# app.py

from flask import Flask, request, jsonify
from claims_processor import process_claims_excel
from flask_cors import CORS

app = Flask(__name__)
CORS(app)  # Allow cross-origin calls from Power Apps

@app.route('/process', methods=['POST'])
def process():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    try:
        file = request.files['file']
        file_bytes = file.read()
        processed_data = process_claims_excel(file_bytes)
        return jsonify(processed_data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)
