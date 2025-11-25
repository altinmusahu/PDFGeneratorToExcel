from flask import Flask, render_template, request, send_file
import pandas as pd
from PyPDF2 import PdfReader
import io
import re

app = Flask(__name__)

def extract_case_info_from_form(file):
    reader = PdfReader(file)
    fields = reader.get_fields()
    if not fields:
        return {"Case Name": "", "Street": "", "City": ""}

    name = ""
    street = ""
    city = ""

    for key, field in fields.items():
        value = field.get("/V")
        if not value:
            continue
        key_lower = key.lower()

        if "case" in key_lower and "name" in key_lower:
            name = str(value)
        elif "street" in key_lower or "street" in key_lower:
            street = str(value)
        elif "city" in key_lower:
            city = str(value)

    return {
        "Case Name": name.strip(),
        "Street": street.strip(),
        "City": city.strip(),
    }

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    files = request.files.getlist('pdfs')
    extracted_data = []

    for file in files:
        case_info = extract_case_info_from_form(file)

        name = case_info.get("Case Name")
        street = case_info.get("Street")
        city = case_info.get("City")
    
        extracted_data.append({
            "Name": name,
            "Street": street,
            "City": city
        })

    # Convert to Excel
    df = pd.DataFrame(extracted_data)
    output = io.BytesIO()
    df.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)

    return send_file(
        output,
        download_name="output.xlsx",
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    print(f"Starting Flask app on port {port}")
    app.run(host="0.0.0.0", port=port)
