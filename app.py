from flask import Flask, render_template, request, send_file
import pandas as pd
import os

app = Flask(__name__)

EXCEL_FILE = "responses.xlsx"

@app.route('/', methods=['GET'])
def form():
    return render_template('form.html')

@app.route('/submit', methods=['POST'])
def submit():
    data = {
        "Name": request.form.get("name"),
        "Department": request.form.get("department"),
        "Item Code": request.form.get("item_code"),
        "Description": request.form.get("description"),
        "Project": request.form.get("project"),
        "Part No": request.form.get("part_no"),
        "Operation": request.form.get("operation"),
        "Reason for Breakage": ", ".join(request.form.getlist("reason")),
        "Explanation": request.form.get("explanation"),
        "Breakdown Time": request.form.get("breakdown_time"),
        "Cost of Tool": request.form.get("cost"),
        "Operator": request.form.get("operator"),
        "Supervisor": request.form.get("supervisor"),
        "Tooling Engineer": request.form.get("engineer")
    }

    # Convert to DataFrame
    new_entry = pd.DataFrame([data])

    # Append to Excel file
    if os.path.exists(EXCEL_FILE):
        existing_data = pd.read_excel(EXCEL_FILE)
        combined = pd.concat([existing_data, new_entry], ignore_index=True)
    else:
        combined = new_entry

    combined.to_excel(EXCEL_FILE, index=False)

    return "✅ Submitted successfully!"

@app.route('/download')
def download():
    if os.path.exists(EXCEL_FILE):
        return send_file(EXCEL_FILE, as_attachment=Tru_

