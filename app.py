from flask import Flask, request, render_template
import pandas as pd
import os

app = Flask(__name__)

excel_file = 'tool_breakage_responses.xlsx'

@app.route('/')
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
        "Reason": ", ".join(request.form.getlist("reason")),
        "Explanation": request.form.get("explanation"),
        "Breakdown Time": request.form.get("breakdown_time"),
        "Cost of Tool": request.form.get("cost"),
        "Operator": request.form.get("operator"),
        "Supervisor": request.form.get("supervisor"),
        "Tooling Engineer": request.form.get("engineer")
    }

    df = pd.DataFrame([data])

    if not os.path.exists(excel_file):
        df.to_excel(excel_file, index=False)
    else:
        old_df = pd.read_excel(excel_file)
        new_df = pd.concat([old_df, df], ignore_index=True)
        new_df.to_excel(excel_file, index=False)

    return "✅ Response recorded successfully!"

if __name__ == '__main__':
    app.run(debug=True)
