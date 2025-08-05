
from flask import Flask, request, jsonify
import os
import pandas as pd
import requests

app = Flask(__name__)

# === CONFIG ===
folder_path = r'C:/Users/ECrofts.TAFTELECTRIC/OneDrive - Taft Electric Company/CTC_Exports'
sheet_id = '1173339876839300'

def push_ctc_to_smartsheet():
    api_token = os.getenv("SMARTSHEET_API_TOKEN")
    headers = {
        'Authorization': f'Bearer {api_token}',
        'Content-Type': 'application/json'
    }

    # Find latest Excel file
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    if not excel_files:
        return "❌ No Excel files found."

    latest_file = max(excel_files, key=lambda f: os.path.getctime(os.path.join(folder_path, f)))
    latest_path = os.path.join(folder_path, latest_file)
    print(f"📂 Using file: {latest_file}")

    # Load Data
    df = pd.read_excel(latest_path, sheet_name='Data Shuttle Export')

    # Get sheet columns
    sheet_info = requests.get(f'https://api.smartsheet.com/2.0/sheets/{sheet_id}', headers=headers)
    columns = sheet_info.json()['columns']
    column_map = {col['title']: col['id'] for col in columns}

    # Build rows
    rows = []
    for _, row in df.iterrows():
        if pd.isna(row.get("Phase Code")):
            continue
        cells = []
        for col_name in ['Phase Code', 'Task Description', 'Budgeted', 'Start Date', 'End Date']:
            if col_name in row and pd.notna(row[col_name]):
                cells.append({ 'columnId': column_map[col_name], 'value': row[col_name] })
        rows.append({ 'toBottom': True, 'cells': cells })

    # Push to Smartsheet
    url = f'https://api.smartsheet.com/2.0/sheets/{sheet_id}/rows'
    response = requests.post(url, headers=headers, json={'rows': rows})

    if response.status_code == 200:
        return "✅ Upload complete."
    else:
        return f"❌ Upload failed: {response.text}"

@app.route('/trigger', methods=['POST'])
def trigger_upload():
    result = push_ctc_to_smartsheet()
    return jsonify({"status": result})

@app.route('/', methods=['GET'])
def health_check():
    return "✅ Smartsheet webhook is live."

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
