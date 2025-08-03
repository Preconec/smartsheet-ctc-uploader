
import os
import pandas as pd
import requests

# === CONFIG ===
folder_path = r'C:/Users/ECrofts.TAFTELECTRIC/OneDrive - Taft Electric Company/CTC_Exports'
api_token = os.getenv("SMARTSHEET_API_TOKEN")
sheet_id = '1173339876839300'

# === GET LATEST FILE ===
excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
latest_file = max(excel_files, key=lambda f: os.path.getctime(os.path.join(folder_path, f)))
latest_path = os.path.join(folder_path, latest_file)

print(f'📂 Using file: {latest_file}')

# === LOAD DATA ===
df = pd.read_excel(latest_path, sheet_name='Data Shuttle Export')

# === CONNECT TO SMARTSHEET ===
headers = {
    'Authorization': f'Bearer {api_token}',
    'Content-Type': 'application/json'
}

# Get Smartsheet column IDs
sheet_info = requests.get(f'https://api.smartsheet.com/2.0/sheets/{sheet_id}', headers=headers)
columns = sheet_info.json()['columns']
column_map = {col['title']: col['id'] for col in columns}

# === BUILD ROWS ===
rows = []
for _, row in df.iterrows():
    if pd.isna(row.get("Phase Code")):
        continue
    cells = []
    for col_name in ['Phase Code', 'Task Description', 'Budgeted', 'Start Date', 'End Date']:
        if col_name in row and pd.notna(row[col_name]):
            cells.append({ 'columnId': column_map[col_name], 'value': row[col_name] })
    rows.append({ 'toBottom': True, 'cells': cells })

# === PUSH TO SMARTSHEET ===
url = f'https://api.smartsheet.com/2.0/sheets/{sheet_id}/rows'
response = requests.post(url, headers=headers, json={'rows': rows})

# === RESULT ===
if response.status_code == 200:
    print("✅ Upload complete.")
else:
    print("❌ Upload failed:", response.text)
