import os
import json
import requests

WORKSPACE_DIR = r"e:\user\OneDrive - Bangladesh Telecommunication Regulatory Commission\ABM\Sunnah\AL karim hisab"
DATA_DIR = os.path.join(WORKSPACE_DIR, "data")
JSON_DIR = os.path.join(DATA_DIR, "json")
INDEX_PATH = os.path.join(JSON_DIR, "index_data.json")
SETTINGS_PATH = os.path.join(WORKSPACE_DIR, "settings.json")

def load_settings():
    if os.path.exists(SETTINGS_PATH):
        with open(SETTINGS_PATH, "r") as f:
            return json.load(f)
    return {}

def compile_data():
    if not os.path.exists(INDEX_PATH):
        print("Error: index_data.json not found!")
        return None, None, None
        
    with open(INDEX_PATH, "r", encoding="utf-8") as f:
        index_list = json.load(f)
        
    index_data = []
    debit_data = []
    credit_data = []
    
    # Track mapping of ledger page to client name
    ledger_page_to_name = {}
    for entry in index_list:
        p_num = int(entry.get("page", 0))
        if p_num:
            ledger_page_to_name[p_num] = entry.get("party_name", "")
            
    print("Compiling transactions from JSON files...")
    for entry in index_list:
        p_num = int(entry.get("page", 0))
        pdf_page_num = p_num // 2
        name = entry.get("party_name", "")
        
        index_data.append({
            "no": entry.get("no", ""),
            "party_name": name,
            "address": entry.get("address", ""),
            "page": p_num,
            "pdf_page": pdf_page_num,
            "notes": entry.get("notes", "")
        })
        
        # Left (Debit) Page
        left_file = os.path.join(JSON_DIR, f"page_{pdf_page_num}_left.json")
        if os.path.exists(left_file):
            with open(left_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            t_name = data.get("party_name", name) or name
            for row in data.get("debit_table", []):
                debit_data.append({
                    "party_name": t_name,
                    "ledger_page": data.get("ledger_page_number", str(p_num)),
                    "no": row.get("no", ""),
                    "date": row.get("date", ""),
                    "bi_ka": row.get("bi_ka", ""),
                    "description": row.get("description", ""),
                    "size": row.get("size", ""),
                    "model": row.get("model", ""),
                    "pd": row.get("pd", ""),
                    "bill": row.get("bill", ""),
                    "qty": row.get("qty", ""),
                    "taka": row.get("taka", ""),
                    "total": row.get("total", ""),
                    "remarks": row.get("remarks", "")
                })
                
        # Right (Credit) Page
        right_file = os.path.join(JSON_DIR, f"page_{pdf_page_num}_right.json")
        if os.path.exists(right_file):
            with open(right_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            t_name = data.get("party_name", name) or name
            for row in data.get("credit_table", []):
                credit_data.append({
                    "party_name": t_name,
                    "ledger_page": data.get("ledger_page_number", str(p_num + 1)),
                    "no": row.get("no", ""),
                    "date": row.get("date", ""),
                    "amount": row.get("amount", ""),
                    "remarks": row.get("remarks", "")
                })
                
    return index_data, debit_data, credit_data

def main():
    settings = load_settings()
    url = settings.get("google_apps_script_url", "")
    
    if not url or url == "YOUR_APPS_SCRIPT_WEB_APP_URL_HERE":
        print("\n[ALERT] Please update settings.json with your deployed Google Apps Script Web App URL first!")
        return
        
    index_data, debit_data, credit_data = compile_data()
    if index_data is None:
        return
        
    payload = {
        "action": "bulk_init",
        "index_data": index_data,
        "debit_data": debit_data,
        "credit_data": credit_data
    }
    
    print(f"Uploading {len(index_data)} clients, {len(debit_data)} sales bills, and {len(credit_data)} payments...")
    try:
        response = requests.post(url, json=payload)
        if response.status_code == 200:
            try:
                res_json = response.json()
                if res_json.get("status") == "success":
                    print("SUCCESS: Google Sheets database populated successfully!")
                else:
                    print(f"Error: {res_json.get('message')}")
            except ValueError as je:
                print(f"Upload failed to decode JSON response: {je}")
                with open("error_response.html", "w", encoding="utf-8") as ef:
                    ef.write(response.text)
                print("Saved the server's HTML response to 'error_response.html' for inspection.")
        else:
            print(f"HTTP Error: {response.status_code}")
            with open("error_response.html", "w", encoding="utf-8") as ef:
                ef.write(response.text)
            print("Saved the server's HTML response to 'error_response.html' for inspection.")
    except Exception as e:
        print(f"Upload failed: {e}")

if __name__ == "__main__":
    main()
