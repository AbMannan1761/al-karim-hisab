import os
import json
import requests

WORKSPACE_DIR = r"e:\user\OneDrive - Bangladesh Telecommunication Regulatory Commission\ABM\Sunnah\AL karim hisab"
JSON_DIR = os.path.join(WORKSPACE_DIR, "data", "json")
SETTINGS_PATH = os.path.join(WORKSPACE_DIR, "settings.json")

def load_settings():
    if os.path.exists(SETTINGS_PATH):
        with open(SETTINGS_PATH, "r") as f:
            return json.load(f)
    return {}

def main():
    settings = load_settings()
    url = settings.get("google_apps_script_url", "")
    
    if not url or url == "YOUR_APPS_SCRIPT_WEB_APP_URL_HERE":
        print("Error: Please configure google_apps_script_url in settings.json first.")
        return
        
    print("Fetching current database from Google Sheets...")
    try:
        response = requests.get(url)
        if response.status_code != 200:
            print(f"HTTP Error: {response.status_code}")
            return
            
        data = response.json()
    except Exception as e:
        print(f"Failed to fetch data: {e}")
        return
        
    client_index = data.get("Client_Index", [])
    debit_transactions = data.get("Debit_Transactions", [])
    credit_transactions = data.get("Credit_Transactions", [])
    
    print(f"Received {len(client_index)} clients, {len(debit_transactions)} debits, {len(credit_transactions)} credits.")
    
    # 1. Update index_data.json
    index_list = []
    page_to_client = {}  # Map page to client info
    
    for row in client_index:
        p_num = int(row.get("Ledger Page", 0))
        party_name = row.get("Party Name", "")
        if p_num:
            page_to_client[p_num] = {
                "party_name": party_name,
                "address": row.get("Address", ""),
                "notes": row.get("Notes", "")
            }
        index_list.append({
            "no": int(row.get("Client No", 0)) if row.get("Client No") else "",
            "party_name": party_name,
            "page": p_num,
            "address": row.get("Address", ""),
            "notes": row.get("Notes", "")
        })
        
    # Write index_data.json
    with open(os.path.join(JSON_DIR, "index_data.json"), "w", encoding="utf-8") as f:
        json.dump(index_list, f, indent=2, ensure_ascii=False)
    print("Saved index_data.json")
    
    # 2. Group transactions by page
    debits_by_page = {}
    credits_by_page = {}
    
    for row in debit_transactions:
        p_num = int(row.get("Ledger Page", 0))
        if p_num:
            if p_num not in debits_by_page:
                debits_by_page[p_num] = []
            debits_by_page[p_num].append({
                "no": row.get("No", ""),
                "date": row.get("Date", ""),
                "bi_ka": row.get("Details", row.get("বিবরণ", "")),
                "description": row.get("Description", row.get("কাপড়", "")),
                "size": row.get("Size", ""),
                "model": row.get("Model", ""),
                "pd": row.get("PD", row.get("মোট", "")),
                "bill": row.get("Bill No", ""),
                "qty": row.get("Qty", ""),
                "taka": row.get("Rate", ""),
                "total": row.get("Total", row.get("সর্বশেষ বিল", "")),
                "remarks": row.get("Remarks", "")
            })
            
    for row in credit_transactions:
        p_num = int(row.get("Ledger Page", 0))
        if p_num:
            if p_num not in credits_by_page:
                credits_by_page[p_num] = []
            credits_by_page[p_num].append({
                "no": row.get("No", ""),
                "date": row.get("Date", ""),
                "amount": row.get("Amount", ""),
                "remarks": row.get("Remarks", "")
            })
            
    # 3. Regenerate page JSON files
    # Find all relevant PDF page numbers (page // 2)
    all_pages = set(list(page_to_client.keys()) + list(debits_by_page.keys()) + list(credits_by_page.keys()))
    pdf_pages = set(p // 2 for p in all_pages)
    
    for pdf_page in pdf_pages:
        left_page = pdf_page * 2
        right_page = left_page + 1
        
        # Left Page (Debit)
        left_client = page_to_client.get(left_page, {})
        left_data = {
            "party_name": left_client.get("party_name", ""),
            "phone": "",  # phone isn't explicitly in index, can leave empty or defaults
            "ledger_page_number": str(left_page),
            "debit_table": debits_by_page.get(left_page, [])
        }
        with open(os.path.join(JSON_DIR, f"page_{pdf_page}_left.json"), "w", encoding="utf-8") as f:
            json.dump(left_data, f, indent=2, ensure_ascii=False)
            
        # Right Page (Credit)
        right_client = page_to_client.get(right_page, {})
        right_data = {
            "party_name": right_client.get("party_name", ""),
            "phone": "",
            "ledger_page_number": str(right_page),
            "credit_table": credits_by_page.get(right_page, [])
        }
        with open(os.path.join(JSON_DIR, f"page_{pdf_page}_right.json"), "w", encoding="utf-8") as f:
            json.dump(right_data, f, indent=2, ensure_ascii=False)
            
    print(f"Successfully backed up all {len(pdf_pages)} page files.")

if __name__ == "__main__":
    main()
