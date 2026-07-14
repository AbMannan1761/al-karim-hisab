import os
import json
import requests
import pandas as pd
import streamlit as st
from datetime import datetime
# from export_to_excel import export

# Paths
WORKSPACE_DIR = r"e:\user\OneDrive - Bangladesh Telecommunication Regulatory Commission\ABM\Sunnah\AL karim hisab"
DATA_DIR = os.path.join(WORKSPACE_DIR, "data")
JSON_DIR = os.path.join(DATA_DIR, "json")
INDEX_PATH = os.path.join(JSON_DIR, "index_data.json")
SETTINGS_PATH = os.path.join(WORKSPACE_DIR, "settings.json")

# Page config
st.set_page_config(
    page_title="Al Karim Hisab Dashboard",
    page_icon="📖",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling
st.markdown("""
<style>
    .kpi-card {
        padding: 20px;
        border-radius: 10px;
        background-color: #f8f9fa;
        border: 1px solid #e9ecef;
        text-align: center;
    }
    .kpi-val {
        font-size: 24px;
        font-weight: bold;
        color: #366092;
    }
    .kpi-title {
        font-size: 14px;
        color: #6c757d;
        text-transform: uppercase;
        margin-bottom: 5px;
    }
</style>
""", unsafe_allow_html=True)

# Helper to clean numbers
def clean_num(val):
    if val is None:
        return 0.0
    val_str = str(val).strip().replace(",", "")
    if not val_str or val_str == "〃" or val_str == '"':
        return 0.0
    try:
        return float(val_str)
    except ValueError:
        return 0.0

# ----------------- CONFIG & MODE CHECK -----------------
def get_sheets_url():
    if os.path.exists(SETTINGS_PATH):
        try:
            with open(SETTINGS_PATH, "r") as f:
                data = json.load(f)
                url = data.get("google_apps_script_url", "")
                if url and url != "YOUR_APPS_SCRIPT_WEB_APP_URL_HERE":
                    return url
        except Exception:
            pass
    return None

SHEETS_URL = get_sheets_url()

# ----------------- GOOGLE SHEETS API DATA LAYER -----------------
@st.cache_data(ttl=5) # Cache data for 5 seconds to keep it snappy but fresh
def fetch_sheets_data(url):
    try:
        res = requests.get(url, timeout=10)
        if res.status_code == 200:
            return res.json()
    except Exception as e:
        st.sidebar.error(f"গুগল শিটের সাথে সংযোগ করা যায়নি: {e}")
    return None

def send_sheets_post(url, payload):
    try:
        res = requests.post(url, json=payload, timeout=10)
        if res.status_code == 200:
            # Clear Streamlit cache to force reload fresh data
            fetch_sheets_data.clear()
            return True
    except Exception as e:
        st.error(f"গুগল শিটে ডেটা পাঠাতে সমস্যা হয়েছে: {e}")
    return False

# ----------------- LOCAL DATA LAYER -----------------
@st.cache_data
def load_local_index():
    if os.path.exists(INDEX_PATH):
        with open(INDEX_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

@st.cache_data
def load_local_page_data(pdf_page_num):
    left_file = os.path.join(JSON_DIR, f"page_{pdf_page_num}_left.json")
    right_file = os.path.join(JSON_DIR, f"page_{pdf_page_num}_right.json")
    
    left_data = {"party_name": "", "phone": "", "ledger_page_number": str(2 * pdf_page_num), "debit_table": []}
    right_data = {"party_name": "", "phone": "", "ledger_page_number": str(2 * pdf_page_num + 1), "credit_table": []}
    
    if os.path.exists(left_file):
        with open(left_file, "r", encoding="utf-8") as f:
            left_data = json.load(f)
            
    if os.path.exists(right_file):
        with open(right_file, "r", encoding="utf-8") as f:
            right_data = json.load(f)
            
    return left_data, right_data

def save_local_page_data(pdf_page_num, side, data):
    file_path = os.path.join(JSON_DIR, f"page_{pdf_page_num}_{side}.json")
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    # Clear local page data cache to force reload fresh data
    load_local_page_data.clear()
    # Regenerate local Excel (Disabled: using Google Sheets mode primary)
    # try:
    #     export()
    # except Exception as e:
    #     st.error(f"Error regenerating Excel sheet: {e}")

# ----------------- DATA PREPARATION -----------------
index_data = []
all_debit_list = []
all_credit_list = []

if SHEETS_URL:
    st.sidebar.success("⚡ গুগল শিট সিঙ্ক মোড চালু আছে (Google Sheets Mode)")
    sheets_db = fetch_sheets_data(SHEETS_URL)
    if sheets_db:
        # Load Client Index
        for row in sheets_db.get("Client_Index", []):
            index_data.append({
                "no": str(row.get("Client No", "")),
                "party_name": str(row.get("Party Name", "")),
                "address": str(row.get("Address", "")),
                "page": int(clean_num(row.get("Ledger Page", 0))),
                "notes": str(row.get("Notes", ""))
            })
        
        # Load Debit Transactions
        all_debit_list = sheets_db.get("Debit_Transactions", [])
        # Load Credit Transactions
        all_credit_list = sheets_db.get("Credit_Transactions", [])
    else:
        st.error("গুগল শিট থেকে ডেটা লোড করা যায়নি। সাময়িকভাবে লোকাল মোডে কাজ করুন।")
        SHEETS_URL = None # Fallback to local

if not SHEETS_URL:
    st.sidebar.warning("💾 লোকাল অফলাইন মোড চালু আছে (Local Offline Mode)")
    local_idx = load_local_index()
    for row in local_idx:
        index_data.append({
            "no": str(row.get("no", "")),
            "party_name": str(row.get("party_name", "")),
            "address": str(row.get("address", "")),
            "page": int(clean_num(row.get("page", 0))),
            "notes": str(row.get("notes", ""))
        })

# Application header
st.title("📖 আল করিম হিসাব (Al Karim Hisab)")
st.subheader("হ্যান্ডরিটেন লেজার বুক অটোমেশন সিস্টেম")

# Initialize session states
if "current_panel" not in st.session_state:
    st.session_state.current_panel = "📈 সর্বমোট সারসংক্ষেপ (General Summary)"
if "selected_client" not in st.session_state:
    st.session_state.selected_client = index_data[0] if index_data else None

# Sidebar Navigation Panel
st.sidebar.markdown("## 🧭 মেনু নেভিগেশন (Menu)")
panel_options = [
    "📈 সর্বমোট সারসংক্ষেপ (General Summary)",
    "📊 গ্রাহক খতিয়ান (Client Ledger Sheet)",
    "✍️ নতুন ডাটা ইনপুট প্যানেল (Data Input Panel)",
    "🛠️ সংশোধন প্যানেল (Correction Panel)"
]

# Render the panel selector linked to session state
st.session_state.current_panel = st.sidebar.selectbox(
    "প্যানেল নির্বাচন করুন (Select Panel):",
    options=panel_options,
    index=panel_options.index(st.session_state.current_panel),
    key="panel_selector"
)

st.sidebar.divider()

# Sidebar Client Selection list (Always Visible)
st.sidebar.markdown("## 👥 গ্রাহক তালিকা (Client List)")
sidebar_search = st.sidebar.text_input("গ্রাহক খুঁজুন (Search Client):", "", placeholder="নাম বা ঠিকানা লিখুন...")

# Filter index based on query
filtered_index = index_data
if sidebar_search:
    filtered_index = [
        x for x in index_data 
        if sidebar_search.lower() in x.get("party_name", "").lower() or 
           sidebar_search.lower() in x.get("address", "").lower()
    ]

# Client options list
client_options = []
for entry in filtered_index:
    name = entry.get("party_name", "Unknown")
    page = entry.get("page", 0)
    address = entry.get("address", "")
    client_options.append({
        "label": f"পৃষ্ঠা: {page} - {name} ({address})",
        "entry": entry
    })

# Callback for when the client is selected
def handle_client_selection():
    st.session_state.current_panel = "📊 গ্রাহক খতিয়ান (Client Ledger Sheet)"
    st.session_state.panel_selector = "📊 গ্রাহক খতিয়ান (Client Ledger Sheet)"

if client_options:
    default_idx = 0
    if "selected_client_label" in st.session_state:
        labels = [x["label"] for x in client_options]
        if st.session_state.selected_client_label in labels:
            default_idx = labels.index(st.session_state.selected_client_label)

    selected_client_option = st.sidebar.selectbox(
        "গ্রাহক সিলেক্ট করুন (Select Client):",
        options=client_options,
        index=default_idx,
        format_func=lambda x: x["label"],
        key="global_client_selectbox",
        on_change=handle_client_selection
    )
    st.session_state.selected_client = selected_client_option["entry"]
    st.session_state.selected_client_label = selected_client_option["label"]
else:
    st.sidebar.warning("কোনো গ্রাহক পাওয়া যায়নি।")
    st.session_state.selected_client = None

# Active Page Routing
menu = st.session_state.current_panel

# ----------------- 1. GENERAL SUMMARY PANEL -----------------
if menu == "📈 সর্বমোট সারসংক্ষেপ (General Summary)":
    st.header("📊 সর্বমোট ব্যবসার সারসংক্ষেপ (Business Summary)")
    
    global_debit = 0.0
    global_credit = 0.0
    summary_rows = []
    
    if SHEETS_URL:
        # Pre-sum tables from Google Sheets data in bulk
        debit_sums = {}
        for row in all_debit_list:
            name = row.get("Client Name", "")
            total = clean_num(row.get("Total", 0))
            debit_sums[name] = debit_sums.get(name, 0.0) + total
            
        credit_sums = {}
        for row in all_credit_list:
            name = row.get("Client Name", "")
            amount = clean_num(row.get("Amount", 0))
            credit_sums[name] = credit_sums.get(name, 0.0) + amount
            
        for entry in index_data:
            name = entry.get("party_name", "")
            p_num = entry.get("page", 0)
            debit_sum = debit_sums.get(name, 0.0)
            credit_sum = credit_sums.get(name, 0.0)
            
            global_debit += debit_sum
            global_credit += credit_sum
            
            summary_rows.append({
                "Client No": entry.get("no", ""),
                "Name (নাম)": name,
                "Address (ঠিকানা)": entry.get("address", ""),
                "Ledger Page (পৃষ্ঠা)": p_num,
                "Total Sales (মোট ক্রয়) ৳": debit_sum,
                "Total Payments (মোট পরিশোধ) ৳": credit_sum,
                "Balance (বকেয়া) ৳": debit_sum - credit_sum,
                "Notes (মন্তব্য)": entry.get("notes", "")
            })
    else:
        # Local file sums
        for entry in index_data:
            p_num = int(entry.get("page", 0))
            pdf_page_num = p_num // 2
            left, right = load_local_page_data(pdf_page_num)
            
            debit_sum = sum(clean_num(row.get("total")) for row in left.get("debit_table", []))
            credit_sum = sum(clean_num(row.get("amount")) for row in right.get("credit_table", []))
            
            global_debit += debit_sum
            global_credit += credit_sum
            
            summary_rows.append({
                "Client No": entry.get("no", ""),
                "Name (নাম)": entry.get("party_name", ""),
                "Address (ঠিকানা)": entry.get("address", ""),
                "Ledger Page (পৃষ্ঠা)": p_num,
                "Total Sales (মোট ক্রয়) ৳": debit_sum,
                "Total Payments (মোট পরিশোধ) ৳": credit_sum,
                "Balance (বকেয়া) ৳": debit_sum - credit_sum,
                "Notes (মন্তব্য)": entry.get("notes", "")
            })
            
    global_balance = global_debit - global_credit
    
    # KPI Display
    kcol1, kcol2, kcol3 = st.columns(3)
    with kcol1:
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Global Sales (সর্বমোট বিক্রয়)</div><div class="kpi-val">{global_debit:,.2f} ৳</div></div>', unsafe_allow_html=True)
    with kcol2:
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Global Payments (সর্বমোট আদায়)</div><div class="kpi-val" style="color: #2e7d32;">{global_credit:,.2f} ৳</div></div>', unsafe_allow_html=True)
    with kcol3:
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Total Receivables (সর্বমোট বকেয়া পাওনা)</div><div class="kpi-val" style="color: {"#c62828" if global_balance > 0 else "#2e7d32"};">{global_balance:,.2f} ৳</div></div>', unsafe_allow_html=True)
        
    st.write("")
    st.subheader("📋 সকল গ্রাহকের হিসাব তালিকা")
    
    df_summary = pd.DataFrame(summary_rows)
    table_search = st.text_input("তালিকা ফিল্টার করুন (নাম/ঠিকানা লিখুন):", "")
    if table_search:
        df_summary = df_summary[
            df_summary["Name (নাম)"].str.contains(table_search, case=False, na=False) |
            df_summary["Address (ঠিকানা)"].str.contains(table_search, case=False, na=False)
        ]
        
    st.dataframe(
        df_summary,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Total Sales (মোট ক্রয়) ৳": st.column_config.NumberColumn(format="%.2f"),
            "Total Payments (মোট পরিশোধ) ৳": st.column_config.NumberColumn(format="%.2f"),
            "Balance (বকেয়া) ৳": st.column_config.NumberColumn(format="%.2f")
        }
    )

# ----------------- 2. CLIENT LEDGER VIEW PANEL -----------------
elif menu == "📊 গ্রাহক খতিয়ান (Client Ledger Sheet)":
    st.header("📊 গ্রাহক খতিয়ান ভিউয়ার (Client Ledger Viewer)")
    
    selected_client = st.session_state.selected_client
    if selected_client is None:
        st.info("দয়া করে সাইডবার (বামপাশ) থেকে একটি গ্রাহক সিলেক্ট করুন।")
        st.stop()
        
    ledger_page = int(selected_client.get("page", 0))
    
    debit_list = []
    credit_list = []
    phone_number = "N/A"
    
    if SHEETS_URL:
        # Filter in-memory from Google Sheets database
        name = selected_client.get("party_name", "")
        
        # We need to clean and fetch phone from Client_Index notes or page data
        phone_number = "N/A"
        
        debit_rows = [x for x in all_debit_list if x.get("Client Name") == name]
        for row in debit_rows:
            debit_list.append({
                "no": str(row.get("No", "")),
                "date": str(row.get("Date", "")),
                "bi_ka": str(row.get("Details", "")),
                "description": str(row.get("Description", "")),
                "size": str(row.get("Size", "")),
                "model": str(row.get("Model", "")),
                "pd": str(row.get("PD", "")),
                "bill": str(row.get("Bill No", "")),
                "qty": str(row.get("Qty", "")),
                "taka": str(row.get("Rate", "")),
                "total": str(row.get("Total", "")),
                "remarks": str(row.get("Remarks", ""))
            })
            
        credit_rows = [x for x in all_credit_list if x.get("Client Name") == name]
        for row in credit_rows:
            credit_list.append({
                "no": str(row.get("No", "")),
                "date": str(row.get("Date", "")),
                "amount": str(row.get("Amount", "")),
                "remarks": str(row.get("Remarks", ""))
            })
    else:
        # Load locally
        pdf_page_num = ledger_page // 2
        left_data, right_data = load_local_page_data(pdf_page_num)
        phone_number = left_data.get("phone", "N/A")
        debit_list = left_data.get("debit_table", [])
        credit_list = right_data.get("credit_table", [])
        
    st.write("")
    # Banner
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(f"👤 **গ্রাহকের নাম:** {selected_client.get('party_name', 'N/A')}")
        st.markdown(f"📍 **ঠিকানা:** {selected_client.get('address', 'N/A')}")
    with col2:
        st.markdown(f"📄 **লেজার পৃষ্ঠা:** {ledger_page}")
        st.markdown(f"📞 **ফোন নম্বর:** {phone_number}")
    with col3:
        st.markdown(f"📝 **মন্তব্য:** {selected_client.get('notes', 'N/A')}")
        
    st.divider()
    
    total_debit = sum(clean_num(row.get("total")) for row in debit_list)
    total_credit = sum(clean_num(row.get("amount")) for row in credit_list)
    outstanding_balance = total_debit - total_credit
    
    # Client KPIs
    kcol1, kcol2, kcol3 = st.columns(3)
    with kcol1:
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Total Purchases (মোট ক্রয়)</div><div class="kpi-val">{total_debit:,.2f} ৳</div></div>', unsafe_allow_html=True)
    with kcol2:
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Total Payments (মোট পরিশোধ)</div><div class="kpi-val" style="color: #2e7d32;">{total_credit:,.2f} ৳</div></div>', unsafe_allow_html=True)
    with kcol3:
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Balance / Outstanding (বকেয়া)</div><div class="kpi-val" style="color: {"#c62828" if outstanding_balance > 0 else "#2e7d32"};">{outstanding_balance:,.2f} ৳</div></div>', unsafe_allow_html=True)
        
    st.write("")
    
    # Transaction Tables
    col_left, col_right = st.columns([7, 3])
    with col_left:
        st.subheader("📈 ডেবিট এন্ট্রি সমূহ (Debit - Bills/Sales)")
        if debit_list:
            df_debit = pd.DataFrame(debit_list)
            cols = ["no", "date", "bi_ka", "description", "size", "model", "qty", "taka", "total", "remarks"]
            existing_cols = [c for c in cols if c in df_debit.columns]
            st.dataframe(df_debit[existing_cols], use_container_width=True, hide_index=True)
        else:
            st.info("কোনো বিক্রয় রেকর্ড পাওয়া যায়নি।")
            
    with col_right:
        st.subheader("📉 ক্রেডিট এন্ট্রি সমূহ (Credit - Payments)")
        if credit_list:
            df_credit = pd.DataFrame(credit_list)
            cols = ["no", "date", "amount", "remarks"]
            existing_cols = [c for c in cols if c in df_credit.columns]
            st.dataframe(df_credit[existing_cols], use_container_width=True, hide_index=True)
        else:
            st.info("কোনো পেমেন্ট রেকর্ড পাওয়া যায়নি।")

# ----------------- 3. DATA INPUT PANEL -----------------
elif menu == "✍️ নতুন ডাটা ইনপুট প্যানেল (Data Input Panel)":
    st.header("✍️ নতুন ট্রানজেকশন ডাটা ইনপুট প্যানেল (Data Input Panel)")
    
    selected_client = st.session_state.selected_client
    if selected_client is None:
        st.info("দয়া করে সাইডবার (বামপাশ) থেকে একটি গ্রাহক সিলেক্ট করুন।")
        st.stop()
        
    ledger_page = int(selected_client.get("page", 0))
    name = selected_client.get("party_name", "")
    
    debit_list = []
    credit_list = []
    
    if SHEETS_URL:
        debit_list = [x for x in all_debit_list if x.get("Client Name") == name]
        credit_list = [x for x in all_credit_list if x.get("Client Name") == name]
    else:
        pdf_page_num = ledger_page // 2
        left_data, right_data = load_local_page_data(pdf_page_num)
        debit_list = left_data.get("debit_table", [])
        credit_list = right_data.get("credit_table", [])
        
    st.write("")
    st.success(f"নির্বাচিত গ্রাহক: **{name}** | ঠিকানা: {selected_client.get('address')} | লেজার পৃষ্ঠা: {ledger_page}")
    
    entry_type = st.radio(
        "ট্রানজেকশন টাইপ নির্বাচন করুন (Select Transaction Type):",
        ["ডেবিট / বিক্রয় বিল (Debit - Sales Bill)", "ক্রেডিট / পেমেন্ট আদায় (Credit - Payment Received)"]
    )
    
    st.divider()
    
    if entry_type == "ডেবিট / বিক্রয় বিল (Debit - Sales Bill)":
        st.subheader(f"➕ নতুন বিক্রয় বিল যোগ করুন ({name})")
        
        with st.form("input_bill_form", clear_on_submit=True):
            col_b1, col_b2, col_b3 = st.columns(3)
            with col_b1:
                bill_date = st.text_input("তারিখ (Date - DD.MM.YY):", datetime.today().strftime('%d.%m.%y'))
                bill_desc = st.text_input("বিবরণ (Description):", placeholder="উদাঃ ক্রীম বা সাদা")
                bill_size = st.text_input("সাইজ (Size):", placeholder="উদাঃ 40/44")
            with col_b2:
                bill_model = st.text_input("মডেল (Model):", placeholder="উদাঃ S.M")
                bill_bi_ka = st.text_input("বিঃ কাঃ (Details):", value="B+K+H")
                bill_pd = st.text_input("পিঃ ডিঃ (PD):", "")
            with col_b3:
                bill_bill_no = st.text_input("বিল নং (Bill No):", "")
                bill_qty = st.number_input("পরিমাণ (Qty):", min_value=0.0, step=1.0, value=0.0)
                bill_rate = st.number_input("দর (Rate):", min_value=0.0, step=10.0, value=0.0)
                
            bill_remarks = st.text_input("মন্তব্য (Remarks):", "")
            
            submitted = st.form_submit_button("বিলটি সংরক্ষণ করুন (Save Bill)")
            if submitted:
                bill_total = bill_qty * bill_rate
                new_no = str(len(debit_list) + 1)
                
                qty_str = str(int(bill_qty)) if bill_qty.is_integer() else str(bill_qty)
                taka_str = str(int(bill_rate)) if bill_rate.is_integer() else str(bill_rate)
                total_str = str(int(bill_total)) if bill_total.is_integer() else str(bill_total)
                
                if SHEETS_URL:
                    payload = {
                        "action": "add_debit",
                        "party_name": name,
                        "ledger_page": ledger_page,
                        "no": new_no,
                        "date": bill_date,
                        "bi_ka": bill_bi_ka,
                        "description": bill_desc,
                        "size": bill_size,
                        "model": bill_model,
                        "pd": bill_pd,
                        "bill": bill_bill_no,
                        "qty": qty_str,
                        "taka": taka_str,
                        "total": total_str,
                        "remarks": bill_remarks
                    }
                    if send_sheets_post(SHEETS_URL, payload):
                        st.success("নতুন বিক্রয় বিল গুগল শিটে সফলভাবে যোগ করা হয়েছে!")
                        st.rerun()
                else:
                    new_row = {
                        "no": new_no, "date": bill_date, "bi_ka": bill_bi_ka, "description": bill_desc,
                        "size": bill_size, "model": bill_model, "pd": bill_pd, "bill": bill_bill_no,
                        "qty": qty_str, "taka": taka_str, "total": total_str, "remarks": bill_remarks
                    }
                    debit_list.append(new_row)
                    left_data["debit_table"] = debit_list
                    save_local_page_data(pdf_page_num, "left", left_data)
                    st.success("নতুন বিক্রয় বিল লোকাল ফাইলে সফলভাবে সংরক্ষণ করা হয়েছে!")
                    st.rerun()
                
    else:
        st.subheader(f"➕ নতুন পেমেন্ট আদায় এন্ট্রি ({name})")
        
        with st.form("input_payment_form", clear_on_submit=True):
            col_p1, col_p2 = st.columns(2)
            with col_p1:
                pay_date = st.text_input("তারিখ (Date - DD.MM.YY):", datetime.today().strftime('%d.%m.%y'))
                pay_amount = st.number_input("টাকার পরিমাণ (Amount):", min_value=0.0, step=100.0, value=0.0)
            with col_p2:
                pay_remarks = st.text_input("মন্তব্য (Remarks):", placeholder="নগদ বা ব্যাংক বা বিল পরিশোধ")
                
            submitted = st.form_submit_button("পেমেন্ট সংরক্ষণ করুন (Save Payment)")
            if submitted:
                if pay_amount <= 0:
                    st.error("দয়া করে পেমেন্টের সঠিক পরিমাণ দিন।")
                else:
                    new_no = str(len(credit_list) + 1)
                    amt_str = str(int(pay_amount)) if pay_amount.is_integer() else str(pay_amount)
                    
                    if SHEETS_URL:
                        payload = {
                            "action": "add_credit",
                            "party_name": name,
                            "ledger_page": ledger_page + 1, # Credit page is even+1
                            "no": new_no,
                            "date": pay_date,
                            "amount": amt_str,
                            "remarks": pay_remarks
                        }
                        if send_sheets_post(SHEETS_URL, payload):
                            st.success("পেমেন্ট রেকর্ড গুগল শিটে সফলভাবে জমা করা হয়েছে!")
                            st.rerun()
                    else:
                        new_row = {
                            "no": new_no, "date": pay_date, "amount": amt_str, "remarks": pay_remarks
                        }
                        credit_list.append(new_row)
                        right_data["credit_table"] = credit_list
                        save_local_page_data(pdf_page_num, "right", right_data)
                        st.success("পেমেন্ট রেকর্ড লোকাল ফাইলে সফলভাবে সংরক্ষণ করা হয়েছে!")
                        st.rerun()

# ----------------- 4. CORRECTION PANEL -----------------
elif menu == "🛠️ সংশোধন প্যানেল (Correction Panel)":
    st.header("🛠️ হিসাব সংশোধন প্যানেল (Correction Panel)")
    
    selected_client = st.session_state.selected_client
    if selected_client is None:
        st.info("দয়া করে সাইডবার (বামপাশ) থেকে একটি গ্রাহক সিলেক্ট করুন।")
        st.stop()
        
    ledger_page = int(selected_client.get("page", 0))
    name = selected_client.get("party_name", "")
    
    debit_list = []
    credit_list = []
    phone_number = "N/A"
    
    if SHEETS_URL:
        # Load sheets transactions
        debit_rows = [x for x in all_debit_list if x.get("Client Name") == name]
        for row in debit_rows:
            debit_list.append({
                "no": str(row.get("No", "")), "date": str(row.get("Date", "")), "bi_ka": str(row.get("Details", "")),
                "description": str(row.get("Description", "")), "size": str(row.get("Size", "")), "model": str(row.get("Model", "")),
                "pd": str(row.get("PD", "")), "bill": str(row.get("Bill No", "")), "qty": str(row.get("Qty", "")),
                "taka": str(row.get("Rate", "")), "total": str(row.get("Total", "")), "remarks": str(row.get("Remarks", ""))
            })
            
        credit_rows = [x for x in all_credit_list if x.get("Client Name") == name]
        for row in credit_rows:
            credit_list.append({
                "no": str(row.get("No", "")), "date": str(row.get("Date", "")),
                "amount": str(row.get("Amount", "")), "remarks": str(row.get("Remarks", ""))
            })
    else:
        pdf_page_num = ledger_page // 2
        left_data, right_data = load_local_page_data(pdf_page_num)
        phone_number = left_data.get("phone", "N/A")
        debit_list = left_data.get("debit_table", [])
        credit_list = right_data.get("credit_table", [])
        
    st.success(f"সংশোধনের জন্য নির্বাচিত গ্রাহক: **{name}** | লেজার পৃষ্ঠা: {ledger_page}")
    
    corr_type = st.radio(
        "কি সংশোধন করতে চান সিলেক্ট করুন:",
        ["১. গ্রাহকের প্রোফাইল তথ্য (Client Info)", "২. বিক্রয় বিল সংশোধন/ডিলিট (Edit/Delete Bills)", "৩. পেমেন্ট রেকর্ড সংশোধন/ডিলিট (Edit/Delete Payments)"]
    )
    
    st.divider()
    
    # 1. Edit Client Profile
    if corr_type == "১. গ্রাহকের প্রোফাইল তথ্য (Client Info)":
        st.subheader("👤 গ্রাহকের প্রোফাইল তথ্য সংশোধন")
        
        with st.form("edit_profile_form"):
            new_name = st.text_input("গ্রাহকের নাম (Party Name):", name)
            new_address = st.text_input("ঠিকানা (Address):", selected_client.get("address", ""))
            new_phone = st.text_input("ফোন নম্বর (Phone):", phone_number)
            new_notes = st.text_input("মন্তব্য (Index Notes):", selected_client.get("notes", ""))
            
            submit_profile = st.form_submit_button("পরিবর্তনসমূহ সংরক্ষণ করুন (Save Changes)")
            if submit_profile:
                if not new_name.strip():
                    st.error("গ্রাহকের নাম ফাঁকা রাখা যাবে না।")
                else:
                    if SHEETS_URL:
                        payload = {
                            "action": "edit_profile",
                            "client_no": selected_client.get("no"),
                            "party_name": new_name.strip(),
                            "address": new_address.strip(),
                            "notes": new_notes.strip()
                        }
                        if send_sheets_post(SHEETS_URL, payload):
                            st.success("গ্রাহকের প্রোফাইল গুগল শিটে সফলভাবে সংশোধন করা হয়েছে!")
                            st.rerun()
                    else:
                        # Local edit
                        for entry in index_data:
                            if entry.get("no") == selected_client.get("no"):
                                entry["party_name"] = new_name.strip()
                                entry["address"] = new_address.strip()
                                entry["notes"] = new_notes.strip()
                                break
                        with open(INDEX_PATH, "w", encoding="utf-8") as f:
                            json.dump(index_data, f, ensure_ascii=False, indent=2)
                        load_local_index.clear()
                        
                        left_data["party_name"] = new_name.strip()
                        left_data["phone"] = new_phone.strip()
                        right_data["party_name"] = new_name.strip()
                        right_data["phone"] = new_phone.strip()
                        
                        save_local_page_data(pdf_page_num, "left", left_data)
                        save_local_page_data(pdf_page_num, "right", right_data)
                        st.success("গ্রাহকের প্রোফাইল লোকাল ফাইলে সফলভাবে সংশোধন করা হয়েছে!")
                        st.rerun()
                        
    # 2. Edit/Delete Sales Bills
    elif corr_type == "২. বিক্রয় বিল সংশোধন/ডিলিট (Edit/Delete Bills)":
        st.subheader("📈 বিক্রয় বিল সংশোধন বা ডিলিট করুন")
        if not debit_list:
            st.info("এই গ্রাহকের কোনো বিক্রয় বিল নেই।")
        else:
            bill_options = [
                {
                    "label": f"ক্রঃ নং {row.get('no')} | তারিখ: {row.get('date')} | বিবরণ: {row.get('description')} | মোট: {row.get('total')} ৳",
                    "row": row,
                    "index": idx
                }
                for idx, row in enumerate(debit_list)
            ]
            
            selected_bill_opt = st.selectbox(
                "সংশোধনের জন্য বিলটি নির্বাচন করুন (Select Bill):",
                options=bill_options,
                format_func=lambda x: x["label"]
            )
            
            bill_to_edit = selected_bill_opt["row"]
            bill_idx = selected_bill_opt["index"]
            
            st.write("---")
            with st.form("edit_bill_form"):
                col_b1, col_b2, col_b3 = st.columns(3)
                with col_b1:
                    ebill_date = st.text_input("তারিখ (Date):", bill_to_edit.get("date", ""))
                    ebill_desc = st.text_input("বিবরণ (Description):", bill_to_edit.get("description", ""))
                    ebill_size = st.text_input("সাইজ (Size):", bill_to_edit.get("size", ""))
                with col_b2:
                    ebill_model = st.text_input("মডেল (Model):", bill_to_edit.get("model", ""))
                    ebill_bi_ka = st.text_input("বিঃ কাঃ (Details):", bill_to_edit.get("bi_ka", ""))
                    ebill_pd = st.text_input("পিঃ ডিঃ (PD):", bill_to_edit.get("pd", ""))
                with col_b3:
                    ebill_bill_no = st.text_input("বিল নং (Bill No):", bill_to_edit.get("bill", ""))
                    qty_val = clean_num(bill_to_edit.get("qty"))
                    rate_val = clean_num(bill_to_edit.get("taka"))
                    ebill_qty = st.number_input("পরিমাণ (Qty):", min_value=0.0, step=1.0, value=qty_val)
                    ebill_rate = st.number_input("দর (Rate):", min_value=0.0, step=10.0, value=rate_val)
                    
                ebill_remarks = st.text_input("মন্তব্য (Remarks):", bill_to_edit.get("remarks", ""))
                
                col_btn1, col_btn2 = st.columns(2)
                with col_btn1:
                    save_bill = st.form_submit_button("বিল আপডেট করুন (Save Bill)")
                with col_btn2:
                    delete_bill = st.form_submit_button("বিলটি ডিলিট করুন (Delete Bill)", type="primary")
                    
                if save_bill:
                    ebill_total = ebill_qty * ebill_rate
                    qty_str = str(int(ebill_qty)) if ebill_qty.is_integer() else str(ebill_qty)
                    taka_str = str(int(ebill_rate)) if ebill_rate.is_integer() else str(ebill_rate)
                    total_str = str(int(ebill_total)) if ebill_total.is_integer() else str(ebill_total)
                    
                    if SHEETS_URL:
                        payload = {
                            "action": "edit_debit",
                            "party_name": name,
                            "ledger_page": ledger_page,
                            "no": bill_to_edit.get("no"),
                            "date": ebill_date,
                            "bi_ka": ebill_bi_ka,
                            "description": ebill_desc,
                            "size": ebill_size,
                            "model": ebill_model,
                            "pd": ebill_pd,
                            "bill": ebill_bill_no,
                            "qty": qty_str,
                            "taka": taka_str,
                            "total": total_str,
                            "remarks": ebill_remarks
                        }
                        if send_sheets_post(SHEETS_URL, payload):
                            st.success("বিলটি গুগল শিটে আপডেট করা হয়েছে!")
                            st.rerun()
                    else:
                        debit_list[bill_idx] = {
                            "no": bill_to_edit.get("no"), "date": ebill_date, "bi_ka": ebill_bi_ka,
                            "description": ebill_desc, "size": ebill_size, "model": ebill_model,
                            "pd": ebill_pd, "bill": ebill_bill_no, "qty": qty_str, "taka": taka_str,
                            "total": total_str, "remarks": ebill_remarks
                        }
                        left_data["debit_table"] = debit_list
                        save_local_page_data(pdf_page_num, "left", left_data)
                        st.success("বিলটি লোকাল ফাইলে আপডেট করা হয়েছে!")
                        st.rerun()
                        
                if delete_bill:
                    if SHEETS_URL:
                        payload = {
                            "action": "delete_debit",
                            "party_name": name,
                            "ledger_page": ledger_page,
                            "no": bill_to_edit.get("no")
                        }
                        if send_sheets_post(SHEETS_URL, payload):
                            st.success("বিলটি গুগল শিট থেকে ডিলিট করা হয়েছে!")
                            st.rerun()
                    else:
                        debit_list.pop(bill_idx)
                        for i, row in enumerate(debit_list):
                            row["no"] = str(i + 1)
                        left_data["debit_table"] = debit_list
                        save_local_page_data(pdf_page_num, "left", left_data)
                        st.success("বিলটি লোকাল ফাইল থেকে ডিলিট করা হয়েছে!")
                        st.rerun()

    # 3. Edit/Delete Payments
    elif corr_type == "৩. পেমেন্ট রেকর্ড সংশোধন/ডিলিট (Edit/Delete Payments)":
        st.subheader("📉 পেমেন্ট রেকর্ড সংশোধন বা ডিলিট করুন")
        if not credit_list:
            st.info("এই গ্রাহকের কোনো পেমেন্ট রেকর্ড নেই।")
        else:
            pay_options = [
                {
                    "label": f"ক্রঃ নং {row.get('no')} | তারিখ: {row.get('date')} | পরিমাণ: {row.get('amount')} ৳ | মন্তব্য: {row.get('remarks')}",
                    "row": row,
                    "index": idx
                }
                for idx, row in enumerate(credit_list)
            ]
            
            selected_pay_opt = st.selectbox(
                "সংশোধনের জন্য পেমেন্ট এন্ট্রি নির্বাচন করুন (Select Payment):",
                options=pay_options,
                format_func=lambda x: x["label"]
            )
            
            pay_to_edit = selected_pay_opt["row"]
            pay_idx = selected_pay_opt["index"]
            
            st.write("---")
            with st.form("edit_payment_form"):
                col_p1, col_p2 = st.columns(2)
                with col_p1:
                    epay_date = st.text_input("তারিখ (Date):", pay_to_edit.get("date", ""))
                    amt_val = clean_num(pay_to_edit.get("amount"))
                    epay_amount = st.number_input("টাকার পরিমাণ (Amount):", min_value=0.0, step=100.0, value=amt_val)
                with col_p2:
                    epay_remarks = st.text_input("মন্তব্য (Remarks):", pay_to_edit.get("remarks", ""))
                    
                col_btn1, col_btn2 = st.columns(2)
                with col_btn1:
                    save_pay = st.form_submit_button("পেমেন্ট আপডেট করুন (Save Payment)")
                with col_btn2:
                    delete_pay = st.form_submit_button("পেমেন্ট ডিলিট করুন (Delete Payment)", type="primary")
                    
                if save_pay:
                    amt_str = str(int(epay_amount)) if epay_amount.is_integer() else str(epay_amount)
                    
                    if SHEETS_URL:
                        payload = {
                            "action": "edit_credit",
                            "party_name": name,
                            "ledger_page": ledger_page + 1,
                            "no": pay_to_edit.get("no"),
                            "date": epay_date,
                            "amount": amt_str,
                            "remarks": epay_remarks
                        }
                        if send_sheets_post(SHEETS_URL, payload):
                            st.success("পেমেন্ট এন্ট্রি গুগল শিটে আপডেট করা হয়েছে!")
                            st.rerun()
                    else:
                        credit_list[pay_idx] = {
                            "no": pay_to_edit.get("no"), "date": epay_date, "amount": amt_str, "remarks": epay_remarks
                        }
                        right_data["credit_table"] = credit_list
                        save_local_page_data(pdf_page_num, "right", right_data)
                        st.success("পেমেন্ট এন্ট্রি লোকাল ফাইলে আপডেট করা হয়েছে!")
                        st.rerun()
                        
                if delete_pay:
                    if SHEETS_URL:
                        payload = {
                            "action": "delete_credit",
                            "party_name": name,
                            "ledger_page": ledger_page + 1,
                            "no": pay_to_edit.get("no")
                        }
                        if send_sheets_post(SHEETS_URL, payload):
                            st.success("পেমেন্ট এন্ট্রি গুগল শিট থেকে ডিলিট করা হয়েছে!")
                            st.rerun()
                    else:
                        credit_list.pop(pay_idx)
                        for i, row in enumerate(credit_list):
                            row["no"] = str(i + 1)
                        right_data["credit_table"] = credit_list
                        save_local_page_data(pdf_page_num, "right", right_data)
                        st.success("পেমেন্ট এন্ট্রি লোকাল ফাইল থেকে ডিলিট করা হয়েছে!")
                        st.rerun()
