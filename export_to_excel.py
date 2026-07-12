import os
import sys
import json
import pandas as pd
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Paths
WORKSPACE_DIR = r"e:\user\OneDrive - Bangladesh Telecommunication Regulatory Commission\ABM\Sunnah\AL karim hisab"
DATA_DIR = os.path.join(WORKSPACE_DIR, "data")
JSON_DIR = os.path.join(DATA_DIR, "json")
OUTPUT_EXCEL = os.path.join(WORKSPACE_DIR, "hisab_alkarim_digitized.xlsx")

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

def load_data():
    # Load index data
    index_file = os.path.join(JSON_DIR, "index_data.json")
    index_map = {}
    index_list = []
    if os.path.exists(index_file):
        with open(index_file, "r", encoding="utf-8") as f:
            index_list = json.load(f)
            for entry in index_list:
                try:
                    p_num = int(entry.get("page", 0))
                    if p_num:
                        if p_num not in index_map:
                            index_map[p_num] = []
                        index_map[p_num].append(entry)
                except Exception:
                    pass

    debit_rows = []
    credit_rows = []
    
    # We will also keep track of debit/credit sums per ledger page
    page_sums = {}

    print("Compiling JSON files...")
    for page_num in range(3, 111):
        left_ledger = 2 * page_num
        right_ledger = 2 * page_num + 1
        
        # 1. Left Page (Debit)
        left_file = os.path.join(JSON_DIR, f"page_{page_num}_left.json")
        if os.path.exists(left_file):
            with open(left_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            
            transcribed_name = str(data.get("party_name") or "").strip()
            phone = str(data.get("phone") or "").strip()
            ledger_page = str(data.get("ledger_page_number") or left_ledger).strip()
            
            # Find expected index info
            idx_infos = index_map.get(left_ledger, [])
            idx_names = ", ".join([x.get("party_name", "") for x in idx_infos])
            idx_address = ", ".join([x.get("address", "") for x in idx_infos])
            idx_notes = "; ".join([x.get("notes", "") for x in idx_infos if x.get("notes", "")])
            
            table = data.get("debit_table", [])
            for row in table:
                qty = clean_num(row.get("qty"))
                taka = clean_num(row.get("taka"))
                total = clean_num(row.get("total"))
                
                # If total is 0 but qty and rate exist, compute it
                if total == 0.0 and qty > 0 and taka > 0:
                    total = qty * taka
                    
                row_data = {
                    "PDF Page": page_num,
                    "Ledger Page": ledger_page,
                    "Transcribed Name": transcribed_name,
                    "Expected Client (Index)": idx_names if idx_names else None,
                    "Address (Index)": idx_address if idx_address else None,
                    "Notes (Index)": idx_notes if idx_notes else None,
                    "No": row.get("no", ""),
                    "Date": row.get("date", ""),
                    "Details (বিঃ কাঃ)": row.get("bi_ka", ""),
                    "Description": row.get("description", ""),
                    "Size": row.get("size", ""),
                    "Model": row.get("model", ""),
                    "PD": row.get("pd", ""),
                    "Bill No": row.get("bill", ""),
                    "Qty": qty if qty > 0 else row.get("qty", ""),
                    "Rate (Taka)": taka if taka > 0 else row.get("taka", ""),
                    "Total": total if total > 0 else row.get("total", ""),
                    "Remarks": row.get("remarks", "")
                }
                debit_rows.append(row_data)
                
                # Add to ledger page sum
                if left_ledger not in page_sums:
                    page_sums[left_ledger] = {"debit": 0.0, "credit": 0.0}
                page_sums[left_ledger]["debit"] += total

        # 2. Right Page (Credit)
        right_file = os.path.join(JSON_DIR, f"page_{page_num}_right.json")
        if os.path.exists(right_file):
            with open(right_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                
            transcribed_name = str(data.get("party_name") or "").strip()
            phone = str(data.get("phone") or "").strip()
            ledger_page = str(data.get("ledger_page_number") or right_ledger).strip()
            
            # Find expected index info
            idx_infos = index_map.get(left_ledger, []) # Often credit page is indexed by the main left page
            if not idx_infos:
                idx_infos = index_map.get(right_ledger, [])
                
            idx_names = ", ".join([x.get("party_name", "") for x in idx_infos])
            idx_address = ", ".join([x.get("address", "") for x in idx_infos])
            idx_notes = "; ".join([x.get("notes", "") for x in idx_infos if x.get("notes", "")])
            
            table = data.get("credit_table", [])
            for row in table:
                amount = clean_num(row.get("amount"))
                row_data = {
                    "PDF Page": page_num,
                    "Ledger Page": ledger_page,
                    "Transcribed Name": transcribed_name,
                    "Expected Client (Index)": idx_names if idx_names else None,
                    "Address (Index)": idx_address if idx_address else None,
                    "Notes (Index)": idx_notes if idx_notes else None,
                    "No": row.get("no", ""),
                    "Date": row.get("date", ""),
                    "Amount": amount if amount > 0 else row.get("amount", ""),
                    "Remarks": row.get("remarks", "")
                }
                credit_rows.append(row_data)
                
                # Add to ledger page sum (attribute to left ledger page since they are the same client)
                client_page = left_ledger
                if client_page not in page_sums:
                    page_sums[client_page] = {"debit": 0.0, "credit": 0.0}
                page_sums[client_page]["credit"] += amount

    # 3. Create Summary list based on Index entries
    summary_rows = []
    for idx_entry in index_list:
        p_num = int(idx_entry.get("page", 0))
        debit_sum = 0.0
        credit_sum = 0.0
        if p_num in page_sums:
            debit_sum = page_sums[p_num]["debit"]
            credit_sum = page_sums[p_num]["credit"]
            
        summary_rows.append({
            "Client No": idx_entry.get("no", ""),
            "Party Name": idx_entry.get("party_name", ""),
            "Address": idx_entry.get("address", ""),
            "Ledger Page": p_num,
            "PDF Page": p_num // 2 if p_num > 0 else "",
            "Total Debit (Sales)": debit_sum,
            "Total Credit (Payments)": credit_sum,
            "Outstanding Balance": debit_sum - credit_sum,
            "Notes (Index)": idx_entry.get("notes", "")
        })

    return summary_rows, debit_rows, credit_rows

def export():
    summary_rows, debit_rows, credit_rows = load_data()
    
    df_summary = pd.DataFrame(summary_rows)
    df_debit = pd.DataFrame(debit_rows)
    df_credit = pd.DataFrame(credit_rows)
    
    print(f"Exporting to {OUTPUT_EXCEL}...")
    
    with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:
        df_summary.to_excel(writer, sheet_name="Client Summary", index=False)
        df_debit.to_excel(writer, sheet_name="Debit Entries (Sales)", index=False)
        df_credit.to_excel(writer, sheet_name="Credit Entries (Payments)", index=False)
        
        # Access sheets for styling
        workbook = writer.book
        
        # Styles
        font_family = "Segoe UI"
        header_font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        title_font = Font(name=font_family, size=11, bold=False)
        border_side = Side(border_style="thin", color="D3D3D3")
        cell_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)
        
        for name in workbook.sheetnames:
            sheet = workbook[name]
            
            # Format header row
            for col_idx in range(1, sheet.max_column + 1):
                cell = sheet.cell(row=1, column=col_idx)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            
            # Auto-fit column widths
            for col in sheet.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    cell.font = Font(name=font_family, size=10)
                    cell.border = cell_border
                    
                    # Formatting values
                    if cell.row > 1:
                        # Right-align numeric columns
                        if isinstance(cell.value, (int, float)):
                            cell.alignment = Alignment(horizontal="right")
                            # Format as currency/number
                            if name == "Client Summary" and cell.column in [6, 7, 8]:
                                cell.number_format = "#,##0.00"
                            elif name == "Debit Entries (Sales)" and cell.column in [15, 16, 17]:
                                cell.number_format = "#,##0.00"
                            elif name == "Credit Entries (Payments)" and cell.column in [9]:
                                cell.number_format = "#,##0.00"
                        else:
                            cell.alignment = Alignment(horizontal="left", vertical="center")
                            
                    val_str = str(cell.value or "")
                    if len(val_str) > max_len:
                        max_len = len(val_str)
                
                sheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
            
            # Set header row height
            sheet.row_dimensions[1].height = 28
            
    print("Excel file created successfully!")

if __name__ == "__main__":
    export()
