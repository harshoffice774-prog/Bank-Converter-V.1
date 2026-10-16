import streamlit as st
import re
import pandas as pd
from openpyxl.styles import Border, Side, Alignment, Font
from datetime import datetime
import io

# --- HELPER FUNCTIONS ---
def extract_metadata(text):
    flat_text = text.replace('\n', '  ')
    meta = {
        "name": "", "address": "", "cust_id": "", "ifsc": "", 
        "acct_num": "", "branch": "", "date_from": "-", "date_to": "-"
    }
    
    name_m = re.search(r'Account holder name:\s*(.*?)\s+(?:Account holder address:|Customer ID:)', flat_text, re.IGNORECASE)
    if name_m: meta["name"] = name_m.group(1).strip()
    
    addr_m = re.search(r'Account holder address:\s*(.*?)\s+(?:Customer ID:|IFSC:|Account number:)', flat_text, re.IGNORECASE)
    if addr_m: meta["address"] = addr_m.group(1).strip()
    
    cust_m = re.search(r'Customer ID:\s*(\d+)', flat_text, re.IGNORECASE)
    if cust_m: meta["cust_id"] = cust_m.group(1).strip()
    
    ifsc_m = re.search(r'IFSC:\s*([A-Z0-9]+)', flat_text, re.IGNORECASE)
    if ifsc_m: meta["ifsc"] = ifsc_m.group(1).strip()
    
    acct_m = re.search(r'Account number:\s*(\d+)', flat_text, re.IGNORECASE)
    if acct_m: meta["acct_num"] = acct_m.group(1).strip()
    
    branch_m = re.search(r'Branch Name:\s*([A-Za-z]+)', flat_text, re.IGNORECASE)
    if branch_m: meta["branch"] = branch_m.group(1).strip()
    
    from_m = re.search(r'from:\s*([\d-]+)\s+to:\s*([\d-]+)', flat_text, re.IGNORECASE)
    if from_m: 
        meta["date_from"] = from_m.group(1).strip()
        meta["date_to"] = from_m.group(2).strip()
        
    return meta

def build_excel_bytes(full_text):
    lines = full_text.split('\n')
    meta = extract_metadata(full_text)

    transactions = []
    current_txn = None
    trigger_pattern = re.compile(r'^(\d+)\s+(\d{2}-\d{2}-\d{4})\s+(.+?)\s+([\d,]+\.\d{2})$')
    balance_pattern = re.compile(r'^₹\s*([\d,]+\.\d{2})')

    for line in lines:
        line = line.strip()
        if not line or line.startswith('---'): continue

        trigger_match = trigger_pattern.match(line)
        if trigger_match:
            if current_txn and 'Balance' in current_txn: transactions.append(current_txn)
            current_txn = {
                'Date': trigger_match.group(2), 'Remarks': trigger_match.group(3),
                'Amount': float(trigger_match.group(4).replace(',', '')), 'Balance': None
            }
            continue

        balance_match = balance_pattern.match(line)
        if balance_match and current_txn:
            current_txn['Balance'] = float(balance_match.group(1).replace(',', ''))
            continue

        if current_txn:
            lower_line = line.lower()
            if "discrepancy" in lower_line or "passwords" in lower_line or "atm" in lower_line or "page" in lower_line:
                continue
            current_txn['Remarks'] += " " + line

    if current_txn and 'Balance' in current_txn: transactions.append(current_txn)
    
    if not transactions:
        return None

    transactions.reverse()
    df = pd.DataFrame(transactions)

    debits, credits = [], []
    for i in range(len(df)):
        amt = df.loc[i, 'Amount']
        remarks = df.loc[i, 'Remarks'].upper()
        if i > 0:
            diff = round(df.loc[i, 'Balance'] - df.loc[i-1, 'Balance'], 2)
            if diff > 0: credits.append(amt); debits.append(None)
            elif diff < 0: debits.append(amt); credits.append(None)
            else:
                if '/CR/' in remarks or 'CREDIT' in remarks: credits.append(amt); debits.append(None)
                else: debits.append(amt); credits.append(None)
        else:
            if '/CR/' in remarks or 'CREDIT' in remarks: credits.append(amt); debits.append(None)
            else: debits.append(amt); credits.append(None)

    df['Debit'] = debits
    df['Credit'] = credits
    df = df[['Date', 'Remarks', 'Debit', 'Credit', 'Balance']]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Statement', index=False, startrow=13)
        worksheet = writer.sheets['Statement']

        worksheet.merge_cells('A1:E2')
        worksheet['A1'] = "Detailed Statement"
        worksheet['A1'].font = Font(size=18, bold=True)
        worksheet['A1'].alignment = Alignment(horizontal='center', vertical='center')
        
        worksheet['E3'] = f"Date: {datetime.today().strftime('%d/%m/%Y')}"
        worksheet['E3'].font = Font(bold=True)
        worksheet['E3'].alignment = Alignment(horizontal='right')

        worksheet['A4'] = "Account holder name:"
        worksheet['B4'] = meta["name"]
        worksheet['C4'] = "Account holder address:"
        worksheet.merge_cells('D4:E4')
        worksheet['D4'] = meta["address"]
        
        worksheet['A5'] = "Customer ID:"
        worksheet['B5'] = meta["cust_id"]
        worksheet['C5'] = "IFSC:"
        worksheet.merge_cells('D5:E5')
        worksheet['D5'] = meta["ifsc"]
        
        worksheet['A6'] = "Account number:"
        worksheet['B6'] = meta["acct_num"]
        worksheet['C6'] = "Branch Name:"
        worksheet.merge_cells('D6:E6')
        worksheet['D6'] = meta["branch"]
        
        box_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                            top=Side(style='thin'), bottom=Side(style='thin'))
        for r in range(4, 7):
            for c in range(1, 6):
                cell = worksheet.cell(row=r, column=c)
                cell.border = box_border
                cell.alignment = Alignment(vertical='center', wrap_text=True)
                if c in [1, 3]: cell.font = Font(size=10)
                else: cell.font = Font(size=10, bold=True)

        worksheet['A8'] = "Transaction Date"
        worksheet['B8'] = f"from: {meta['date_from']}"
        worksheet['C8'] = f"to: {meta['date_to']}"
        worksheet['A9'] = "Amount"
        worksheet['B9'] = "from: -"
        worksheet['C9'] = "to: -"
        worksheet['A10'] = "Cheque"
        worksheet['B10'] = "from: -"
        worksheet['C10'] = "to: -"
        worksheet['A11'] = "Transaction type: All"
        
        for r in range(8, 12):
            for c in range(1, 4):
                worksheet.cell(row=r, column=c).font = Font(bold=True, size=10)

        headers = ["DATE", "REMARKS", "DEBIT", "CREDIT", "BALANCE"]
        for col_num, header in enumerate(headers, 1):
            cell = worksheet.cell(row=14, column=col_num, value=header)
            cell.font = Font(bold=True)
            cell.border = box_border
            cell.alignment = Alignment(horizontal='center', vertical='center')

        max_row = worksheet.max_row
        for row in worksheet.iter_rows(min_row=15, max_row=max_row, min_col=1, max_col=5):
            for cell in row:
                cell.border = box_border
                cell.alignment = Alignment(wrap_text=True, vertical='center')

        for r in range(16, max_row + 1):
            worksheet[f"E{r}"] = f"=E{r-1}+SUM(D{r})-SUM(C{r})"

        worksheet.column_dimensions['A'].width = 18
        worksheet.column_dimensions['B'].width = 40
        worksheet.column_dimensions['C'].width = 15
        worksheet.column_dimensions['D'].width = 15
        worksheet.column_dimensions['E'].width = 18
        worksheet.row_dimensions[4].height = 30 

    return output.getvalue()


# --- STREAMLIT UI & LOGIN SYSTEM ---
st.set_page_config(page_title="Statement Converter")

# Initialize session state for authentication
if 'authenticated' not in st.session_state:
    st.session_state['authenticated'] = False

if not st.session_state['authenticated']:
    # Show Login Form
    st.title("Login Required")
    
    user_id = st.text_input("User ID")
    password = st.text_input("Password", type="password")
    
    if st.button("Login"):
        if user_id == "Harsh" and password == "9806":
            st.session_state['authenticated'] = True
            st.rerun()  # Refreshes the page to show the main app
        else:
            st.error("Incorrect User ID or Password.")

else:
    # Show Main App
    st.title("Bank Statement Text to Excel Converter")
    
    # Optional logout button in the sidebar
    if st.sidebar.button("Logout"):
        st.session_state['authenticated'] = False
        st.rerun()

    uploaded_file = st.file_uploader("Upload raw statement text file (.txt)", type=["txt"])

    if uploaded_file is not None:
        raw_text = uploaded_file.read().decode("utf-8")
        
        if st.button("Process Statement"):
            with st.spinner("Converting..."):
                excel_data = build_excel_bytes(raw_text)
                
                if excel_data:
                    st.success("Conversion successful!")
                    st.download_button(
                        label="Download Final Statement (Excel)",
                        data=excel_data,
                        file_name="final_statement.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("No valid transactions found in the uploaded text.")