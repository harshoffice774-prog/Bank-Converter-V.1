import re
import pandas as pd
from openpyxl.styles import Border, Side, Alignment, Font
from datetime import datetime
import os
import streamlit as st

def extract_metadata(text):
    # Flattens the text to make finding the account details easier
    flat_text = text.replace('\n', '  ')
    meta = {
        "name": "", "address": "", "cust_id": "", "ifsc": "", 
        "acct_num": "", "branch": "", "date_from": "-", "date_to": "-"
    }
    
    # Scans for the specific labels and captures the data next to them
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

def build_excel_from_text(text_path, excel_path):
    print(f"Reading raw data from {text_path}...")
    with open(text_path, 'r', encoding='utf-8') as f:
        full_text = f.read()
        lines = full_text.split('\n')

    # 1. Grab the Account Details for the Header
    meta = extract_metadata(full_text)

    # 2. Process Transactions
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
        print("No transactions found!")
        return

    # Flip to Chronological Order
    transactions.reverse()
    df = pd.DataFrame(transactions)

    # Math Validation (Debit vs Credit)
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

    # 3. WRITE THE BEAUTIFUL EXCEL FILE
    print("Applying Professional Formatting...")
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        # Push the table down to row 14 so we have room for the header
        df.to_excel(writer, sheet_name='Statement', index=False, startrow=13)
        worksheet = writer.sheets['Statement']

        # --- TITLE & DATE ---
        worksheet.merge_cells('A1:E2')
        worksheet['A1'] = "Detailed Statement"
        worksheet['A1'].font = Font(size=18, bold=True)
        worksheet['A1'].alignment = Alignment(horizontal='center', vertical='center')
        
        worksheet['E3'] = f"Date: {datetime.today().strftime('%d/%m/%Y')}"
        worksheet['E3'].font = Font(bold=True)
        worksheet['E3'].alignment = Alignment(horizontal='right')

        # --- THE INFO BOX ---
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
        
        # Apply Box Borders
        box_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                            top=Side(style='thin'), bottom=Side(style='thin'))
        for r in range(4, 7):
            for c in range(1, 6):
                cell = worksheet.cell(row=r, column=c)
                cell.border = box_border
                cell.alignment = Alignment(vertical='center', wrap_text=True)
                if c in [1, 3]: cell.font = Font(size=10) # Normal font for labels
                else: cell.font = Font(size=10, bold=True) # Bold font for your details

        # --- FILTERS ---
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

        # --- FORMAT TRANSACTION TABLE ---
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

        # Inject Running Balance Formula
        for r in range(16, max_row + 1):
            worksheet[f"E{r}"] = f"=E{r-1}+SUM(D{r})-SUM(C{r})"

        # Adjust Column Widths & Heights to fit everything beautifully
        worksheet.column_dimensions['A'].width = 18
        worksheet.column_dimensions['B'].width = 40
        worksheet.column_dimensions['C'].width = 15
        worksheet.column_dimensions['D'].width = 15
        worksheet.column_dimensions['E'].width = 18
        worksheet.row_dimensions[4].height = 30 # Gives extra room for address to wrap

    print(f"Success! Final professional statement saved to: {os.path.abspath(excel_path)}")


# --- STREAMLIT ONLINE INTERFACE & LOGIN ---
if __name__ == "__main__":
    st.set_page_config(page_title="Statement Converter")

    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False

    # 1. Login Screen
    if not st.session_state.logged_in:
        st.title("Login to Statement Converter")
        user_id = st.text_input("USER ID")
        password = st.text_input("Password", type="password")

        if st.button("Login"):
            if user_id == "Harsh" and password == "9806":
                st.session_state.logged_in = True
                st.rerun()
            else:
                st.error("Invalid USER ID or Password")

    # 2. Main Converter Screen (After Login)
    else:
        st.title("Bank Statement Converter")
        
        if st.sidebar.button("Logout"):
            st.session_state.logged_in = False
            st.rerun()

        uploaded_file = st.file_uploader("Upload your raw_data.txt file", type=["txt"])

        if uploaded_file is not None:
            if st.button("Convert to Excel"):
                # Remove old output file if it exists
                if os.path.exists("final_statement.xlsx"):
                    os.remove("final_statement.xlsx")

                # Save uploaded file as raw_data.txt so your original function can read it
                with open("raw_data.txt", "wb") as f:
                    f.write(uploaded_file.getbuffer())

                # Run your exact function
                build_excel_from_text("raw_data.txt", "final_statement.xlsx")

                # Check if the Excel file was created and offer download
                if os.path.exists("final_statement.xlsx"):
                    st.success("Success! Final professional statement is ready.")
                    with open("final_statement.xlsx", "rb") as excel_file:
                        st.download_button(
                            label="Download final_statement.xlsx",
                            data=excel_file,
                            file_name="final_statement.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.error("No transactions found!")
