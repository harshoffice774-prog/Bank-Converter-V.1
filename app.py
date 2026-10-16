import re
import pandas as pd
from openpyxl.styles import Border, Side, Alignment, Font
from datetime import datetime
import os
import pdfplumber
from pdfplumber.utils.exceptions import PdfminerException
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

def build_excel_from_pdf(pdf_path, excel_path, pdf_password=""):
    print(f"Reading PDF data from {pdf_path}...")
    full_text = ""
    
    with pdfplumber.open(pdf_path, password=pdf_password) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                full_text += page_text + "\n"

    # Ensure ₹ Balance is always on its own line if the PDF extractor merges it onto the same line
    full_text = re.sub(r'(?<!^)\s*(₹\s*[\d,]+\.\d{2})', r'\n\1', full_text, flags=re.MULTILINE)
    lines = full_text.split('\n')

    # 1. Grab the Account Details for the Header
    meta = extract_metadata(full_text)

    # 2. Process Transactions
    transactions = []
    current_txn = None
    trigger_pattern = re.compile(r'^
