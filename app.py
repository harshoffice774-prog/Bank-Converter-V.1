import streamlit as st
import io
import re
import time
import pandas as pd
import pdfplumber
import fitz  # PyMuPDF - renders PDF pages to images (no Poppler needed)
import pytesseract
from PIL import Image, ImageOps
from pdfminer.pdfdocument import PDFPasswordIncorrect
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment, Font, PatternFill
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.properties import PageSetupProperties

# Windows only: if Tesseract is not in PATH, uncomment and fix this path
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


# --- HELPER: TIME FORMATTER ---
def format_seconds(sec):
    if sec < 60:
        return f"{sec:.1f} sec"
    mins = int(sec // 60)
    rem_sec = sec % 60
    return f"{mins}m {rem_sec:.1f}s"


# --- HELPER: UNIVERSAL DATE PARSER ---
def parse_any_date(date_str):
    clean_d = re.sub(r'\s+', '-', date_str.strip()).replace('/', '-')
    for fmt in ('%d-%m-%Y', '%d-%m-%y', '%d-%b-%Y', '%d-%b-%y', '%d-%B-%Y', '%Y-%m-%d'):
        try:
            return datetime.strptime(clean_d, fmt)
        except ValueError:
            continue
    return None


# --- OCR HELPERS ---
def tesseract_available():
    try:
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


def render_page_image(doc, page_idx, dpi):
    """Draws one PDF page as an image in memory, then cleans it for OCR."""
    page = doc[page_idx]
    zoom = dpi / 72
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    img = ImageOps.grayscale(img)
    img = ImageOps.autocontrast(img, cutoff=2)
    # img = img.point(lambda p: 255 if p > 170 else 0)
    img = ImageOps.expand(img, border=20, fill=255)
    return img


def clean_ocr_text(text):
    """Fix common OCR glitches so the amount/date patterns still match."""
    text = text.replace('|', ' ').replace('_', ' ')
    text = re.sub(r'\s[=~]\s', ' ', text)                  # stray '=' between columns
    text = re.sub(r'(\d)\s+\.(\d{2})\b', r'\1.\2', text)    # "1,234 .56" -> "1,234.56"
    text = re.sub(r'(\d),\s(\d{3})', r'\1,\2', text)         # "1, 234.00" -> "1,234.00"
    text = re.sub(r'[ \t]+', ' ', text)
    return text


def ocr_page_text(img, lang):
    """Tries several page-segmentation modes and falls back to English."""
    last_err = None
    langs = [lang] if lang == "eng" else [lang, "eng"]

    for lg in langs:
        for psm in (6, 4, 3, 11):
            try:
                raw = pytesseract.image_to_string(
                    img, lang=lg, config=f"--oem 3 --psm {psm}"
                )
                if raw and len(raw.strip()) > 20:
                    return clean_ocr_text(raw)
            except Exception as e:
                last_err = e
                continue

    if last_err:
        raise ValueError(f"OCR failed on this page: {last_err}")
    return ""


# --- UNIVERSAL METADATA EXTRACTOR ---
KNOWN_BANKS = [
    ("State Bank of India", r"State\s+Bank\s+of\s+India|\bSBI\b"),
    ("HDFC Bank", r"HDFC\s+Bank"),
    ("ICICI Bank", r"ICICI\s+Bank"),
    ("Axis Bank", r"Axis\s+Bank"),
    ("Kotak Mahindra Bank", r"Kotak\s+Mahindra"),
    ("Punjab National Bank", r"Punjab\s+National\s+Bank|\bPNB\b"),
    ("Bank of Baroda", r"Bank\s+of\s+Baroda"),
    ("Canara Bank", r"Canara\s+Bank"),
    ("Union Bank of India", r"Union\s+Bank\s+of\s+India"),
    ("Central Bank of India", r"Central\s+Bank\s+of\s+India"),
    ("Bank of Maharashtra", r"Bank\s+of\s+Maharashtra"),
    ("Bank of India", r"Bank\s+of\s+India"),
    ("IDBI Bank", r"IDBI\s+Bank"),
    ("Yes Bank", r"Yes\s+Bank"),
    ("IndusInd Bank", r"IndusInd\s+Bank"),
    ("Federal Bank", r"Federal\s+Bank"),
    ("IDFC FIRST Bank", r"IDFC\s+FIRST\s+Bank"),
    ("Indian Overseas Bank", r"Indian\s+Overseas\s+Bank"),
    ("Indian Bank", r"Indian\s+Bank"),
    ("UCO Bank", r"UCO\s+Bank"),
    ("Bandhan Bank", r"Bandhan\s+Bank"),
    ("Punjab & Sind Bank", r"Punjab\s*(?:&|and)\s*Sind\s+Bank"),
]

IFSC_BANK_PREFIX = {
    "SBIN": "State Bank of India", "HDFC": "HDFC Bank", "ICIC": "ICICI Bank",
    "UTIB": "Axis Bank", "KKBK": "Kotak Mahindra Bank", "PUNB": "Punjab National Bank",
    "BARB": "Bank of Baroda", "CNRB": "Canara Bank", "UBIN": "Union Bank of India",
    "BKID": "Bank of India", "IBKL": "IDBI Bank", "YESB": "Yes Bank",
    "INDB": "IndusInd Bank", "FDRL": "Federal Bank", "IDFB": "IDFC FIRST Bank",
    "IDIB": "Indian Bank", "CBIN": "Central Bank of India", "UCBA": "UCO Bank",
    "BDBL": "Bandhan Bank", "MAHB": "Bank of Maharashtra", "IOBA": "Indian Overseas Bank",
    "PSIB": "Punjab & Sind Bank",
}


def extract_metadata(text):
    flat_text = text.replace('\n', '  ')
    meta = {
        "bank": "-", "name": "-", "address": "-", "cust_id": "-", "ifsc": "-",
        "acct_num": "-", "acct_type": "-", "branch": "-", "email": "-",
        "date_from": "-", "date_to": "-"
    }

    name_m = re.search(r'(?:Account\s*holder\s*name|Customer\s*Name|Account\s*Name|Name)\s*[:\-]\s*(.*?)\s+(?:Account|Customer|Address|Branch|IFSC|CIF)', flat_text, re.IGNORECASE)
    if name_m: meta["name"] = name_m.group(1).strip()

    addr_m = re.search(r'(?:Account\s*holder\s*address|Customer\s*Address|Address)\s*[:\-]\s*(.*?)\s+(?:Customer|IFSC|Account|Branch|Email|Mobile|Phone)', flat_text, re.IGNORECASE)
    if addr_m: meta["address"] = addr_m.group(1).strip()

    cust_m = re.search(r'(?:Customer\s*ID|Cust\s*ID|CIF\s*No|CIF\s*ID)\s*[:\-]?\s*([A-Za-z0-9]+)', flat_text, re.IGNORECASE)
    if cust_m: meta["cust_id"] = cust_m.group(1).strip()

    ifsc_m = re.search(r'(?:IFSC\s*Code|RTGS\/NEFT\s*IFSC|IFSC)\s*[:\-]?\s*([A-Z]{4}0[A-Z0-9]{6})', flat_text, re.IGNORECASE)
    if ifsc_m: meta["ifsc"] = ifsc_m.group(1).strip().upper()

    acct_m = re.search(r'(?:Account\s*number|Account\s*No|A\/c\s*No|Ac\s*No)\s*[:\-]?\s*(\d{9,18})', flat_text, re.IGNORECASE)
    if acct_m: meta["acct_num"] = acct_m.group(1).strip()

    branch_m = re.search(r'(?:Branch\s*Name|Home\s*Branch|Branch)\s*[:\-]\s*([A-Za-z0-9\s,\-]+?)(?:\s{2,}|IFSC|Account|Customer|Date|Period)', flat_text, re.IGNORECASE)
    if branch_m: meta["branch"] = branch_m.group(1).strip()

    from_m = re.search(r'(?:from|Period)\s*[:\-]?\s*(\d{1,2}[-/\s][A-Za-z0-9]{2,3}[-/\s]\d{2,4})\s+to\s*[:\-]?\s*(\d{1,2}[-/\s][A-Za-z0-9]{2,3}[-/\s]\d{2,4})', flat_text, re.IGNORECASE)
    if from_m:
        meta["date_from"] = from_m.group(1).strip()
        meta["date_to"] = from_m.group(2).strip()

    # Account type: "Account Type: Savings" or "Savings Account"
    type_m = re.search(r'(?:Account\s*Type|A\/c\s*Type|Type\s*of\s*Account|Scheme\s*Name)\s*[:\-]\s*([A-Za-z ]{3,40}?)(?=\s{2,}|\s+(?:Account|Customer|IFSC|Branch|Date|Currency|MICR|Name|Address)\b|$)', flat_text, re.IGNORECASE)
    if type_m:
        meta["acct_type"] = type_m.group(1).strip()
    else:
        kw_m = re.search(r'\b(Savings|Current|Salary|Cash\s*Credit|Overdraft)\s+(?:Bank\s+)?(?:Account|A\/c)\b', flat_text, re.IGNORECASE)
        if kw_m: meta["acct_type"] = kw_m.group(1).strip().title() + " Account"

    email_m = re.search(r'[\w.\-+]+@[\w\-]+\.[\w.\-]+', flat_text)
    if email_m: meta["email"] = email_m.group(0)

    # Bank name: known names in the top of the statement, then IFSC prefix as backup
    head = flat_text[:3000]
    for display, pattern in KNOWN_BANKS:
        if re.search(pattern, head, re.IGNORECASE):
            meta["bank"] = display
            break
    if meta["bank"] == "-" and meta["ifsc"] != "-":
        meta["bank"] = IFSC_BANK_PREFIX.get(meta["ifsc"][:4], "-")

    return meta


# --- PROFESSIONAL EXCEL BUILDER ---
NAVY, BLUE, LIGHT, GREY = "1F3864", "2F5597", "D9E1F2", "F2F2F2"
_THIN = Side(style="thin", color="BFBFBF")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_NUM = '#,##0.00'


def _style(cell, bold=False, size=10, color="000000", fill=None, h=None, v="center",
           wrap=False, fmt=None, border=True, italic=False):
    cell.font = Font(name="Calibri", size=size, bold=bold, color=color, italic=italic)
    if fill:
        cell.fill = PatternFill("solid", start_color=fill, end_color=fill)
    cell.alignment = Alignment(horizontal=h, vertical=v, wrap_text=wrap)
    if fmt:
        cell.number_format = fmt
    if border:
        cell.border = _BORDER


def _merge(ws, rng, value=None, **kw):
    ws.merge_cells(rng)
    rows = ws[rng]
    rows[0][0].value = value
    for row in rows:
        for c in row:
            _style(c, **kw)


def build_excel(df, meta, opening_value, opening_derived, ocr_used):
    n = len(df)
    period_from = meta["date_from"] if meta["date_from"] != "-" else str(df["Date"].iloc[0])
    period_to = meta["date_to"] if meta["date_to"] != "-" else str(df["Date"].iloc[-1])
    period = f"{period_from}  to  {period_to}"

    wb = Workbook()
    ws_sum = wb.active
    ws_sum.title = "Summary"
    ws = wb.create_sheet("Transactions")
    ws_sum.sheet_view.showGridLines = False
    ws.sheet_view.showGridLines = False

    HEADER_ROW, OPEN_ROW, FIRST = 4, 5, 6
    LAST = OPEN_ROW + n
    TOTAL = LAST + 1

    # ---------------- Transactions sheet ----------------
    _merge(ws, "A1:G1", f"Transaction Details  -  {meta['bank'] if meta['bank'] != '-' else 'Bank Statement'}",
           bold=True, size=16, color="FFFFFF", fill=NAVY, h="center", border=False)
    ws.row_dimensions[1].height = 30
    _merge(ws, "A2:G2",
           f"Account Holder: {meta['name']}   |   Account No: {meta['acct_num']}   |   Period: {period}",
           size=10, color="FFFFFF", fill=BLUE, h="center", border=False)
    ws.row_dimensions[2].height = 20

    headers = ["Date", "Remarks / Narration", "Debit (Dr)", "Credit (Cr)",
               "Balance (Calculated)", "Balance (Per Statement)", "Check"]
    for col, text in enumerate(headers, 1):
        c = ws.cell(row=HEADER_ROW, column=col, value=text)
        _style(c, bold=True, color="FFFFFF", fill=NAVY, h="center", wrap=True)
    ws.row_dimensions[HEADER_ROW].height = 32

    # Opening balance row
    for col in range(1, 8):
        _style(ws.cell(row=OPEN_ROW, column=col), bold=True, fill=LIGHT)
    ws.cell(row=OPEN_ROW, column=2, value="Opening Balance")
    ws.cell(row=OPEN_ROW, column=5, value=opening_value)
    ws.cell(row=OPEN_ROW, column=6, value=f"=E{OPEN_ROW}")
    for col in (5, 6):
        ws.cell(row=OPEN_ROW, column=col).number_format = _NUM
        ws.cell(row=OPEN_ROW, column=col).alignment = Alignment(horizontal="right", vertical="center")

    # Transaction rows
    for i in range(n):
        r = FIRST + i
        row = df.iloc[i]
        band = GREY if i % 2 == 1 else None

        date_val = row["Date"]
        dt = parse_any_date(str(date_val))
        c = ws.cell(row=r, column=1, value=dt if dt else date_val)
        _style(c, fill=band, h="center", fmt="dd-mmm-yyyy" if dt else None)

        c = ws.cell(row=r, column=2, value=row["Remarks"])
        _style(c, fill=band, wrap=True, v="top")

        debit = row["Debit"] if pd.notna(row["Debit"]) else None
        credit = row["Credit"] if pd.notna(row["Credit"]) else None
        c = ws.cell(row=r, column=3, value=debit)
        _style(c, fill=band, h="right", fmt=_NUM, color="C00000")
        c = ws.cell(row=r, column=4, value=credit)
        _style(c, fill=band, h="right", fmt=_NUM, color="00703C")

        # Opening + Credit - Debit = Closing, row by row
        c = ws.cell(row=r, column=5, value=f"=ROUND(E{r-1}+D{r}-C{r},2)")
        _style(c, fill=band, h="right", fmt=_NUM, bold=True)
        c = ws.cell(row=r, column=6, value=float(row["Balance"]))
        _style(c, fill=band, h="right", fmt=_NUM)
        c = ws.cell(row=r, column=7, value=f'=IF(ROUND(F{r-1}+D{r}-C{r}-F{r},2)=0,"OK","CHECK")')
        _style(c, fill=band, h="center", bold=True)

    # Totals row
    for col in range(1, 8):
        _style(ws.cell(row=TOTAL, column=col), bold=True, fill=LIGHT)
    ws.cell(row=TOTAL, column=2, value="TOTAL")
    ws.cell(row=TOTAL, column=3, value=f"=SUM(C{FIRST}:C{LAST})")
    ws.cell(row=TOTAL, column=4, value=f"=SUM(D{FIRST}:D{LAST})")
    ws.cell(row=TOTAL, column=5, value=f"=E{LAST}")
    ws.cell(row=TOTAL, column=6, value=f"=F{LAST}")
    for col in (3, 4, 5, 6):
        cell = ws.cell(row=TOTAL, column=col)
        cell.number_format = _NUM
        cell.alignment = Alignment(horizontal="right", vertical="center")

    # Highlight OK / CHECK
    g_range = f"G{FIRST}:G{LAST}"
    ws.conditional_formatting.add(g_range, CellIsRule(
        operator="equal", formula=['"OK"'],
        font=Font(color="006100", bold=True), fill=PatternFill("solid", start_color="C6EFCE", end_color="C6EFCE")))
    ws.conditional_formatting.add(g_range, CellIsRule(
        operator="equal", formula=['"CHECK"'],
        font=Font(color="9C0006", bold=True), fill=PatternFill("solid", start_color="FFC7CE", end_color="FFC7CE")))

    for col, width in zip("ABCDEFG", (14, 52, 16, 16, 20, 22, 10)):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = f"A{OPEN_ROW}"
    ws.auto_filter.ref = f"A{HEADER_ROW}:G{LAST}"
    ws.print_title_rows = f"{HEADER_ROW}:{HEADER_ROW}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

    # ---------------- Summary sheet ----------------
    for col, width in zip("ABCD", (26, 44, 28, 26)):
        ws_sum.column_dimensions[col].width = width

    _merge(ws_sum, "A1:D1", "BANK ACCOUNT STATEMENT", bold=True, size=20, color="FFFFFF",
           fill=NAVY, h="center", border=False)
    ws_sum.row_dimensions[1].height = 38
    _merge(ws_sum, "A2:D2", meta["bank"] if meta["bank"] != "-" else "Bank name not detected",
           bold=True, size=13, color="FFFFFF", fill=BLUE, h="center", border=False)
    ws_sum.row_dimensions[2].height = 24

    _merge(ws_sum, "A4:D4", "ACCOUNT HOLDER DETAILS", bold=True, size=11, color=NAVY, fill=LIGHT, border=True)

    def pair(row, l_label, l_value, r_label, r_value):
        c = ws_sum.cell(row=row, column=1, value=l_label); _style(c, bold=True, fill=GREY)
        c = ws_sum.cell(row=row, column=2, value=l_value); _style(c, wrap=True)
        c = ws_sum.cell(row=row, column=3, value=r_label); _style(c, bold=True, fill=GREY)
        c = ws_sum.cell(row=row, column=4, value=r_value); _style(c, wrap=True)

    pair(5, "Account Holder Name", meta["name"], "Account Number", meta["acct_num"])
    pair(6, "Address", meta["address"], "Account Type", meta["acct_type"])
    pair(7, "Customer ID", meta["cust_id"], "IFSC Code", meta["ifsc"])
    pair(8, "Bank Name", meta["bank"], "Branch", meta["branch"])
    pair(9, "Email", meta["email"], "Report Generated On", datetime.today().strftime("%d-%b-%Y"))
    ws_sum.row_dimensions[6].height = 48
    c = ws_sum.cell(row=10, column=1, value="Statement Period"); _style(c, bold=True, fill=GREY)
    _merge(ws_sum, "B10:D10", period, bold=True)

    _merge(ws_sum, "A12:D12", "ACCOUNT SUMMARY", bold=True, size=11, color=NAVY, fill=LIGHT, border=True)

    def money_pair(row, l_label, l_formula, r_label, r_formula, fmt=_NUM):
        c = ws_sum.cell(row=row, column=1, value=l_label); _style(c, bold=True, fill=GREY)
        c = ws_sum.cell(row=row, column=2, value=l_formula); _style(c, h="right", bold=True, fmt=fmt)
        c = ws_sum.cell(row=row, column=3, value=r_label); _style(c, bold=True, fill=GREY)
        c = ws_sum.cell(row=row, column=4, value=r_formula); _style(c, h="right", bold=True, fmt=fmt)

    money_pair(13, "Opening Balance", f"=Transactions!E{OPEN_ROW}",
               "Total Credits (+)", f"=Transactions!D{TOTAL}")
    money_pair(14, "No. of Transactions", f"=COUNT(Transactions!F{FIRST}:F{LAST})",
               "Total Debits (-)", f"=Transactions!C{TOTAL}")
    ws_sum["B14"].number_format = "0"
    money_pair(15, "Closing Balance (Calculated)", "=B13+D13-D14",
               "Closing Balance (Per Statement)", f"=Transactions!F{LAST}")

    check_cnt = f'COUNTIF(Transactions!G{FIRST}:G{LAST},"CHECK")'
    c = ws_sum.cell(row=16, column=1, value="Reconciliation"); _style(c, bold=True, fill=GREY)
    _merge(ws_sum, "B16:D16",
           f'=IF(AND({check_cnt}=0,ROUND(B15-D15,2)=0),"All balances verified",'
           f'{check_cnt}&" row(s) need review - compare with the original PDF")',
           bold=True, h="center")
    ws_sum.conditional_formatting.add("B16:D16", FormulaRule(
        formula=['$B$16="All balances verified"'],
        font=Font(color="006100", bold=True), fill=PatternFill("solid", start_color="C6EFCE", end_color="C6EFCE")))
    ws_sum.conditional_formatting.add("B16:D16", FormulaRule(
        formula=['$B$16<>"All balances verified"'],
        font=Font(color="9C0006", bold=True), fill=PatternFill("solid", start_color="FFC7CE", end_color="FFC7CE")))

    notes = ["Closing Balance = Opening Balance + Total Credits - Total Debits."]
    if opening_derived:
        notes.append("Opening balance was not printed in the statement, so it was worked out from the first transaction.")
    if ocr_used:
        notes.append("Some pages were read with OCR. Please check rows marked CHECK against the original PDF.")
    for i, note in enumerate(notes):
        _merge(ws_sum, f"A{18 + i}:D{18 + i}", note, italic=True, size=9, color="595959", border=False)

    ws_sum.page_setup.orientation = "portrait"
    ws_sum.page_setup.fitToWidth = 1
    ws_sum.page_setup.fitToHeight = 0
    ws_sum.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# --- UNIVERSAL MULTI-BANK CONVERSION ENGINE ---
def process_pdf(pdf_file, pdf_password, progress_bar, status_text, timer_text,
                force_ocr=False, ocr_lang="eng", ocr_dpi=300):
    start_time = time.time()
    full_text = ""
    ocr_pages = 0
    pdf_bytes = pdf_file.getvalue()

    # If an image (PNG/JPG) was uploaded, convert it to a 1-page PDF first
    ext = pdf_file.name.lower().rsplit(".", 1)[-1]
    if ext in ("png", "jpg", "jpeg"):
        img_doc = fitz.open(stream=pdf_bytes, filetype=ext)
        pdf_bytes = img_doc.convert_to_pdf()
        img_doc.close()

    # 1. Open PDF
    try:
        pdf = pdfplumber.open(io.BytesIO(pdf_bytes), password=pdf_password if pdf_password else None)
        if len(pdf.pages) > 0:
            _ = pdf.pages[0].chars  # Trigger decryption check
    except PDFPasswordIncorrect:
        raise ValueError("Incorrect password or PDF is password-protected. Please enter the correct password.")
    except Exception as e:
        if "password" in str(e).lower():
            raise ValueError("Incorrect password or PDF is password-protected. Please enter the correct password.")
        raise ValueError(f"Failed to open PDF: {e}")

    mu_doc = None  # PyMuPDF document, opened only if OCR is needed

    try:
        # 2. Extract Text (normal first, OCR fallback per page)
        total_pages = len(pdf.pages)
        for idx, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            # A page counts as "real text" only if it has transaction-like amounts
            amount_count = len(re.findall(r'\d[\d,]*\.\d{2}', text))
            has_text = amount_count >= 3

            if force_ocr or not has_text:
                if not tesseract_available():
                    raise ValueError(
                        "This PDF needs OCR but Tesseract is not installed. "
                        "Install Tesseract (or add packages.txt on Streamlit Cloud)."
                    )
                if mu_doc is None:
                    mu_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                    if mu_doc.needs_pass:
                        mu_doc.authenticate(pdf_password or "")
                status_text.text(f"OCR running: Page {idx + 1} of {total_pages} (scanned page)...")
                img = render_page_image(mu_doc, idx, ocr_dpi)
                text = ocr_page_text(img, ocr_lang)
                ocr_pages += 1

            if text:
                full_text += text + "\n"

            elapsed = time.time() - start_time
            pct = int(((idx + 1) / total_pages) * 75)

            progress_bar.progress(pct)
            status_text.text(f"Scanning PDF: Page {idx + 1} of {total_pages} ({pct}%)")
            timer_text.text(f"Time Taken: {format_seconds(elapsed)}")
    finally:
        pdf.close()
        if mu_doc is not None:
            mu_doc.close()

    st.session_state["ocr_debug"] = full_text

    # 3. Parse Transactions
    status_text.text("Validating transactions & calculating balances (82%)...")
    progress_bar.progress(82)

    lines = full_text.split('\n')
    meta = extract_metadata(full_text)

    transactions = []
    current_txn = None

    date_rgx = r'(?:\d{1,2}[-/]\d{1,2}[-/]\d{2,4}|\d{1,2}[-\s][A-Za-z]{3}[-\s]\d{2,4}|[A-Za-z]{2,3}/[A-Za-z]{2,3}/[A-Za-z]{3,4})'
    line_start_pattern = re.compile(rf'^(?:\d+\s+)?({date_rgx})(?:\s+{date_rgx})?\s+(.+)$')

    amt_rgx = r'(?:₹|Rs\.?|INR)?\s*([\d,]+\.\d{2})\s*(?:Cr|Dr|CR|DR|cr|dr)?'
    three_amts_pattern = re.compile(rf'^(.*?)\s+{amt_rgx}\s+{amt_rgx}\s+{amt_rgx}$')
    two_amts_pattern = re.compile(rf'^(.*?)\s+{amt_rgx}\s+{amt_rgx}$')
    one_amt_pattern = re.compile(rf'^(.*?)\s+{amt_rgx}$')
    standalone_bal_pattern = re.compile(rf'^{amt_rgx}$')
    wrapped_two_amts_pattern = re.compile(rf'^(.*?)\s*{amt_rgx}\s+{amt_rgx}$')

    # Narrower list so real remark lines (ATM, cheque, etc.) are not deleted
    ignore_words = [
        "discrepancy", "passwords", "generated on", "computer generated",
        "statement of", "page ", "date remarks", "date narration",
        "date particulars", "date payment", "paid in", "brought forward", "carried forward"
    ]

    for line in lines:
        line = line.strip()
        line = re.sub(r'^[^\w₹]+', '', line)
        if not line or line.startswith('---'):
            continue

        start_match = line_start_pattern.match(line)
        if start_match:
            if current_txn and current_txn['Balance'] is not None and current_txn['Amount'] is not None:
                transactions.append(current_txn)

            txn_date = start_match.group(1)
            rest = start_match.group(2).strip()

            m3 = three_amts_pattern.match(rest)
            m2 = two_amts_pattern.match(rest)
            m1 = one_amt_pattern.match(rest)

            if m3:
                current_txn = {'Date': txn_date, 'Remarks': m3.group(1).strip(),
                               'Amount': float(m3.group(2).replace(',', '')) if float(m3.group(2).replace(',', '')) > 0 else float(m3.group(3).replace(',', '')),
                               'Balance': float(m3.group(4).replace(',', ''))}
            elif m2:
                current_txn = {'Date': txn_date, 'Remarks': m2.group(1).strip(),
                               'Amount': float(m2.group(2).replace(',', '')), 'Balance': float(m2.group(3).replace(',', ''))}
            elif m1:
                current_txn = {'Date': txn_date, 'Remarks': m1.group(1).strip(),
                               'Amount': float(m1.group(2).replace(',', '')), 'Balance': None}
            else:
                current_txn = {'Date': txn_date, 'Remarks': rest, 'Amount': None, 'Balance': None}
            continue

        if current_txn and current_txn['Balance'] is None:
            bal_match = standalone_bal_pattern.match(line)
            if bal_match and current_txn['Amount'] is not None:
                current_txn['Balance'] = float(bal_match.group(1).replace(',', ''))
                continue

            wrap_m2 = wrapped_two_amts_pattern.match(line)
            if wrap_m2 and current_txn['Amount'] is None:
                extra_rem = wrap_m2.group(1).strip()
                if extra_rem: current_txn['Remarks'] += " " + extra_rem
                current_txn['Amount'] = float(wrap_m2.group(2).replace(',', ''))
                current_txn['Balance'] = float(wrap_m2.group(3).replace(',', ''))
                continue

        if current_txn:
            lower_line = line.lower()
            if any(w in lower_line for w in ignore_words):
                continue
            current_txn['Remarks'] = (current_txn['Remarks'] + " " + line).strip()

    if current_txn and current_txn['Balance'] is not None and current_txn['Amount'] is not None:
        transactions.append(current_txn)

    if not transactions:
        if ocr_pages:
            raise ValueError("OCR ran but no transactions were detected. Try a higher OCR DPI (400), "
                             "check the OCR language, or use a clearer scan.")
        raise ValueError("Could not detect transactions in this statement.")

    # 4. Chronological Check
    reversed_order = False
    if len(transactions) > 1:
        first_dt = parse_any_date(transactions[0]['Date'])
        last_dt = parse_any_date(transactions[-1]['Date'])
        if first_dt and last_dt and first_dt > last_dt:
            transactions.reverse()
            reversed_order = True

    df = pd.DataFrame(transactions)

    # 5. Math Validation
    debits, credits = [], []
    mismatch_rows = 0
    ob = re.search(r'(?:Brought\s*Forward|Opening\s*Balance)[^\d\n]*([\d,]+\.\d{2})', full_text, re.IGNORECASE)
    opening_bal = float(ob.group(1).replace(',', '')) if (ob and not reversed_order) else None
    for i in range(len(df)):
        amt = df.loc[i, 'Amount']
        remarks = df.loc[i, 'Remarks'].upper()
        prev_bal = opening_bal if i == 0 else df.loc[i-1, 'Balance']
        if prev_bal is not None:
            diff = round(df.loc[i, 'Balance'] - prev_bal, 2)
            if abs(abs(diff) - amt) > 0.01:
                mismatch_rows += 1   # possible OCR misread
            if diff > 0:
                credits.append(amt); debits.append(None)
            elif diff < 0:
                debits.append(amt); credits.append(None)
            else:
                if '/CR/' in remarks or 'CREDIT' in remarks or 'BY ' in remarks or 'NEFT CR' in remarks:
                    credits.append(amt); debits.append(None)
                else:
                    debits.append(amt); credits.append(None)
        else:
            if '/CR/' in remarks or 'CREDIT' in remarks or 'BY ' in remarks or 'NEFT CR' in remarks:
                credits.append(amt); debits.append(None)
            else:
                debits.append(amt); credits.append(None)

    df['Debit'] = debits; df['Credit'] = credits
    df = df[['Date', 'Remarks', 'Debit', 'Credit', 'Balance']]

    # 6. Build the formatted Excel file
    status_text.text("Applying professional formatting (92%)...")
    progress_bar.progress(92)

    def _num(v):
        return float(v) if pd.notna(v) else 0.0

    opening_derived = opening_bal is None
    if opening_derived:
        # Work backwards from the first row: opening = balance + debit - credit
        r0 = df.iloc[0]
        opening_value = round(_num(r0['Balance']) + _num(r0['Debit']) - _num(r0['Credit']), 2)
    else:
        opening_value = opening_bal

    output = build_excel(df, meta, opening_value, opening_derived, ocr_pages > 0)

    progress_bar.progress(100)
    status_text.text(f"✅ Complete! Extracted {len(df)} transactions.")
    output.seek(0)
    return output, len(df), ocr_pages, total_pages, mismatch_rows


# --- STREAMLIT UI ---
st.set_page_config(page_title="Universal Bank Converter", page_icon="🏦")
st.title("🏦 Universal Bank Statement Converter")
st.write("Convert your PDF bank statements (including scanned / unselectable PDFs) directly to a formatted Excel file.")

uploaded_file = st.file_uploader("Upload Bank Statement (PDF or image)", type=["pdf", "png", "jpg", "jpeg"])
pdf_password = st.text_input("PDF Password (leave blank if none)", type="password")

with st.expander("⚙️ OCR settings (for scanned PDFs)"):
    force_ocr = st.checkbox("Force OCR on all pages", value=False,
                            help="Normally OCR runs only on pages that have no selectable text.")
    ocr_lang = st.text_input("OCR language", value="eng",
                             help="Example: eng+hin for English + Hindi, eng+guj for English + Gujarati")
    ocr_dpi = st.select_slider("OCR quality (DPI)", options=[200, 250, 300, 400], value=300,
                               help="Higher = more accurate but slower.")

if uploaded_file is not None:
    if st.button("Convert to Excel"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        timer_text = st.empty()

        try:
            excel_buffer, n_txn, ocr_pages, total_pages, mismatch = process_pdf(
                uploaded_file, pdf_password, progress_bar, status_text, timer_text,
                force_ocr=force_ocr, ocr_lang=ocr_lang, ocr_dpi=ocr_dpi
            )

            st.success("Conversion successful!")
            if ocr_pages:
                st.info(f"OCR was used on {ocr_pages} of {total_pages} pages.")
            if mismatch:
                st.warning(f"{mismatch} row(s) have a balance that doesn't match the amount. "
                           "This is usually an OCR misread, so please check those rows against the PDF.")
            st.download_button(
                label="📥 Download Excel File",
                data=excel_buffer,
                file_name=f"{uploaded_file.name.rsplit('.', 1)[0]}_Converted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error: {e}")

if st.session_state.get("ocr_debug"):
    with st.expander("🔍 Show OCR text (debug)"):
        st.text_area("What Tesseract read from your PDF",
                     st.session_state["ocr_debug"], height=400)
