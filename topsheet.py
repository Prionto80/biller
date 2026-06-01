from docx import Document
import pandas as pd
import re

# ===== FILE PATHS =====
DOCX_FILE = "invoices.docx"
OUTPUT_EXCEL = "invoice_summary.xlsx"

# ===== READ DOCX =====
doc = Document(DOCX_FILE)

# Collect text page-wise using page breaks
pages = []
current_page = []

for para in doc.paragraphs:
    text = para.text.strip()

    if text:
        current_page.append(text)

    # Detect page break
    for run in para.runs:
        if "lastRenderedPageBreak" in run._element.xml:
            pages.append("\n".join(current_page))
            current_page = []

# Add final page
if current_page:
    pages.append("\n".join(current_page))

# ===== EXTRACT DATA =====
data = []
grand_sum = 0

for page in pages:

    # -------- BILL NO --------
    bill_match = re.search(r"Bill\s*No\s*(\d+)", page, re.IGNORECASE)
    bill_no = bill_match.group(1) if bill_match else ""

    # -------- NAME --------
    # Usually line after company name contains branch/place name
    lines = page.split("\n")

    name = ""

    for line in lines:
        # Example:
        # Walton Hi-Tech Industries (PLC) Walton Plaza-Bazir Moor
        if "Walton Plaza" in line:
            name = line.split("Walton Hi-Tech Industries (PLC)")[-1].strip()
            break

    # Backup method if above fails
    if not name:
        for line in lines:
            if re.search(r"\d{11}", line):
                # remove phone number
                name = re.sub(r"\d{11}", "", line).strip()
                break

    # -------- GRAND TOTAL --------
    total_match = re.search(
        r"Grand\s*Total\s*=\s*([\d,]+)",
        page,
        re.IGNORECASE
    )

    total_amount = total_match.group(1) if total_match else "0"

    # Convert to integer
    numeric_total = int(total_amount.replace(",", ""))
    grand_sum += numeric_total

    # Store row
    data.append({
        "Name": name,
        "Bill No": bill_no,
        "Total Amount": numeric_total
    })

# ===== ADD FINAL TOTAL ROW =====
data.append({
    "Name": "",
    "Bill No": "TOTAL",
    "Total Amount": grand_sum
})

# ===== EXPORT TO EXCEL =====
df = pd.DataFrame(data)

with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="Invoices")

print(f"Excel file created: {OUTPUT_EXCEL}")