import sys
import re
from pathlib import Path
from typing import List, Dict, Any, Union, Tuple, Optional

from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter


# =========================================================
# NUMBER WORDS → NUMBER
# =========================================================

ONES = {
    "zero": 0,
    "one": 1,
    "two": 2,
    "three": 3,
    "four": 4,
    "five": 5,
    "six": 6,
    "seven": 7,
    "eight": 8,
    "nine": 9,
    "ten": 10,
    "eleven": 11,
    "twelve": 12,
    "thirteen": 13,
    "fourteen": 14,
    "fifteen": 15,
    "sixteen": 16,
    "seventeen": 17,
    "eighteen": 18,
    "nineteen": 19,
}

TENS = {
    "twenty": 20,
    "thirty": 30,
    "forty": 40,
    "fifty": 50,
    "sixty": 60,
    "seventy": 70,
    "eighty": 80,
    "ninety": 90,
}


def words_to_number(text):
    text = text.lower()
    text = text.replace("-", " ")
    text = text.replace(",", " ")

    words = text.split()

    total = 0
    current = 0

    for word in words:

        if word == "and":
            continue

        if word in ONES:
            current += ONES[word]

        elif word in TENS:
            current += TENS[word]

        elif word == "hundred":
            current *= 100

        elif word == "thousand":
            total += current * 1000
            current = 0

        elif word == "lakh":
            total += current * 100000
            current = 0

        elif word == "million":
            total += current * 1000000
            current = 0

    return total + current


# =========================================================
# EXTRACT INVOICES
# =========================================================

def extract_invoices(docx_file):

    doc = Document(docx_file)

    paragraphs = []

    for p in doc.paragraphs:
        text = p.text.strip()

        if text:
            paragraphs.append(text)

    invoices = []
    current = []

    for text in paragraphs:

        # New invoice
        if re.match(r"Bill No\s+\d+", text, re.IGNORECASE):

            if current:
                invoices.append(current)

            current = [text]

        else:
            if current:
                current.append(text)

    if current:
        invoices.append(current)

    return invoices


# =========================================================
# PARSE INVOICE
# =========================================================

def parse_invoice(lines):

    # -----------------------------------------------------
    # Bill Number
    # -----------------------------------------------------

    bill_no = ""

    for line in lines:

        match = re.search(
            r"Bill No\s+(\d+)",
            line,
            re.IGNORECASE
        )

        if match:
            bill_no = match.group(1)
            break

    # -----------------------------------------------------
    # Showroom Name
    # -----------------------------------------------------

    showroom = ""

    for i, line in enumerate(lines):

        if "(COM)" in line.upper():

            after_com = re.split(
                r"\(COM\)",
                line,
                flags=re.IGNORECASE
            )[1].strip()

            if after_com:

                showroom = after_com

                # Some showroom names continue
                # on the next line
                j = i + 1

                while j < len(lines):

                    next_line = lines[j]

                    # Stop at address/date/etc.
                    if (
                        next_line.lower().startswith("date")
                        or next_line.lower().startswith("pr no")
                        or next_line.lower().startswith("subject")
                        or next_line.lower().startswith("in words")
                        or next_line.lower() == "thank you"
                    ):
                        break

                    # Phone number
                    if re.fullmatch(r"[\d\s+-]+", next_line):
                        break

                    showroom += " " + next_line

                    j += 1

            break

    # -----------------------------------------------------
    # Total Amount
    # -----------------------------------------------------

    amount = 0

    for line in lines:

        if "In words:" in line:

            words = re.split(
                r"In words:\s*",
                line,
                flags=re.IGNORECASE
            )[1]

            words = re.sub(
                r"\s*taka\s*only\.?",
                "",
                words,
                flags=re.IGNORECASE
            ).strip()

            amount = words_to_number(words)

            break

    return {
        "showroom": showroom,
        "bill_no": bill_no,
        "amount": amount,
    }


# (Old 5-column create_excel replaced by official 8-column layout below)


# =========================================================
# BANGLADESH DISTRICTS & OFFICIAL 8-COLUMN TOPSHEET
# =========================================================

BANGLADESH_DISTRICTS = [
    "Dhaka", "Faridpur", "Gazipur", "Gopalganj", "Gopalgonj", "Kishoreganj", "Kishoregonj", "Madaripur", "Manikganj", "Manikgonj", 
    "Munshiganj", "Munshigonj", "Narayanganj", "Narayangonj", "N.gonj", "N.ganj", "Narsingdi", "Rajbari", "Shariatpur", "Tangail",
    "Bagerhat", "Chuadanga", "Jashore", "Jessore", "Jhenaidah", "Khulna", "Kushtia", "Magura", "Meherpur", "Narail", "Satkhira",
    "Bogura", "Bogra", "Joypurhat", "Naogaon", "Natore", "Chapainawabganj", "Chapainawabgonj", "Nawabganj", "Nawabgonj", "Pabna", "Rajshahi", "Sirajganj", "Sirajgonj",
    "Dinajpur", "Gaibandha", "Kurigram", "Lalmonirhat", "Nilphamari", "Panchagarh", "Rangpur", "Thakurgaon",
    "Habiganj", "Habigonj", "Moulvibazar", "Moulvibazer", "Sunamganj", "Sunamgonj", "Sylhet",
    "Barguna", "Barishal", "Barisal", "Bhola", "Jhalokathi", "Patuakhali", "Pirojpur",
    "Bandarban", "Brahmanbaria", "B.baria", "Chandpur", "Chattogram", "Chittagong", "Ctg", "Cumilla", "Comilla", "Cox's Bazar", "Coxsbazar", "Feni", "Khagrachhari", "Lakshmipur", "Noakhali", "Rangamati",
    "Mymensingh", "Jamalpur", "Netrokona", "Netrakona", "Sherpur"
]


def find_district(text: str) -> str:
    """Extract district name from address or showroom text using Bangladesh 64 districts."""
    if not text:
        return ""
    clean_text = text.replace(",", " ").replace("-", " ").replace("/", " ").replace("(", " ").replace(")", " ")
    words = clean_text.split()
    
    district_map = {
        "comilla": "Cumilla", "cumilla": "Cumilla",
        "ctg": "Chattogram", "chittagong": "Chattogram", "chattogram": "Chattogram",
        "coxsbazar": "Cox's Bazar", "cox's bazar": "Cox's Bazar",
        "kishoregonj": "Kishoreganj", "kishoreganj": "Kishoreganj",
        "netrakona": "Netrokona", "netrokona": "Netrokona",
        "bogra": "Bogura", "bogura": "Bogura",
        "jessore": "Jashore", "jashore": "Jashore",
        "barisal": "Barishal", "barishal": "Barishal",
        "narayanganj": "Narayanganj", "narayangonj": "Narayanganj", "n.gonj": "Narayanganj", "n.ganj": "Narayanganj",
        "chapainawabganj": "Chapainawabganj", "chapainawabgonj": "Chapainawabganj", "nawabganj": "Chapainawabganj", "nawabgonj": "Chapainawabganj",
        "gopalganj": "Gopalganj", "gopalgonj": "Gopalganj",
        "manikganj": "Manikganj", "manikgonj": "Manikganj",
        "munshiganj": "Munshiganj", "munshigonj": "Munshiganj",
        "sirajganj": "Sirajganj", "sirajgonj": "Sirajganj",
        "habiganj": "Habiganj", "habigonj": "Habiganj",
        "sunamganj": "Sunamganj", "sunamgonj": "Sunamganj",
        "moulvibazar": "Moulvibazar", "moulvibazer": "Moulvibazar",
        "b.baria": "Brahmanbaria", "brahmanbaria": "Brahmanbaria",
    }

    # Check backwards since district is usually the final token in the address
    for w in reversed(words):
        w_strip = re.sub(r"[^a-zA-Z.]", "", w).lower()
        if w_strip in district_map:
            return district_map[w_strip]
        for d in BANGLADESH_DISTRICTS:
            if w_strip == d.lower():
                return district_map.get(d.lower(), d)

    for d in BANGLADESH_DISTRICTS:
        if re.search(r"\b" + re.escape(d) + r"\b", text, re.IGNORECASE):
            return district_map.get(d.lower(), d)
    return ""


def clean_company_name(comp_text: str, outlet_text: str = "", brand_code: str = "") -> str:
    """Format company name matching the client format (e.g. Hi-Tech (COM), Digi-Tech (MOB), Hi-Tech (PLAZA))."""
    if not comp_text:
        comp_text = "Hi-Tech (COM)"
    comp_upper = comp_text.upper()
    code_upper = brand_code.upper() if brand_code else ""
    
    if "DIGI-TECH" in comp_upper or "MOBILE" in comp_upper or "(MOB)" in comp_upper or "MOB" in code_upper:
        return "Digi-Tech (MOB)"
    
    if "(PLAZA)" in comp_upper or "PLAZA" in code_upper:
        return "Hi-Tech (PLAZA)"
        
    if "HI-TECH" in comp_upper or "WALTON" in comp_upper:
        return "Hi-Tech (COM)"
        
    return comp_text.strip()


def normalize_work_type(raw_desc: str) -> str:
    """Normalize work type description matching client sheet (e.g. signboard, signboard relocating, signboard structure, pana change)."""
    if not raw_desc:
        return "signboard"
    desc = raw_desc.strip().lower()
    if "relocating" in desc or "relocate" in desc:
        return "signboard relocating"
    if "structure" in desc or "staucture" in desc:
        return "signboard structure"
    if "pana change" in desc:
        return "pana change"
    if "signboard" in desc:
        return "signboard"
    if "light box" in desc or "lightbox" in desc:
        return "light box"
    return desc


def parse_merged_docx(docx_input) -> List[dict]:
    """
    Parse a merged .docx file (or single .docx bill) and extract all invoices into
    structured records with the 8 columns required for the official Top Sheet.
    Works with file path or file-like objects (e.g. BytesIO, UploadFile.file).
    """
    doc = Document(docx_input)
    
    # Iterate through body elements in document order
    # to associate each bill's paragraphs with its table
    bills_data = []
    current_bill = None
    
    for child in doc.element.body:
        if child.tag.endswith("p"):
            p = Paragraph(child, doc)
            text = p.text.strip()
            
            # Start of a new bill
            if re.search(r"^Bill\s*No\s*\d+", text, re.IGNORECASE):
                if current_bill:
                    bills_data.append(current_bill)
                current_bill = {
                    "paragraphs": [text],
                    "tables": []
                }
            elif current_bill and text:
                current_bill["paragraphs"].append(text)
                
        elif child.tag.endswith("tbl"):
            t = Table(child, doc)
            if current_bill:
                current_bill["tables"].append(t)
            else:
                current_bill = {
                    "paragraphs": [],
                    "tables": [t]
                }
                
    if current_bill:
        bills_data.append(current_bill)
        
    # If no Bill No paragraph was detected, fallback to paragraph splitting
    if not bills_data:
        curr = []
        for p in doc.paragraphs:
            txt = p.text.strip()
            if re.search(r"Bill\s*No\s*\d+", txt, re.IGNORECASE):
                if curr:
                    bills_data.append({"paragraphs": curr, "tables": []})
                curr = [txt]
            elif curr and txt:
                curr.append(txt)
        if curr:
            bills_data.append({"paragraphs": curr, "tables": []})
            
    records = []
    for i, bill in enumerate(bills_data, start=1):
        lines = bill.get("paragraphs", [])
        tables = bill.get("tables", [])
        
        bno = ""
        comp = ""
        outlet = ""
        addr = ""
        pr = ""
        work_type = ""
        amount = 0
        
        # 1. Bill Number
        for line in lines:
            m = re.search(r"Bill\s*No\s*(\d+)", line, re.IGNORECASE)
            if m:
                bno = m.group(1)
                break
                
        # 2. Company & Outlet Name (P2)
        for line in lines:
            if "WALTON" in line.upper() or "HI-TECH" in line.upper() or "DIGI-TECH" in line.upper():
                parts = [p.strip() for p in re.split(r"\s{3,}|\t", line) if p.strip()]
                if parts:
                    comp = parts[0]
                    if len(parts) > 1:
                        outlet = parts[1]
                break
                
        # 3. Address (P3)
        for line in lines:
            if (
                line != comp
                and line != outlet
                and not line.startswith("Bill No")
                and not line.startswith("Date")
                and not line.startswith("PR No")
                and not line.startswith("Subject")
                and not line.startswith("In words")
                and not line.startswith("Thank you")
                and not line.startswith("Grameen Media")
            ):
                if not any(k in line for k in ["Walton", "Hi-Tech", "Digi-Tech"]):
                    parts = [p.strip() for p in re.split(r"\s{3,}|\t", line) if p.strip()]
                    if parts:
                        addr = parts[0]
                    break
                    
        # 4. PR / TN Number
        for line in lines:
            m = re.search(r"(?:PR|TN)\s*(?:No)?[:\s/]+([^\n\r]+)", line, re.IGNORECASE)
            if m:
                pr = m.group(1).strip()
                break
                
        # 5. Subject
        for line in lines:
            m = re.search(r"Subject[:\s]+(?:Bill for\s*(?:supply of\s*)?)?([^\n\r.]+)", line, re.IGNORECASE)
            if m:
                work_type = m.group(1).strip()
                break
                
        # 6. Table extraction (particles & total amount)
        table_particles = []
        table_amount = 0
        if tables:
            tbl = tables[0]
            for row in tbl.rows:
                cells = [c.text.strip() for c in row.cells]
                if len(cells) >= 2 and cells[0].isdigit():
                    table_particles.append(cells[1])
                row_str = " ".join(cells)
                if "Grand Total" in row_str or "Total Amount" in row_str:
                    last_cell = cells[-1].replace("/-", "").replace(",", "").strip()
                    m_num = re.search(r"(\d+)", last_cell)
                    if m_num:
                        val = int(m_num.group(1))
                        if "Grand Total" in row_str:
                            table_amount = val
                        elif table_amount == 0:
                            table_amount = val
        elif i - 1 < len(doc.tables):
            tbl = doc.tables[i - 1]
            for row in tbl.rows:
                cells = [c.text.strip() for c in row.cells]
                if len(cells) >= 2 and cells[0].isdigit():
                    table_particles.append(cells[1])
                row_str = " ".join(cells)
                if "Grand Total" in row_str or "Total Amount" in row_str:
                    last_cell = cells[-1].replace("/-", "").replace(",", "").strip()
                    m_num = re.search(r"(\d+)", last_cell)
                    if m_num:
                        val = int(m_num.group(1))
                        if "Grand Total" in row_str:
                            table_amount = val
                        elif table_amount == 0:
                            table_amount = val
                            
        # Work type resolution
        if table_particles:
            work_type = normalize_work_type(table_particles[0])
        elif work_type:
            work_type = normalize_work_type(work_type)
        else:
            work_type = "Signboard"
            
        # Amount resolution
        if table_amount > 0:
            amount = table_amount
        else:
            for line in lines:
                if "In words:" in line:
                    words = re.split(r"In words:\s*", line, flags=re.IGNORECASE)[1]
                    words = re.sub(r"\s*taka\s*only\.?", "", words, flags=re.IGNORECASE).strip()
                    amt = words_to_number(words)
                    if amt > 0:
                        amount = amt
                        break
                m_gt = re.search(r"Grand\s*Total\s*=\s*([\d,]+)", line, re.IGNORECASE)
                if m_gt:
                    amount = int(m_gt.group(1).replace(",", ""))
                    break
                    
        # District resolution
        district = find_district(addr) or find_district(outlet) or ""
        
        # Combined Outlet Name + Outlet Address
        if outlet and addr:
            if addr.lower() in outlet.lower():
                combined_outlet_addr = outlet
            elif outlet.lower() in addr.lower():
                combined_outlet_addr = addr
            else:
                combined_outlet_addr = f"{outlet}, {addr}" if "-" in outlet else f"{outlet}-{addr}"
        elif outlet:
            combined_outlet_addr = outlet
        else:
            combined_outlet_addr = addr
            
        final_company = clean_company_name(comp, outlet)
        
        records.append({
            "sl": i,
            "company_name": final_company,
            "work_type": work_type,
            "district": district,
            "outlet_address": combined_outlet_addr,
            "pr_tn": pr,
            "bill_no": bno,
            "total_amount": amount,
        })
        
    return records


def create_official_topsheet(records: List[dict], output_file: Union[str, Path]):
    """
    Generate an Excel Top-Sheet matching the client format shown in media_1793480112419_0904e18b.png:
      Row 1: Grameen Media Top-Sheet (A1:H1 merged, bold, centered)
      Row 2: Bill No {min}-{max} (A2:H2 merged, bold, left-aligned)
      Row 3: (Blank separator)
      Row 4: Headers: SL No | Company Name | Work Type | District | Outlet + Address | PR/TN | Bill No | Amount
      Rows 5+: Data rows with solid thin borders
      Total Row: Merged E:F for 'Total Amount =', blank Col G, sum in Col H.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    
    # Page Setup
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "1:4"
    
    # Column dimensions
    col_widths = {
        "A": 8.0,    # SL No
        "B": 18.0,   # Company Name
        "C": 22.0,   # Work Type
        "D": 16.0,   # District
        "E": 52.0,   # Outlet + Address
        "F": 15.0,   # PR/TN
        "G": 12.0,   # Bill No
        "H": 15.0,   # Amount
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width
        
    thin_side = Side(style="thin", color="000000")
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    
    # -----------------------------------------------------
    # Row 1: Grameen Media Top-Sheet
    # -----------------------------------------------------
    ws.merge_cells("A1:H1")
    title_cell = ws["A1"]
    title_cell.value = "Grameen Media Top-Sheet"
    title_cell.font = Font(name="Arial", size=18, bold=True, color="000000")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36
    
    for col in range(1, 9):
        ws.cell(row=1, column=col).border = thin_border
        
    # -----------------------------------------------------
    # Row 2: Bill No Range (e.g. Bill No 3490-3502)
    # -----------------------------------------------------
    bill_numbers = []
    for r in records:
        bno_raw = str(r.get("bill_no", "")).strip()
        m = re.search(r"\d+", bno_raw)
        if m:
            try:
                bill_numbers.append(int(m.group(0)))
            except ValueError:
                pass
                
    if bill_numbers:
        min_b = min(bill_numbers)
        max_b = max(bill_numbers)
        bill_range_str = f"Bill No {min_b}-{max_b}" if min_b != max_b else f"Bill No {min_b}"
    else:
        bill_range_str = "Bill No"
        
    ws.merge_cells("A2:H2")
    range_cell = ws["A2"]
    range_cell.value = bill_range_str
    range_cell.font = Font(name="Arial", size=13, bold=True, color="000000")
    range_cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 26
    
    for col in range(1, 9):
        ws.cell(row=2, column=col).border = thin_border
        
    # -----------------------------------------------------
    # Row 3: Blank separator row
    # -----------------------------------------------------
    ws.row_dimensions[3].height = 14
    
    # -----------------------------------------------------
    # Row 4: Column Headers
    # -----------------------------------------------------
    headers = [
        "SL No",
        "Company Name",
        "Work Type",
        "District",
        "Outlet + Address",
        "PR/TN",
        "Bill No",
        "Amount"
    ]
    
    header_font = Font(name="Arial", size=10, bold=True, color="000000")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[4].height = 26
    
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=col_idx, value=header)
        cell.font = header_font
        cell.alignment = header_align
        cell.border = thin_border
        
    # -----------------------------------------------------
    # Rows 5+: Data Rows
    # -----------------------------------------------------
    data_font = Font(name="Arial", size=9, bold=False, color="000000")
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    
    start_row = 5
    for idx, item in enumerate(records, start=1):
        row_num = start_row + idx - 1
        ws.row_dimensions[row_num].height = 22
        
        # 1. SL No
        c1 = ws.cell(row=row_num, column=1, value=item.get("sl", idx))
        c1.alignment = align_center
        
        # 2. Company Name
        c2 = ws.cell(row=row_num, column=2, value=item.get("company_name", ""))
        c2.alignment = align_center
        
        # 3. Work Type
        work_type_val = item.get("work_type", "")
        if work_type_val.lower() == "signboard":
            work_type_val = "signboard"
        elif work_type_val.lower() == "signboard structure":
            work_type_val = "signboard structure"
        c3 = ws.cell(row=row_num, column=3, value=work_type_val)
        c3.alignment = align_left
        
        # 4. District
        c4 = ws.cell(row=row_num, column=4, value=item.get("district", ""))
        c4.alignment = align_left
        
        # 5. Outlet + Address
        c5 = ws.cell(row=row_num, column=5, value=item.get("outlet_address", ""))
        c5.alignment = align_left
        
        # 6. PR/TN
        pr_val = item.get("pr_tn", "")
        c6 = ws.cell(row=row_num, column=6, value=pr_val)
        c6.alignment = align_right
        
        # 7. Bill No
        c7 = ws.cell(row=row_num, column=7, value=item.get("bill_no", ""))
        c7.alignment = align_center
        
        # 8. Amount
        amt_raw = item.get("total_amount", 0)
        try:
            amt_num = int(round(float(amt_raw)))
        except (ValueError, TypeError):
            amt_num = 0
        c8 = ws.cell(row=row_num, column=8, value=amt_num)
        c8.alignment = align_right
        c8.number_format = "#,##0"
        
        for c_idx in range(1, 9):
            cell = ws.cell(row=row_num, column=c_idx)
            cell.font = data_font
            cell.border = thin_border
            
    # -----------------------------------------------------
    # Total Row
    # -----------------------------------------------------
    total_row = start_row + len(records)
    ws.row_dimensions[total_row].height = 24
    
    total_amount = sum(
        int(round(float(r.get("total_amount", 0) or 0)))
        for r in records
    )
    
    # Merge E to F for "Total Amount ="
    ws.merge_cells(start_row=total_row, start_column=5, end_row=total_row, end_column=6)
    label_cell = ws.cell(row=total_row, column=5, value="Total Amount =")
    label_cell.font = Font(name="Arial", size=10, bold=True, color="000000")
    label_cell.alignment = Alignment(horizontal="right", vertical="center")
    
    # Border for Col 5 and 6
    for col in (5, 6):
        ws.cell(row=total_row, column=col).border = thin_border
        
    # Blank Bill No in Col 7
    c_blank = ws.cell(row=total_row, column=7, value="")
    c_blank.border = thin_border
    
    # Total sum in Col 8
    total_cell = ws.cell(row=total_row, column=8, value=total_amount)
    total_cell.font = Font(name="Arial", size=10, bold=True, color="000000")
    total_cell.alignment = Alignment(horizontal="right", vertical="center")
    total_cell.number_format = "#,##0"
    total_cell.border = thin_border
    
    wb.save(output_file)
    return output_file


def generate_official_topsheet_from_docx(docx_input, output_excel_path: Union[str, Path]) -> Tuple[List[dict], Path]:
    """Parse docx input and write official 8-column Excel topsheet."""
    records = parse_merged_docx(docx_input)
    out_file = create_official_topsheet(records, output_excel_path)
    return records, Path(out_file)


def create_excel(invoices: List[dict], output_file: Union[str, Path]):
    """
    Generate Top-Sheet matching the client format in media_1793480112419_0904e18b.png.
    Accepts full 8-column records or legacy invoice dicts.
    """
    records = []
    for idx, inv in enumerate(invoices, start=1):
        if "outlet_address" in inv:
            records.append(inv)
        else:
            showroom = inv.get("showroom", "")
            bno = inv.get("bill_no", "")
            amt = inv.get("amount", 0)
            comp = clean_company_name(inv.get("company", inv.get("brand", "")))
            dist = find_district(showroom)
            records.append({
                "sl": idx,
                "company_name": comp,
                "work_type": inv.get("work_type", "signboard"),
                "district": dist,
                "outlet_address": showroom,
                "pr_tn": inv.get("pr_no", ""),
                "bill_no": str(bno),
                "total_amount": amt,
            })
    return create_official_topsheet(records, output_file)


# =========================================================
# MAIN
# =========================================================

def main():
    if len(sys.argv) != 2:
        print("Usage:\npython topsheet.py <input.docx>")
        return

    input_file = Path(sys.argv[1])
    if not input_file.exists():
        print(f"File not found: {input_file}")
        return

    # Automatically generate output filename
    output_file = input_file.with_name(
        input_file.stem + "_top_sheet.xlsx"
    )

    print(f"Reading: {input_file}")

    records, out_path = generate_official_topsheet_from_docx(input_file, output_file)
    print(f"Found {len(records)} invoices.")
    for r in records:
        print(
            f"{r.get('bill_no')} | "
            f"{r.get('company_name')} | "
            f"{r.get('outlet_address')} | "
            f"{r.get('total_amount', 0):,}"
        )

    print()
    print("================================")
    print("Done!")
    print("================================")
    print(f"Output: {output_file}")


if __name__ == "__main__":
    main()