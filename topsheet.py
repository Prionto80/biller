import sys
import re
from pathlib import Path

from docx import Document
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
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


# =========================================================
# CREATE BEAUTIFUL TOP-SHEET
# =========================================================

def create_excel(invoices, output_file):

    wb = Workbook()
    ws = wb.active
    ws.title = "Top-Sheet"

    # -----------------------------------------------------
    # Dimensions
    # -----------------------------------------------------

    ws.column_dimensions["A"].width = 9
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 18

    # -----------------------------------------------------
    # Title
    # -----------------------------------------------------

    ws.merge_cells("A1:E1")

    title = ws["A1"]
    title.value = "Grameen Media Top-Sheet"

    title.font = Font(
        name="Arial",
        size=24,
        bold=True
    )

    title.alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    ws.row_dimensions[1].height = 42

    # -----------------------------------------------------
    # Bill Number Range
    # -----------------------------------------------------

    bill_numbers = []

    for invoice in invoices:

        try:
            bill_numbers.append(
                int(invoice["bill_no"])
            )
        except ValueError:
            pass

    if bill_numbers:

        first_bill = min(bill_numbers)
        last_bill = max(bill_numbers)

        bill_range = f"Bill No {first_bill}-{last_bill}"

    else:
        bill_range = "Bill No"

    ws.merge_cells("A2:E2")

    ws["A2"] = bill_range

    ws["A2"].font = Font(
        name="Arial",
        size=16
    )

    ws["A2"].alignment = Alignment(
        horizontal="left",
        vertical="center"
    )

    ws.row_dimensions[2].height = 30

    # -----------------------------------------------------
    # Header
    # -----------------------------------------------------

    header_row = 4

    headers = [
        "SL No",
        "Particulars",
        "Bill No",
        "Amount",
        "Remarks"
    ]

    for col, header in enumerate(headers, start=1):

        cell = ws.cell(
            row=header_row,
            column=col
        )

        cell.value = header

        cell.font = Font(
            name="Arial",
            size=12,
            bold=True
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    ws.row_dimensions[header_row].height = 28

    # -----------------------------------------------------
    # Borders
    # -----------------------------------------------------

    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    # -----------------------------------------------------
    # Invoice Data
    # -----------------------------------------------------

    start_row = 5

    for index, invoice in enumerate(
        invoices,
        start=1
    ):

        row = start_row + index - 1

        # SL No
        ws.cell(
            row=row,
            column=1,
            value=index
        )

        # Particulars
        ws.cell(
            row=row,
            column=2,
            value=invoice["showroom"]
        )

        # Bill No
        ws.cell(
            row=row,
            column=3,
            value=invoice["bill_no"]
        )

        # Amount
        ws.cell(
            row=row,
            column=4,
            value=invoice["amount"]
        )

        # Remarks
        ws.cell(
            row=row,
            column=5,
            value=""
        )

        # Formatting
        for col in range(1, 6):

            cell = ws.cell(
                row=row,
                column=col
            )

            cell.font = Font(
                name="Arial",
                size=11
            )

            cell.border = thin_border

            cell.alignment = Alignment(
                vertical="center"
            )

        # Center specific columns
        ws.cell(row=row, column=1).alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        ws.cell(row=row, column=3).alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        ws.cell(row=row, column=4).alignment = Alignment(
            horizontal="right",
            vertical="center"
        )

        # Amount formatting
        ws.cell(
            row=row,
            column=4
        ).number_format = '#,##0'

        ws.row_dimensions[row].height = 24

    # -----------------------------------------------------
    # Total Row
    # -----------------------------------------------------

    total_row = start_row + len(invoices)

    # Merge A:B
    ws.merge_cells(
        start_row=total_row,
        start_column=1,
        end_row=total_row,
        end_column=2
    )

    total_label = ws.cell(
        row=total_row,
        column=1
    )

    total_label.value = "Total Amount ="

    total_label.font = Font(
        name="Arial",
        size=12,
        bold=True
    )

    total_label.alignment = Alignment(
        horizontal="right",
        vertical="center"
    )

    # Blank Bill No cell
    ws.cell(
        row=total_row,
        column=3
    ).value = ""

    # Total amount
    total_amount = sum(
        invoice["amount"]
        for invoice in invoices
    )

    total_cell = ws.cell(
        row=total_row,
        column=4
    )

    total_cell.value = total_amount

    total_cell.number_format = '#,##0'

    total_cell.font = Font(
        name="Arial",
        size=12,
        bold=True
    )

    total_cell.alignment = Alignment(
        horizontal="right",
        vertical="center"
    )

    # Remarks cell
    ws.cell(
        row=total_row,
        column=5
    ).value = ""

    # Borders for total row
    for col in range(1, 6):

        cell = ws.cell(
            row=total_row,
            column=col
        )

        cell.border = thin_border

    ws.row_dimensions[total_row].height = 28

    # -----------------------------------------------------
    # Apply borders to title + bill range
    # -----------------------------------------------------

    for row in range(1, 3):

        for col in range(1, 6):

            ws.cell(
                row=row,
                column=col
            ).border = thin_border

    # -----------------------------------------------------
    # Page Setup
    # -----------------------------------------------------

    ws.freeze_panes = "A5"

    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    ws.page_setup.orientation = "landscape"

    ws.print_title_rows = "1:4"

    # -----------------------------------------------------
    # Save
    # -----------------------------------------------------

    wb.save(output_file)


# =========================================================
# MAIN
# =========================================================

def main():

    if len(sys.argv) != 2:

        print(
            "Usage:\n"
            "python invoice_to_excel.py <input.docx>"
        )

        return

    input_file = Path(sys.argv[1])

    if not input_file.exists():

        print(
            f"File not found: {input_file}"
        )

        return

    # Automatically generate output filename
    output_file = input_file.with_name(
        input_file.stem + "_top_sheet.xlsx"
    )

    print(f"Reading: {input_file}")

    # Extract invoices
    raw_invoices = extract_invoices(
        input_file
    )

    print(
        f"Found {len(raw_invoices)} invoices."
    )

    # Parse invoices
    invoices = []

    for lines in raw_invoices:

        invoice = parse_invoice(lines)

        invoices.append(invoice)

        print(
            f'{invoice["bill_no"]} | '
            f'{invoice["showroom"]} | '
            f'{invoice["amount"]:,}'
        )

    # Create Excel
    create_excel(
        invoices,
        output_file
    )

    print()
    print("================================")
    print("Done!")
    print("================================")
    print(f"Output: {output_file}")


if __name__ == "__main__":
    main()