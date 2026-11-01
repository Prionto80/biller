import os
import re
import json
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import List, Optional, Union

from fastapi import FastAPI, Request, Form, HTTPException, Query, UploadFile, File
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel, Field
from docxtpl import DocxTemplate

from inwords import number_to_words
from scanner import scan_workorder_image
import project_manager as pm
from topsheet import (
    parse_merged_docx,
    create_official_topsheet,
    generate_official_topsheet_from_docx,
    find_district,
    clean_company_name,
)

# Paths
BASE_DIR = Path(__file__).resolve().parent
TEMPLATES_DIR = BASE_DIR / "templates"
STATIC_DIR = BASE_DIR / "static"
OUTPUT_DIR = BASE_DIR / "output"
TEMPLATE_DOCX = BASE_DIR / "bill.docx"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
STATIC_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(
    title="Biller - Professional Invoice & Bill Generator",
    description="FastAPI service for generating professional bills and Word documents.",
    version="1.0.0",
)

# Mount static files and templates
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")
templates = Jinja2Templates(directory=str(TEMPLATES_DIR))


# -------------------------------------------------------------
# Data Models & Schemas
# -------------------------------------------------------------

class BoardItem(BaseModel):
    w: float = Field(..., description="Width in feet")
    h: float = Field(..., description="Height in feet")
    q: Union[int, str, float] = Field(default=1, description="Quantity")
    r: Optional[float] = Field(default=None, description="Rate in BDT (auto-set from description if not typed)")
    description: Optional[str] = Field(default="New Signboard (Horizontal)", description="Board particle description")


class BillPayload(BaseModel):
    bill_no: str
    pr_no: str
    brand: str
    showroom: str
    address: str
    mobile: str
    boards_specifications: Optional[str] = ""
    boards: Optional[List[BoardItem]] = None
    date: Optional[str] = None
    words: Optional[str] = None
    subject: Optional[str] = None
    project_id: Optional[str] = Field(default="default-project", description="Target project stack ID")


class ProjectCreatePayload(BaseModel):
    name: str = Field(..., description="Project Name")


# -------------------------------------------------------------
# Helper Functions & Rate Constants
# -------------------------------------------------------------

def resolve_brand(brand_input: str):
    """
    Resolve brand code and display name based on user input or preset keys:
    'm' -> Walton Digi-Tech Mobile (MOB)
    'p' -> Walton Hi-Tech Industries (PLC) (PLAZA)
    'h' -> Walton Hi-Tech Industries (PLC) (HAP)
    'e' -> Walton Hi-Tech Industries (PLC) (EAP)
    'c' or default -> Walton Hi-Tech Industries (PLC) (COM)
    """
    if not brand_input:
        return "Walton Hi-Tech Industries (PLC)", "(COM)"

    b = brand_input.strip().lower()

    if b in ("m", "mob", "mobile") or "mobile" in b:
        return "Walton Digi-Tech Mobile", "(MOB)"
    elif b in ("p", "plaza") or "plaza" in b:
        return "Walton Hi-Tech Industries (PLC)", "(PLAZA)"
    elif b in ("h", "hap") or "hap" in b:
        return "Walton Hi-Tech Industries (PLC)", "(HAP)"
    elif b in ("e", "eap") or "eap" in b:
        return "Walton Hi-Tech Industries (PLC)", "(EAP)"
    elif b in ("c", "com") or "com" in b:
        return "Walton Hi-Tech Industries (PLC)", "(COM)"
    else:
        match = re.search(r"\(([^)]+)\)", brand_input)
        if match:
            code = match.group(0)
            name = brand_input.replace(code, "").strip()
            return name, code
        return brand_input.strip(), "(COM)"


def compute_sft(w: float, h: float, q_val: Union[int, str, float]) -> float:
    """Calculate square feet with quantity support."""
    try:
        q_str = str(q_val).strip()
        if q_str.isdigit():
            return round(w * h * int(q_str), 2)
        elif "/" in q_str:
            parts = q_str.split("/")
            multiplier = float(parts[0]) / float(parts[1])
            return round(w * h * multiplier, 2)
        else:
            return round(w * h * float(q_str), 2)
    except Exception:
        return round(w * h * 1.0, 2)


def compute_qty(q_val: Union[int, str, float]) -> int:
    """Extract integer quantity."""
    try:
        q_str = str(q_val).strip()
        if q_str.isdigit():
            return int(q_str)
        elif len(q_str) > 0 and q_str[0].isdigit():
            return int(q_str[0])
        return 1
    except Exception:
        return 1


BOARD_TYPE_HORIZONTAL = "New Signboard (Horizontal)"
BOARD_TYPE_VERTICAL = "New Signboard (Vertical)"
BOARD_TYPE_PANA = "Pana Change and Servicing"

DEFAULT_BOARD_RATES = {
    BOARD_TYPE_HORIZONTAL: 365.0,
    BOARD_TYPE_VERTICAL: 270.0,
    BOARD_TYPE_PANA: 88.0,
}

def get_rate_for_description(desc: Optional[str]) -> float:
    """
    Get the standard rate based on description:
    - New Signboard (Horizontal): 365
    - New Signboard (Vertical): 270
    - Pana Change and Servicing: 88
    """
    d = (desc or "").strip().lower()
    if "vertical" in d:
        return 270.0
    elif "horizontal" in d:
        return 365.0
    elif "pana" in d:
        return 88.0
    elif "new signboard" in d:
        return 365.0
    return 365.0


def normalize_board_description(desc: Optional[str], rate: Optional[float] = None) -> str:
    """
    Ensure the board description matches one of the standard board types:
    - New Signboard (Horizontal)
    - New Signboard (Vertical)
    - Pana Change and Servicing
    """
    d = (desc or "").strip().lower()
    if "vertical" in d:
        return BOARD_TYPE_VERTICAL
    elif "horizontal" in d:
        return BOARD_TYPE_HORIZONTAL
    elif "pana" in d or "servicing" in d:
        return BOARD_TYPE_PANA
    elif "new signboard" in d:
        return BOARD_TYPE_HORIZONTAL
    elif desc and desc.strip():
        return desc.strip()

    if rate is not None:
        if rate == 270:
            return BOARD_TYPE_VERTICAL
        elif rate == 88:
            return BOARD_TYPE_PANA
        elif rate == 365 or rate > 95:
            return BOARD_TYPE_HORIZONTAL
        elif rate > 50:
            return BOARD_TYPE_PANA

    return BOARD_TYPE_HORIZONTAL


def compute_description(r: float, current_desc: Optional[str] = None) -> str:
    """Determine board description from current description or rate."""
    return normalize_board_description(current_desc, r)


def parse_boards_specifications(text_input: str) -> List[dict]:
    """
    Parses human-entered board specifications text into structured board items.
    Supports formats like:
      - '10 x 5, 2, 365'
      - '10 x 5, 2' (auto-assigns rate from description or default 365)
      - '10x5 q:2 r:270 "New Signboard (Vertical)"'
    """
    if not text_input or not text_input.strip():
        return []

    boards = []
    lines = [line.strip() for line in text_input.strip().split("\n") if line.strip()]

    for line in lines:
        # Check if line contains a description in quotes
        desc_match = re.search(r"['\"]([^'\"]+)['\"]", line)
        custom_desc = desc_match.group(1) if desc_match else ""
        cleaned_line = re.sub(r"['\"][^'\"]+['\"]", "", line)

        nums = [float(n) for n in re.findall(r"[-+]?(?:\d*\.\d+|\d+)", cleaned_line)]
        norm_desc = normalize_board_description(custom_desc)

        if len(nums) >= 4:
            w, h, q, r = nums[0], nums[1], nums[2], nums[3]
        elif len(nums) == 3:
            w, h, q, r = nums[0], nums[1], 1, nums[2]
        elif len(nums) == 2:
            w, h, q, r = nums[0], nums[1], 1, get_rate_for_description(norm_desc)
        else:
            continue

        boards.append({
            "w": w,
            "h": h,
            "q": q,
            "r": r,
            "description": norm_desc
        })

    return boards


def process_bill_data(
    bill_no: str,
    pr_no: str,
    brand_input: str,
    showroom: str,
    address: str,
    mobile: str,
    raw_specs: Optional[str] = None,
    structured_boards: Optional[List[Union[BoardItem, dict]]] = None,
    bill_date: Optional[str] = None,
    custom_words: Optional[str] = None,
):
    """
    Core business logic that calculates amounts, VAT, TDS, grand total,
    words representation, and returns both calculation summary and docx context.
    """
    # 1. Resolve boards from structured list or text specifications
    boards_arr = []

    # Prioritize structured boards if provided and non-empty
    if structured_boards and len(structured_boards) > 0:
        for b in structured_boards:
            if isinstance(b, BoardItem):
                b_dict = b.model_dump()
            elif isinstance(b, dict):
                b_dict = b
            else:
                continue

            try:
                w = float(b_dict.get("w", 0))
                h = float(b_dict.get("h", 0))
                q_val = b_dict.get("q", 1)
            except (ValueError, TypeError):
                continue

            if w <= 0 or h <= 0:
                continue

            desc = normalize_board_description(b_dict.get("description"), b_dict.get("r"))

            r_val = b_dict.get("r")
            try:
                if r_val is None or str(r_val).strip() == "" or float(r_val) <= 0:
                    r = get_rate_for_description(desc)
                else:
                    r = float(r_val)
            except (ValueError, TypeError):
                r = get_rate_for_description(desc)

            q = compute_qty(q_val)
            sft = compute_sft(w, h, q_val)
            amnt = round(sft * r)

            boards_arr.append({
                "w": w,
                "h": h,
                "q": q,
                "r": r,
                "sft": sft,
                "amnt": amnt,
                "amn": f"{amnt:,}",
                "description": desc,
            })

    # If no boards found from structured list, parse raw specifications text
    if not boards_arr and raw_specs:
        parsed = parse_boards_specifications(raw_specs)
        for p in parsed:
            w = p["w"]
            h = p["h"]
            q_val = p["q"]
            desc = p.get("description", BOARD_TYPE_HORIZONTAL)
            r_val = p.get("r")
            try:
                if r_val is None or str(r_val).strip() == "" or float(r_val) <= 0:
                    r = get_rate_for_description(desc)
                else:
                    r = float(r_val)
            except (ValueError, TypeError):
                r = get_rate_for_description(desc)

            q = compute_qty(q_val)
            sft = compute_sft(w, h, q_val)
            amnt = round(sft * r)

            boards_arr.append({
                "w": w,
                "h": h,
                "q": q,
                "r": r,
                "sft": sft,
                "amnt": amnt,
                "amn": f"{amnt:,}",
                "description": desc,
            })

    # Fallback default board if empty
    if not boards_arr:
        default_w, default_h, default_q = 10.0, 4.0, 1
        desc = BOARD_TYPE_HORIZONTAL
        default_r = DEFAULT_BOARD_RATES[BOARD_TYPE_HORIZONTAL]
        sft = compute_sft(default_w, default_h, default_q)
        amnt = round(sft * default_r)
        boards_arr.append({
            "w": default_w,
            "h": default_h,
            "q": default_q,
            "r": default_r,
            "sft": sft,
            "amnt": amnt,
            "amn": f"{amnt:,}",
            "description": desc,
        })

    # 2. Financial Calculations
    sub1 = sum(item["amnt"] for item in boards_arr)
    vat = round(sub1 * 0.15)
    sub2 = sub1 + vat
    tds = round(sub1 * 0.05)
    total = sub1 - tds
    words = custom_words.strip() if (custom_words and custom_words.strip()) else number_to_words(total)

    # 3. Brand details
    brand_name, brand_code = resolve_brand(brand_input)

    # 4. Formatted Date
    if bill_date and bill_date.strip():
        # Handle ISO date YYYY-MM-DD or existing DD.MM.YYYY
        try:
            if "-" in bill_date:
                parsed_d = datetime.strptime(bill_date.strip(), "%Y-%m-%d")
                formatted_date = parsed_d.strftime("%d.%m.%Y")
            else:
                formatted_date = bill_date.strip()
        except Exception:
            formatted_date = date.today().strftime("%d.%m.%Y")
    else:
        formatted_date = date.today().strftime("%d.%m.%Y")

    # 5. Template context
    context = {
        "bno": bill_no.strip(),
        "PR": pr_no.strip(),
        "bn": brand_name,
        "brand": brand_code,
        "b": boards_arr,
        "show": showroom.strip(),
        "addr": address.strip(),
        "sub1": f"{sub1:,}",
        "vat": f"{vat:,}",
        "sub2": f"{sub2:,}",
        "tds": f"{tds:,}",
        "total": f"{total:,}",
        "words": words,
        "mob": mobile.strip(),
        "date": formatted_date,
    }

    summary = {
        "bill_no": bill_no.strip(),
        "pr_no": pr_no.strip(),
        "brand_name": brand_name,
        "brand_code": brand_code,
        "showroom": showroom.strip(),
        "address": address.strip(),
        "mobile": mobile.strip(),
        "date": formatted_date,
        "total_boards": len(boards_arr),
        "total_sft": round(sum(b["sft"] for b in boards_arr), 2),
        "sub1": sub1,
        "sub1_formatted": f"{sub1:,}",
        "vat": vat,
        "vat_formatted": f"{vat:,}",
        "sub2": sub2,
        "sub2_formatted": f"{sub2:,}",
        "tds": tds,
        "tds_formatted": f"{tds:,}",
        "total": total,
        "total_formatted": f"{total:,}",
        "words": words,
        "boards": boards_arr,
    }

    return context, summary


def generate_docx_file(context: dict, bill_no: str, showroom: str) -> str:
    """Renders the docx template and returns the absolute saved file path."""
    if not TEMPLATE_DOCX.exists():
        raise FileNotFoundError(f"Template '{TEMPLATE_DOCX}' not found!")

    doc = DocxTemplate(str(TEMPLATE_DOCX))
    doc.render(context)

    # Sanitize filename for Windows & Web
    clean_bno = re.sub(r'[\\/*?:"<>|]', "_", bill_no).strip()
    clean_show = re.sub(r'[\\/*?:"<>|]', "_", showroom).strip()
    filename = f"{clean_bno}_{clean_show}.docx" if (clean_bno or clean_show) else "Bill.docx"

    file_path = OUTPUT_DIR / filename
    doc.save(str(file_path))
    return filename


# -------------------------------------------------------------
# Routes
# -------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    """Render the modern web form."""
    return templates.TemplateResponse(
        request=request,
        name="index.html",
        context={
            "today_date": date.today().strftime("%Y-%m-%d"),
            "formatted_today": date.today().strftime("%d.%m.%Y"),
        },
    )


@app.post("/api/calculate")
async def api_calculate(payload: BillPayload):
    """API endpoint to calculate bill breakdown in real-time."""
    try:
        _, summary = process_bill_data(
            bill_no=payload.bill_no or "DRAFT",
            pr_no=payload.pr_no or "DRAFT",
            brand_input=payload.brand or "p",
            showroom=payload.showroom or "Showroom",
            address=payload.address or "Address",
            mobile=payload.mobile or "01XXXXXXXXX",
            raw_specs=payload.boards_specifications,
            structured_boards=payload.boards,
            bill_date=payload.date,
            custom_words=payload.words,
        )
        return {"success": True, "summary": summary}
    except Exception as e:
        return JSONResponse(status_code=400, content={"success": False, "error": str(e)})


@app.post("/api/scan")
async def api_scan(
    request: Request,
    file: Optional[UploadFile] = File(None),
    use_sample: Optional[bool] = Form(False),
):
    """
    Scan workorder image using Google GenAI Gemini vision model.
    Accepts multipart/form-data upload or JSON payload with use_sample=true or image_base64.
    Extracts showroom, PR number, brand, date, and board specifications with auto-rates.
    """
    try:
        image_bytes: Optional[bytes] = None
        mime_type: str = "image/jpeg"

        # Check if JSON payload was sent
        content_type = request.headers.get("content-type", "")
        if "application/json" in content_type:
            body = await request.json()
            if body.get("use_sample"):
                sample_file = BASE_DIR / "bill.jpeg"
                if not sample_file.exists():
                    raise HTTPException(status_code=404, detail="Sample bill.jpeg not found on server")
                with open(sample_file, "rb") as f:
                    image_bytes = f.read()
                mime_type = "image/jpeg"
            elif body.get("image_base64"):
                import base64
                b64_str = body["image_base64"]
                if "," in b64_str:
                    header, b64_data = b64_str.split(",", 1)
                    if "image/png" in header:
                        mime_type = "image/png"
                    elif "image/webp" in header:
                        mime_type = "image/webp"
                else:
                    b64_data = b64_str
                image_bytes = base64.b64decode(b64_data)

        # Handle multipart file upload or use_sample form flag
        if image_bytes is None:
            if use_sample:
                sample_file = BASE_DIR / "bill.jpeg"
                if not sample_file.exists():
                    raise HTTPException(status_code=404, detail="Sample bill.jpeg not found on server")
                with open(sample_file, "rb") as f:
                    image_bytes = f.read()
                mime_type = "image/jpeg"
            elif file is not None:
                image_bytes = await file.read()
                mime_type = file.content_type or "image/jpeg"

        if not image_bytes:
            raise HTTPException(status_code=400, detail="No workorder image or sample flag provided.")

        # Run vision extraction via Gemini API
        extracted = scan_workorder_image(image_bytes=image_bytes, mime_type=mime_type)

        # Compute totals using our bill calculation engine
        _, summary = process_bill_data(
            bill_no=extracted.get("bill_no") or "DRAFT",
            pr_no=extracted.get("pr_no") or "",
            brand_input=extracted.get("brand") or "c",
            showroom=extracted.get("showroom") or "",
            address=extracted.get("address") or "",
            mobile=extracted.get("mobile") or "",
            structured_boards=extracted.get("boards"),
            bill_date=extracted.get("date"),
        )

        return {
            "success": True,
            "data": extracted,
            "summary": summary,
            "message": f"Successfully extracted workorder data: {len(extracted.get('boards', []))} boards found.",
        }
    except HTTPException:
        raise
    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "error": str(e)})


@app.post("/api/scan-and-stack")
async def api_scan_and_stack(
    request: Request,
    file: Optional[UploadFile] = File(None),
    use_sample: Optional[bool] = Form(False),
    project_id: str = Form("default-project"),
    bill_no_override: Optional[str] = Form(None),
    default_start_bill: Optional[int] = Form(None),
):
    """
    Scan a single workorder photo with Gemini Vision, generate its bill DOCX,
    and directly stack it into the active project.
    """
    try:
        image_bytes: Optional[bytes] = None
        mime_type: str = "image/jpeg"

        content_type = request.headers.get("content-type", "")
        if "application/json" in content_type:
            body = await request.json()
            if body.get("use_sample"):
                sample_file = BASE_DIR / "bill.jpeg"
                if not sample_file.exists():
                    raise HTTPException(status_code=404, detail="Sample bill.jpeg not found")
                with open(sample_file, "rb") as f:
                    image_bytes = f.read()
                mime_type = "image/jpeg"
                project_id = body.get("project_id") or project_id
                bill_no_override = body.get("bill_no_override") or bill_no_override
            elif body.get("image_base64"):
                import base64
                b64_str = body["image_base64"]
                if "," in b64_str:
                    header, b64_data = b64_str.split(",", 1)
                    if "image/png" in header:
                        mime_type = "image/png"
                    elif "image/webp" in header:
                        mime_type = "image/webp"
                else:
                    b64_data = b64_str
                image_bytes = base64.b64decode(b64_data)
                project_id = body.get("project_id") or project_id
                bill_no_override = body.get("bill_no_override") or bill_no_override

        if image_bytes is None:
            if use_sample:
                sample_file = BASE_DIR / "bill.jpeg"
                if not sample_file.exists():
                    raise HTTPException(status_code=404, detail="Sample bill.jpeg not found")
                with open(sample_file, "rb") as f:
                    image_bytes = f.read()
                mime_type = "image/jpeg"
            elif file is not None:
                image_bytes = await file.read()
                mime_type = file.content_type or "image/jpeg"

        if not image_bytes:
            raise HTTPException(status_code=400, detail="No workorder image provided.")

        # 1. Vision Extraction
        extracted = scan_workorder_image(image_bytes=image_bytes, mime_type=mime_type)

        # 2. Determine Bill Number
        bno = (bill_no_override or "").strip()
        if not bno:
            bno = (extracted.get("bill_no") or "").strip()
        if not bno:
            bno = pm.get_next_bill_no(project_id, default_start=default_start_bill or 1001)

        showroom = (extracted.get("showroom") or "").strip()
        if not showroom:
            showroom = f"Showroom {bno}"

        # 3. Process & Compute Financial Totals
        context, summary = process_bill_data(
            bill_no=bno,
            pr_no=extracted.get("pr_no") or "",
            brand_input=extracted.get("brand") or "c",
            showroom=showroom,
            address=extracted.get("address") or "",
            mobile=extracted.get("mobile") or "",
            structured_boards=extracted.get("boards"),
            bill_date=extracted.get("date"),
        )

        # 4. Generate DOCX file
        filename = generate_docx_file(context, bno, showroom)

        # 5. Add to project stack
        project_data = pm.add_bill_to_project(
            project_id=project_id,
            bill_data={
                "bill_no": bno,
                "showroom": showroom,
                "pr_no": extracted.get("pr_no") or "",
                "brand": extracted.get("brand") or "c",
                "amount": summary["total"],
                "date": extracted.get("date") or date.today().strftime("%d.%m.%Y"),
                "boards_count": len(summary.get("boards", [])),
            },
            docx_file_path=OUTPUT_DIR / filename,
        )

        return {
            "success": True,
            "bill_no": bno,
            "showroom": showroom,
            "pr_no": extracted.get("pr_no"),
            "brand": extracted.get("brand"),
            "amount": summary["total"],
            "filename": filename,
            "boards_count": len(summary.get("boards", [])),
            "project": project_data,
            "data": extracted,
        }
    except HTTPException:
        raise
    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "error": str(e)})


@app.post("/api/scan-batch")
async def api_scan_batch(
    files: List[UploadFile] = File(...),
    project_id: str = Form("default-project"),
    starting_bill_no: Optional[int] = Form(None),
):
    """
    Batch scan multiple workorder images all at once.
    Processes each image sequentially, generates DOCX, and stacks each bill into the project.
    """
    results = []
    current_bill_no = starting_bill_no

    for file in files:
        try:
            image_bytes = await file.read()
            mime_type = file.content_type or "image/jpeg"
            extracted = scan_workorder_image(image_bytes=image_bytes, mime_type=mime_type)

            bno = (extracted.get("bill_no") or "").strip()
            if not bno:
                if current_bill_no is not None:
                    bno = str(current_bill_no)
                    current_bill_no += 1
                else:
                    bno = pm.get_next_bill_no(project_id)
                    try:
                        current_bill_no = int(bno) + 1
                    except ValueError:
                        pass
            else:
                try:
                    current_bill_no = int(bno) + 1
                except ValueError:
                    pass

            showroom = (extracted.get("showroom") or "").strip()
            if not showroom:
                showroom = f"Showroom {bno}"

            context, summary = process_bill_data(
                bill_no=bno,
                pr_no=extracted.get("pr_no") or "",
                brand_input=extracted.get("brand") or "c",
                showroom=showroom,
                address=extracted.get("address") or "",
                mobile=extracted.get("mobile") or "",
                structured_boards=extracted.get("boards"),
                bill_date=extracted.get("date"),
            )

            filename = generate_docx_file(context, bno, showroom)

            pm.add_bill_to_project(
                project_id=project_id,
                bill_data={
                    "bill_no": bno,
                    "showroom": showroom,
                    "pr_no": extracted.get("pr_no") or "",
                    "brand": extracted.get("brand") or "c",
                    "amount": summary["total"],
                    "date": extracted.get("date") or date.today().strftime("%d.%m.%Y"),
                    "boards_count": len(summary.get("boards", [])),
                },
                docx_file_path=OUTPUT_DIR / filename,
            )

            results.append({
                "success": True,
                "file_name": file.filename,
                "bill_no": bno,
                "showroom": showroom,
                "pr_no": extracted.get("pr_no"),
                "amount": summary["total"],
                "boards_count": len(summary.get("boards", [])),
            })
        except Exception as e:
            results.append({
                "success": False,
                "file_name": file.filename,
                "error": str(e),
            })

    updated_proj = pm.get_project(project_id)
    return {
        "success": True,
        "processed_count": len(results),
        "results": results,
        "project": updated_proj,
    }


@app.post("/api/generate")
async def api_generate(payload: BillPayload):
    """
    API endpoint to generate the bill docx and return JSON with summary
    and direct download URL.
    """
    try:
        if not payload.bill_no.strip():
            raise HTTPException(status_code=422, detail="Bill No is required")
        if not payload.showroom.strip():
            raise HTTPException(status_code=422, detail="Showroom name is required")

        context, summary = process_bill_data(
            bill_no=payload.bill_no,
            pr_no=payload.pr_no,
            brand_input=payload.brand,
            showroom=payload.showroom,
            address=payload.address,
            mobile=payload.mobile,
            raw_specs=payload.boards_specifications,
            structured_boards=payload.boards,
            bill_date=payload.date,
            custom_words=payload.words,
        )

        filename = generate_docx_file(context, payload.bill_no, payload.showroom)

        # Automatically add bill to project stack if project_id is provided
        project_data = None
        if payload.project_id:
            try:
                project_data = pm.add_bill_to_project(
                    project_id=payload.project_id,
                    bill_data={
                        "bill_no": payload.bill_no,
                        "showroom": payload.showroom,
                        "pr_no": payload.pr_no,
                        "brand": payload.brand,
                        "amount": summary["total"],
                        "date": payload.date or date.today().strftime("%d.%m.%Y"),
                        "boards_count": len(summary.get("boards", [])),
                    },
                    docx_file_path=OUTPUT_DIR / filename,
                )
            except Exception as pe:
                print(f"[WARN] Error adding bill to project: {pe}")

        return {
            "success": True,
            "filename": filename,
            "download_url": f"/download/{filename}",
            "summary": summary,
            "project": project_data,
        }
    except HTTPException:
        raise
    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "error": str(e)})


@app.post("/generate")
async def form_generate(
    bill_no: str = Form(...),
    pr_no: str = Form(...),
    brand: str = Form(...),
    showroom: str = Form(...),
    address: str = Form(...),
    mobile: str = Form(...),
    boards_specifications: Optional[str] = Form(None),
    boards_json: Optional[str] = Form(None),
    date_val: Optional[str] = Form(None),
):
    """
    Standard HTML Form POST submission endpoint.
    Generates the DOCX and serves it directly as an attachment.
    """
    try:
        structured_boards = None
        if boards_json:
            try:
                parsed_json = json.loads(boards_json)
                if isinstance(parsed_json, list):
                    structured_boards = parsed_json
            except Exception:
                pass

        context, _ = process_bill_data(
            bill_no=bill_no,
            pr_no=pr_no,
            brand_input=brand,
            showroom=showroom,
            address=address,
            mobile=mobile,
            raw_specs=boards_specifications,
            structured_boards=structured_boards,
            bill_date=date_val,
        )

        filename = generate_docx_file(context, bill_no, showroom)
        file_path = OUTPUT_DIR / filename

        return FileResponse(
            path=str(file_path),
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/download/{filename}")
async def download_file(filename: str):
    """Download a generated DOCX file."""
    # Prevent directory traversal
    clean_filename = Path(filename).name
    file_path = OUTPUT_DIR / clean_filename

    if not file_path.exists():
        raise HTTPException(status_code=404, detail="File not found or expired.")

    return FileResponse(
        path=str(file_path),
        filename=clean_filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


# -------------------------------------------------------------
# Project Stack, Top-Sheet & DOCX Merge Endpoints
# -------------------------------------------------------------

@app.get("/api/projects")
async def api_list_projects():
    """List all projects with bill counts and total amounts."""
    return {"success": True, "projects": pm.list_projects()}


@app.post("/api/projects")
async def api_create_project(payload: ProjectCreatePayload):
    """Create a new project."""
    try:
        project = pm.create_project(payload.name)
        return {"success": True, "project": project}
    except Exception as e:
        return JSONResponse(status_code=400, content={"success": False, "error": str(e)})


@app.get("/api/projects/{project_id}")
async def api_get_project(project_id: str):
    """Get project details and list of bills in its stack."""
    try:
        project = pm.get_project(project_id)
        return {"success": True, "project": project}
    except FileNotFoundError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "error": str(e)})


@app.delete("/api/projects/{project_id}/bills/{bill_no}")
async def api_remove_bill(project_id: str, bill_no: str):
    """Remove a bill from project stack."""
    try:
        project = pm.remove_bill_from_project(project_id, bill_no)
        return {"success": True, "project": project}
    except Exception as e:
        return JSONResponse(status_code=400, content={"success": False, "error": str(e)})


@app.post("/api/projects/{project_id}/clear")
async def api_clear_project(project_id: str):
    """Clear all bills from project stack."""
    try:
        project = pm.clear_project_bills(project_id)
        return {"success": True, "project": project}
    except Exception as e:
        return JSONResponse(status_code=400, content={"success": False, "error": str(e)})


@app.get("/api/projects/{project_id}/topsheet")
async def api_generate_topsheet(project_id: str):
    """
    Generate professional Excel Top-Sheet via topsheet.py (create_excel)
    and download it.
    """
    try:
        output_file = pm.generate_topsheet_for_project(project_id)
        return FileResponse(
            path=str(output_file),
            filename=output_file.name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to generate top-sheet: {str(e)}")


@app.get("/api/projects/{project_id}/merged")
async def api_merge_docx(project_id: str):
    """
    Merge all project bill DOCX files into a single master document via merge.py
    (docxcompose) and download it.
    """
    try:
        output_file = pm.merge_docx_for_project(project_id)
        return FileResponse(
            path=str(output_file),
            filename=output_file.name,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except (ValueError, FileNotFoundError) as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to merge DOCX files: {str(e)}")


@app.get("/api/projects/{project_id}/zip")
async def api_export_zip(project_id: str):
    """
    Download complete project package (all individual DOCX bills + merged DOCX + Excel top-sheet).
    """
    try:
        output_file = pm.export_project_zip(project_id)
        return FileResponse(
            path=str(output_file),
            filename=output_file.name,
            media_type="application/zip",
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to create ZIP package: {str(e)}")


@app.post("/api/generate-topsheet-from-docx")
async def api_generate_topsheet_from_docx(
    file: UploadFile = File(...),
    preview: bool = Query(default=False),
):
    """
    Accepts an uploaded merged DOCX file (or single DOCX bill),
    extracts all invoice records (SL#, Company Name, Work Type, District,
    Outlet Name + Outlet Address, PR/TN, Bill Ref No., Total Amount),
    and generates the official 8-column Excel Top-Sheet matching the client format.
    """
    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Invalid file type. Please upload a .docx file.")

    try:
        content = await file.read()
        if not content:
            raise HTTPException(status_code=400, detail="Uploaded file is empty.")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        clean_stem = re.sub(r"[^\w\s-]", "", Path(file.filename).stem).strip().replace(" ", "_") or "Merged"
        saved_docx_path = OUTPUT_DIR / f"{clean_stem}_{timestamp}.docx"
        with open(saved_docx_path, "wb") as f:
            f.write(content)

        output_excel_name = f"{clean_stem}_{timestamp}_TopSheet.xlsx"
        output_excel_path = OUTPUT_DIR / output_excel_name

        records, final_excel_path = generate_official_topsheet_from_docx(str(saved_docx_path), output_excel_path)

        if not records:
            raise HTTPException(
                status_code=400,
                detail="Could not detect any bills or invoice tables in the uploaded document. Ensure it contains Walton bills.",
            )

        total_amount = sum(r.get("total_amount", 0) for r in records)

        if preview:
            return JSONResponse(
                content={
                    "success": True,
                    "filename": output_excel_name,
                    "download_url": f"/api/download-topsheet?file={output_excel_name}",
                    "total_bills": len(records),
                    "total_amount": total_amount,
                    "total_amount_formatted": f"{total_amount:,}",
                    "bills": records,
                }
            )

        return FileResponse(
            path=str(final_excel_path),
            filename=f"{clean_stem}_TopSheet.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to generate top-sheet from DOCX: {str(e)}")


@app.get("/api/download-topsheet")
async def api_download_topsheet(file: str = Query(...)):
    """Download an existing generated top-sheet Excel file from OUTPUT_DIR."""
    clean_filename = Path(file).name
    file_path = OUTPUT_DIR / clean_filename
    if not file_path.exists() or not file_path.is_file():
        raise HTTPException(status_code=404, detail="Requested file not found.")

    return FileResponse(
        path=str(file_path),
        filename=clean_filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/api/projects/{project_id}/official-topsheet")
async def api_generate_project_official_topsheet(
    project_id: str,
    preview: bool = Query(default=False),
):
    """
    Generate the official 8-column Excel Top Sheet directly from project bills
    or project merged DOCX.
    """
    try:
        project = pm.get_project(project_id)
        pdir = pm.get_project_dir(project_id)
        bills = project.get("bills", [])
        if not bills:
            raise HTTPException(status_code=400, detail="Current project has no bills in its stack.")
        
        output_filename = f"{pm.slugify(project.get('name', 'TopSheet'))}_Official_TopSheet.xlsx"
        output_file = pdir / output_filename
        
        merged_file = pdir / f"{pm.slugify(project.get('name', 'MergedBills'))}_Merged.docx"
        if not merged_file.exists():
            try:
                merged_file = pm.merge_docx_for_project(project_id)
            except Exception:
                merged_file = None

        records = []
        out_path = output_file
        if merged_file and merged_file.exists():
            try:
                records, out_path = generate_official_topsheet_from_docx(str(merged_file), output_file)
            except Exception:
                records = []

        if not records:
            out_path = pm.generate_topsheet_for_project(project_id)
            records = [
                {
                    "sl": idx,
                    "company_name": clean_company_name(b.get("brand", "")),
                    "work_type": "signboard",
                    "district": find_district(b.get("showroom", "")),
                    "outlet_address": b.get("showroom", ""),
                    "pr_tn": b.get("pr_no", ""),
                    "bill_no": str(b.get("bill_no", "")),
                    "total_amount": int(round(float(b.get("amount", 0)))),
                }
                for idx, b in enumerate(bills, start=1)
            ]

        total_amount = sum(r.get("total_amount", 0) for r in records)

        if preview:
            return JSONResponse(
                content={
                    "success": True,
                    "filename": output_filename,
                    "download_url": f"/api/projects/{project_id}/official-topsheet",
                    "total_bills": len(records),
                    "total_amount": total_amount,
                    "total_amount_formatted": f"{total_amount:,}",
                    "bills": records,
                }
            )
        
        return FileResponse(
            path=str(out_path),
            filename=output_filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except HTTPException:
        raise
    except (ValueError, FileNotFoundError) as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to generate official project top-sheet: {str(e)}")


# -------------------------------------------------------------
# Main Runner
# -------------------------------------------------------------
if __name__ == "__main__":
    import uvicorn

    print("🚀 Starting Biller FastAPI server at http://127.0.0.1:8000")
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)