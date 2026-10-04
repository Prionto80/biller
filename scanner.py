import os
import re
import json
import logging
from pathlib import Path
from typing import Dict, Any, List, Optional
from dotenv import load_dotenv

# Ensure .env is loaded from project directory
BASE_DIR = Path(__file__).resolve().parent
ENV_PATH = BASE_DIR / ".env"
if ENV_PATH.exists():
    load_dotenv(dotenv_path=ENV_PATH)
else:
    load_dotenv()

from google import genai
from google.genai import types
from google.genai import errors

logger = logging.getLogger("biller.scanner")

# Default rates matching project requirements
STANDARD_BOARD_RATES = {
    "New Signboard (Horizontal)": 365.0,
    "New Signboard (Vertical)": 270.0,
    "Pana Change and Servicing": 88.0,
}

SYSTEM_EXTRACTION_PROMPT = """You are an expert OCR and data extraction AI for Walton Group workorders, signboard inspection forms, and billing sheets in Bangladesh.
Document Title typically: 'ওয়ালটন গ্রুপ - স্ট্যাটিক এড এন্ড বিউটিফিকেশন বিভাগ - সাইনবোর্ড ইন্সপেকশন ও প্রাপ্তি স্বীকার পত্র'.

Examine the image carefully and extract all information into a structured JSON format:

1. "bill_no": Bill number if written or noted anywhere. If absent or blank, use "".
2. "pr_no": The PR number / পিআর নাম্বার (e.g., 1127656). Always extract exact digits.
3. "brand": Brand code. Map to one of:
   - "p" for Walton Plaza / Hi-Tech Industries (PLC) (PLAZA)
   - "m" for Walton Digi-Tech Mobile (MOB)
   - "h" for Walton Hi-Tech (HAP)
   - "e" for Walton Hi-Tech (EAP)
   - "c" for Walton Distributor Network / Walton Hi-Tech (COM)
   If the brand says "WDN Walton Distributor" or similar, use "c".
4. "brand_name": Full brand name string found on the sheet (e.g. "WDN Walton Distributor").
5. "showroom": Showroom or Distributor name from the 'ডিস্ট্রিবিউটর/প্লাজার নাম এবং ঠিকানা' box (e.g. "Hossain Electronics-Sreenagar"). Clean up prefixes like 'A-'.
6. "address": Address from the same box (e.g. "Sreenagar, Dhaka South"). Remove phone/ref codes from address.
7. "mobile": Phone/Mobile number from the box or sheet (e.g. "01925501500").
8. "date": Workorder date or print date in DD.MM.YYYY format (e.g., if 2/5/2026 or 5-May-26, format as "02.05.2026" or "05.05.2026").
9. "boards": An array of each signboard line item from the table with:
   - "w": Length / Width (দৈর্ঘ্য) in feet as float (e.g., 32, 20, 2, 3.5).
   - "h": Height (উচ্চতা) in feet as float (e.g., 4, 4, 4, 5).
   - "q": Quantity (সংখ্যা) as integer (e.g., 1, 1, 4, 2).
   - "description": MUST be EXACTLY one of these three options:
       - "New Signboard (Horizontal)"
       - "New Signboard (Vertical)"
       - "Pana Change and Servicing"
     Rule for description:
     * If the type is 'HPC-Pana Change & Serv. Horizontal' or 'VPC-Pana Change & Serv. Vertical' or contains 'Pana Change' or 'Servicing', choose "Pana Change and Servicing".
     * If it is a new signboard horizontal, choose "New Signboard (Horizontal)".
     * If it is a new signboard vertical, choose "New Signboard (Vertical)".
   - "rate": If a rate is explicitly stated, use it as a number; otherwise set to 0.

Return ONLY a valid JSON object matching this schema:
{
  "bill_no": string,
  "pr_no": string,
  "brand": string,
  "brand_name": string,
  "showroom": string,
  "address": string,
  "mobile": string,
  "date": string,
  "boards": [
    {
      "w": number,
      "h": number,
      "q": number,
      "description": string,
      "rate": number
    }
  ]
}
"""

CANDIDATE_MODELS = [
    "gemini-3.8-flash",
    "gemini-3.5-flash",
    "gemini-3.1-flash-lite",
    "gemini-3.1-pro-preview",
]


def get_gemini_client() -> genai.Client:
    """Retrieve and initialize the Google GenAI client with GEMINI_API_KEY."""
    raw_key = os.getenv("GEMINI_API_KEY", "")
    api_key = raw_key.strip().rstrip(".")
    if not api_key:
        raise ValueError(
            "GEMINI_API_KEY is not set in .env. Please configure GEMINI_API_KEY in the .env file."
        )
    return genai.Client(api_key=api_key)


def scan_workorder_image(image_bytes: bytes, mime_type: str = "image/jpeg") -> Dict[str, Any]:
    """
    Scan a workorder image using Google GenAI Gemini vision model.
    Extracts showroom, PR number, address, mobile, date, and board specifications.
    Automatically retries with fallback models if the primary model is busy (503).
    """
    client = get_gemini_client()
    image_part = types.Part.from_bytes(data=image_bytes, mime_type=mime_type)

    last_error = None
    response_text = None
    used_model = None

    for model_name in CANDIDATE_MODELS:
        try:
            logger.info("Attempting extraction with model: %s", model_name)
            response = client.models.generate_content(
                model=model_name,
                contents=[SYSTEM_EXTRACTION_PROMPT, image_part],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.1,
                ),
            )
            response_text = response.text
            used_model = model_name
            break
        except (errors.ServerError, errors.APIError) as e:
            logger.warning("Model %s failed with: %s. Retrying next candidate...", model_name, e)
            last_error = e
            continue
        except Exception as e:
            logger.error("Unexpected error with model %s: %s", model_name, e)
            last_error = e
            continue

    if not response_text:
        raise RuntimeError(
            f"Failed to scan workorder using Gemini Vision models. Last error: {last_error}"
        )

    # Parse and normalize JSON
    try:
        data = json.loads(response_text)
    except json.JSONDecodeError:
        # Fallback regex extraction if wrapped in markdown
        cleaned = re.sub(r"^```(?:json)?", "", response_text.strip())
        cleaned = re.sub(r"```$", "", cleaned.strip())
        data = json.loads(cleaned)

    # Normalization & cleanup
    normalized = normalize_extracted_data(data)
    normalized["_meta"] = {
        "model_used": used_model,
        "success": True,
    }
    return normalized


def normalize_extracted_data(raw: Dict[str, Any]) -> Dict[str, Any]:
    """Normalize extracted fields and ensure all calculations and descriptions align with business rules."""
    showroom = str(raw.get("showroom", "")).strip()
    # Clean up prefixes like 'A- ' or '1. '
    showroom = re.sub(r"^[A-Z0-9][-–—\s.]+", "", showroom).strip()

    address = str(raw.get("address", "")).strip()
    address = re.sub(r",\s*,", ", ", address).strip()

    mobile = str(raw.get("mobile", "")).strip()
    # Clean mobile number
    mobile_match = re.search(r"01[0-9\s-]{9,13}", mobile)
    if mobile_match:
        mobile = mobile_match.group(0).replace(" ", "")

    pr_no = str(raw.get("pr_no", "")).strip()
    bill_no = str(raw.get("bill_no", "")).strip()
    date_val = str(raw.get("date", "")).strip()
    brand_code = str(raw.get("brand", "c")).strip().lower()
    brand_name = str(raw.get("brand_name", "")).strip()

    if brand_code not in ("p", "m", "h", "e", "c"):
        brand_code = "c"

    # Normalize boards
    raw_boards = raw.get("boards", [])
    normalized_boards = []

    for b in raw_boards:
        try:
            w = float(b.get("w", 0))
            h = float(b.get("h", 0))
            q = int(b.get("q", 1))
            desc = str(b.get("description", "Pana Change and Servicing")).strip()
            
            # Map description strictly to the 3 valid choices
            d_lower = desc.lower()
            if "horizontal" in d_lower and "new" in d_lower:
                final_desc = "New Signboard (Horizontal)"
            elif "vertical" in d_lower and "new" in d_lower:
                final_desc = "New Signboard (Vertical)"
            else:
                final_desc = "Pana Change and Servicing"

            # Determine rate
            rate = float(b.get("rate", 0))
            if rate <= 0:
                rate = STANDARD_BOARD_RATES.get(final_desc, 88.0)

            sft = round(w * h * q, 2)
            amnt = round(sft * rate)

            normalized_boards.append({
                "w": w,
                "h": h,
                "q": q,
                "description": final_desc,
                "desc": final_desc,
                "rate": rate,
                "r": rate,
                "sft": sft,
                "amnt": amnt,
                "isCustomRate": False,
            })
        except Exception as e:
            logger.warning("Error normalizing board item %s: %s", b, e)

    return {
        "bill_no": bill_no,
        "pr_no": pr_no,
        "brand": brand_code,
        "brand_name": brand_name,
        "showroom": showroom,
        "address": address,
        "mobile": mobile,
        "date": date_val,
        "boards": normalized_boards,
    }
