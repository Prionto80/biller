import ollama
import json
import sys
import re
from pathlib import Path
from docxtpl import DocxTemplate
from inwords import number_to_words



# ── Step 1: Extract structured data via Claude vision ────────────────────────

EXTRACT_PROMPT = """Extract invoice data from this handwritten note image and return ONLY valid JSON with no markdown, no explanation, no extra text.

The note follows this format (lines may be in any order):
- Bill number (short number, e.g. 3275)
- AD Farm / PR number (longer number, e.g. 1128805)
- Brand code: "w" for Walton Hi-Tech, "m" for Walton Mobile/Digi-Tech
- Showroom name
- Address / area name
- Mobile number (starts with 01)
- One or more signboard rows in format: WIDTH - HEIGHT - QUANTITY - RATE

Return exactly this JSON structure:
{
  "bno": "bill number as string",
  "pr": "PR/farm number as string",
  "brand": "w or m",
  "show": "showroom name",
  "addr": "address",
  "mob": "mobile number",
  "boards": [
    {"w": 15.5, "h": 4.0, "q": 1, "r": 380},
    {"w": 2.5,  "h": 6.0, "q": 2, "r": 280}
  ]
}

Rules:
- w, h, r are floats; q is an integer
- If brand line is missing, default to "w"
- If a field is missing, use an empty string or 0
"""


def extract_from_image(image_path: str) -> dict:
    print("  Sending image to LLaVA (local)...")
    
    with open(image_path, "rb") as f:
        image_data = f.read()

    response = ollama.chat(
        model="llava",
        messages=[{
            "role": "user",
            "content": EXTRACT_PROMPT,
            "images": [image_data],
        }]
    )

    raw = response["message"]["content"].strip()
    raw = re.sub(r"^```(?:json)?", "", raw).strip()
    raw = re.sub(r"```$", "", raw).strip()

    try:
        return json.loads(raw)
    except json.JSONDecodeError as e:
        print(f"\n[ERROR] Could not parse LLaVA response as JSON:\n{raw}\n")
        raise e


# ── Step 2: Compute derived fields (same logic as original generate.py) ───────

def sft(w: float, h: float, q: int) -> float:
    return round(w * h * q, 2)


def build_boards_arr(boards: list[dict]) -> list[dict]:
    result = []
    for b in boards:
        w   = float(b["w"])
        h   = float(b["h"])
        q   = int(b["q"])
        r   = int(b["r"])
        s   = sft(w, h, q)
        amnt = round(s * r)

        if r > 95:
            description = "New Signboard"
        elif r > 50:
            description = "Pana Change and Servicing"
        else:
            description = "Servicing"

        result.append({
            "w":           w,
            "h":           h,
            "q":           q,
            "r":           r,
            "sft":         s,
            "amnt":        amnt,
            "amn":         f"{amnt:,}",
            "description": description,
        })
    return result


def build_context(extracted: dict) -> dict:
    brand = extracted.get("brand", "w").strip().lower()
    brand_name = (
        "Walton Digi-Tech Mobile"
        if brand == "m"
        else "Walton Hi-Tech Industries (PLC)"
    )

    boards_arr = build_boards_arr(extracted.get("boards", []))

    sub1  = sum(b["amnt"] for b in boards_arr)
    vat   = round(sub1 * 0.15)
    sub2  = sub1 + vat
    tds   = round(sub1 * 0.05)
    total = sub1 - tds
    words = number_to_words(total)

    return {
        "bno":   extracted.get("bno", ""),
        "PR":    extracted.get("pr", ""),
        "bn":    brand_name,
        "brand": brand_code,
        "b":     boards_arr,
        "show":  extracted.get("show", ""),
        "addr":  extracted.get("addr", ""),
        "sub1":  f"{sub1:,}",
        "vat":   f"{vat:,}",
        "sub2":  f"{sub2:,}",
        "tds":   f"{tds:,}",
        "total": f"{total:,}",
        "words": words,
        "mob":   extracted.get("mob", ""),
    }


# ── Step 3: Render the docx ───────────────────────────────────────────────────

def generate_invoice(image_path: str, template_path: str = "bill.docx"):
    print(f"\n[1/3] Reading bill photo: {image_path}")
    extracted = extract_from_image(image_path)

    print(f"\n[2/3] Extracted data:")
    print(f"  Bill No   : {extracted.get('bno')}")
    print(f"  PR        : {extracted.get('pr')}")
    print(f"  Brand     : {extracted.get('brand')}")
    print(f"  Showroom  : {extracted.get('show')}")
    print(f"  Address   : {extracted.get('addr')}")
    print(f"  Mobile    : {extracted.get('mob')}")
    for i, b in enumerate(extracted.get("boards", []), 1):
        print(f"  Board {i}   : {b['w']} × {b['h']} ft, qty={b['q']}, rate={b['r']}")

    context = build_context(extracted)

    print(f"\n[3/3] Rendering invoice...")
    doc = DocxTemplate(template_path)
    doc.render(context)

    output_name = f"{context['bno']} {context['show']}.docx"
    doc.save(output_name)
    print(f"\n✓ Invoice saved: {output_name}")
    print(f"  Total payable : ৳{context['total']}")
    print(f"  In words      : {context['words']}")


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python generate.py <bill_photo.jpg> [bill_template.docx]")
        print("Example: python generate.py billphoto.jpeg bill.docx")
        sys.exit(1)

    image_path    = sys.argv[1]
    template_path = sys.argv[2] if len(sys.argv) > 2 else "bill.docx"

    generate_invoice(image_path, template_path)