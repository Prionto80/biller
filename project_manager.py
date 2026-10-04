import os
import re
import json
import shutil
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, List, Optional

from topsheet import create_excel
from merge import merge_folder
from docx import Document
from docxcompose.composer import Composer

BASE_DIR = Path(__file__).resolve().parent
PROJECTS_DIR = BASE_DIR / "projects"
PROJECTS_DIR.mkdir(parents=True, exist_ok=True)

DEFAULT_PROJECT_ID = "default-project"
DEFAULT_PROJECT_NAME = "Default Project"


def slugify(text: str) -> str:
    """Generate filesystem-friendly slug."""
    text = text.strip().lower()
    text = re.sub(r"[^\w\s-]", "", text)
    text = re.sub(r"[\s_-]+", "-", text)
    return text.strip("-") or "project"


def get_project_dir(project_id: str) -> Path:
    """Get project directory path."""
    pdir = PROJECTS_DIR / slugify(project_id)
    pdir.mkdir(parents=True, exist_ok=True)
    return pdir


def get_project_json_path(project_id: str) -> Path:
    return get_project_dir(project_id) / "project.json"


def init_default_project() -> Dict[str, Any]:
    """Ensure a default project exists."""
    p_path = get_project_json_path(DEFAULT_PROJECT_ID)
    if not p_path.exists():
        data = {
            "id": DEFAULT_PROJECT_ID,
            "name": DEFAULT_PROJECT_NAME,
            "created_at": datetime.now().isoformat(),
            "updated_at": datetime.now().isoformat(),
            "bills": [],
        }
        with open(p_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return data
    else:
        with open(p_path, "r", encoding="utf-8") as f:
            return json.load(f)


def list_projects() -> List[Dict[str, Any]]:
    """List all projects with metadata and bill counts."""
    init_default_project()
    projects = []
    for item in PROJECTS_DIR.iterdir():
        if item.is_dir():
            json_file = item / "project.json"
            if json_file.exists():
                try:
                    with open(json_file, "r", encoding="utf-8") as f:
                        data = json.load(f)
                        bills = data.get("bills", [])
                        total_amount = sum(b.get("amount", 0) for b in bills)
                        projects.append({
                            "id": data.get("id", item.name),
                            "name": data.get("name", item.name),
                            "created_at": data.get("created_at"),
                            "updated_at": data.get("updated_at"),
                            "bills_count": len(bills),
                            "total_amount": total_amount,
                        })
                except Exception:
                    pass
    # Sort with default project first, then by updated_at desc
    projects.sort(key=lambda p: (0 if p["id"] == DEFAULT_PROJECT_ID else 1, p.get("updated_at") or ""), reverse=False)
    return projects


def get_project(project_id: str) -> Dict[str, Any]:
    """Retrieve full project details."""
    p_path = get_project_json_path(project_id)
    if not p_path.exists():
        if project_id == DEFAULT_PROJECT_ID:
            return init_default_project()
        raise FileNotFoundError(f"Project '{project_id}' not found.")
    with open(p_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    # Calculate live totals
    bills = data.get("bills", [])
    data["bills_count"] = len(bills)
    data["total_amount"] = sum(b.get("amount", 0) for b in bills)
    return data


def create_project(name: str) -> Dict[str, Any]:
    """Create a new project."""
    clean_name = name.strip() or "New Project"
    project_id = slugify(clean_name)
    
    # Avoid duplicate id collision
    base_id = project_id
    counter = 1
    while (PROJECTS_DIR / project_id).exists():
        project_id = f"{base_id}-{counter}"
        counter += 1

    pdir = get_project_dir(project_id)
    data = {
        "id": project_id,
        "name": clean_name,
        "created_at": datetime.now().isoformat(),
        "updated_at": datetime.now().isoformat(),
        "bills": [],
    }
    with open(pdir / "project.json", "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    return data


def get_next_bill_no(project_id: str, default_start: int = 1001) -> str:
    """Find the next sequential bill number for a project stack."""
    project = get_project(project_id)
    bills = project.get("bills", [])
    int_bills = []
    for b in bills:
        bno = str(b.get("bill_no", "")).strip()
        match = re.search(r"\d+", bno)
        if match:
            try:
                int_bills.append(int(match.group(0)))
            except ValueError:
                pass
    if int_bills:
        return str(max(int_bills) + 1)
    return str(default_start)


def add_bill_to_project(
    project_id: str,
    bill_data: Dict[str, Any],
    docx_file_path: Path,
) -> Dict[str, Any]:
    """
    Save a generated bill docx into the project folder and record its metadata.
    """
    project = get_project(project_id)
    pdir = get_project_dir(project_id)

    # Sanitize destination filename
    dest_filename = docx_file_path.name
    dest_path = pdir / dest_filename
    shutil.copyfile(docx_file_path, dest_path)

    # Normalize bill item record
    bill_no = str(bill_data.get("bill_no", "")).strip()
    showroom = str(bill_data.get("showroom", "")).strip()
    pr_no = str(bill_data.get("pr_no", "")).strip()
    brand = str(bill_data.get("brand", "")).strip()
    amount = int(round(float(bill_data.get("amount", 0))))
    date_val = str(bill_data.get("date", "")).strip()
    boards_count = int(bill_data.get("boards_count", 1))

    new_bill = {
        "bill_no": bill_no,
        "showroom": showroom,
        "pr_no": pr_no,
        "brand": brand,
        "amount": amount,
        "date": date_val,
        "filename": dest_filename,
        "boards_count": boards_count,
        "added_at": datetime.now().isoformat(),
    }

    # Replace existing if same bill_no or append
    bills = project.get("bills", [])
    existing_idx = next((i for i, b in enumerate(bills) if b.get("bill_no") == bill_no), -1)
    if existing_idx >= 0:
        bills[existing_idx] = new_bill
    else:
        bills.append(new_bill)

    project["bills"] = bills
    project["updated_at"] = datetime.now().isoformat()

    with open(get_project_json_path(project_id), "w", encoding="utf-8") as f:
        json.dump(project, f, indent=2, ensure_ascii=False)

    project["bills_count"] = len(bills)
    project["total_amount"] = sum(b.get("amount", 0) for b in bills)
    return project


def remove_bill_from_project(project_id: str, bill_no: str) -> Dict[str, Any]:
    """Remove a bill from project and delete its file."""
    project = get_project(project_id)
    pdir = get_project_dir(project_id)

    bills = project.get("bills", [])
    target = next((b for b in bills if b.get("bill_no") == bill_no), None)
    if target:
        filename = target.get("filename")
        if filename:
            file_to_del = pdir / filename
            if file_to_del.exists():
                file_to_del.unlink()

        bills = [b for b in bills if b.get("bill_no") != bill_no]
        project["bills"] = bills
        project["updated_at"] = datetime.now().isoformat()

        with open(get_project_json_path(project_id), "w", encoding="utf-8") as f:
            json.dump(project, f, indent=2, ensure_ascii=False)

    project["bills_count"] = len(bills)
    project["total_amount"] = sum(b.get("amount", 0) for b in bills)
    return project


def clear_project_bills(project_id: str) -> Dict[str, Any]:
    """Clear all bills from project."""
    project = get_project(project_id)
    pdir = get_project_dir(project_id)

    for b in project.get("bills", []):
        filename = b.get("filename")
        if filename:
            fpath = pdir / filename
            if fpath.exists():
                fpath.unlink()

    project["bills"] = []
    project["updated_at"] = datetime.now().isoformat()

    with open(get_project_json_path(project_id), "w", encoding="utf-8") as f:
        json.dump(project, f, indent=2, ensure_ascii=False)

    project["bills_count"] = 0
    project["total_amount"] = 0
    return project


def generate_topsheet_for_project(project_id: str) -> Path:
    """
    Generate the Excel Top-Sheet for all bills in the project using topsheet.py's create_excel.
    """
    project = get_project(project_id)
    bills = project.get("bills", [])
    if not bills:
        raise ValueError("Project has no bills to generate a top-sheet.")

    invoices = [
        {
            "showroom": b.get("showroom", "Showroom"),
            "bill_no": str(b.get("bill_no", "")),
            "amount": int(round(float(b.get("amount", 0)))),
        }
        for b in bills
    ]

    pdir = get_project_dir(project_id)
    safe_name = slugify(project.get("name", "TopSheet"))
    output_filename = f"{safe_name}_TopSheet.xlsx"
    output_file = pdir / output_filename

    create_excel(invoices, output_file)
    return output_file


def merge_docx_for_project(project_id: str) -> Path:
    """
    Merge all bill DOCX files in the project using docxcompose (as in merge.py).
    """
    project = get_project(project_id)
    bills = project.get("bills", [])
    if not bills:
        raise ValueError("Project has no bills to merge.")

    pdir = get_project_dir(project_id)
    docx_files = []
    for b in bills:
        fn = b.get("filename")
        if fn:
            fp = pdir / fn
            if fp.exists():
                docx_files.append(fp)

    if not docx_files:
        raise FileNotFoundError("No DOCX files found on disk for this project.")

    safe_name = slugify(project.get("name", "MergedBills"))
    output_filename = f"{safe_name}_Merged.docx"
    output_file = pdir / output_filename

    # If only 1 document, copy it
    if len(docx_files) == 1:
        shutil.copyfile(docx_files[0], output_file)
        return output_file

    # First document = master
    master = Document(str(docx_files[0]))
    composer = Composer(master)

    # Append remaining documents with page breaks
    for f in docx_files[1:]:
        doc = Document(str(f))
        master.add_page_break()
        composer.append(doc)

    composer.save(str(output_file))
    return output_file


def export_project_zip(project_id: str) -> Path:
    """
    Create a complete zip archive containing all individual bills,
    the merged DOCX, and the TopSheet Excel file.
    """
    project = get_project(project_id)
    pdir = get_project_dir(project_id)

    # Ensure topsheet and merged docx are generated
    try:
        generate_topsheet_for_project(project_id)
    except Exception:
        pass

    try:
        merge_docx_for_project(project_id)
    except Exception:
        pass

    safe_name = slugify(project.get("name", "Project"))
    zip_filename = f"{safe_name}_Complete_Package.zip"
    zip_file = pdir / zip_filename

    with zipfile.ZipFile(zip_file, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in pdir.iterdir():
            if f.is_file() and f.name != zip_filename and not f.name.endswith(".json"):
                zf.write(f, arcname=f.name)

    return zip_file
