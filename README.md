# Biller - Modern FastAPI Invoice Studio, AI Batch Scanner & Project Stack

A high-performance FastAPI web application to generate, batch-process, and manage professional Walton `.docx` billing invoices. Includes automated area calculations, VAT & TDS computations, number-to-words translations, arrow-key navigation, a live bidirectional editable A4 DOCX sheet, **Multi-Photo AI Batch Workorder Scanner powered by Google Gemini Vision**, **Project Stacks**, **Excel Top-Sheet Generation (`topsheet.py`)**, and **DOCX Document Merging (`merge.py`)**.

---

## ✨ Features

### 1. 📷 Multi-Photo AI Batch Scanner (Google Gemini Vision)
- **Multi-Upload All at Once**: Drag & drop or multi-select dozens of workorder / signboard inspection photos simultaneously.
- **Sequential Auto-Processing**: Scans each photo sequentially with Gemini Vision AI, extracts metadata, formats the bill, generates the `.docx`, and stacks it into the project automatically.
- **Automatic Data Extraction**:
  - **PR Number** (e.g. `1127656`)
  - **Showroom Name** (e.g. `Hossain Electronics-Sreenagar`)
  - **Address & Mobile Number**
  - **Brand Category** (`(COM)`, `(PLAZA)`, `(MOB)`, `(HAP)`, `(EAP)`)
  - **Signboard Items & Dimensions**: Length, Height, Quantity, Particle Type.
  - **Standardized Rates**: Automatically sets ৳365 for Horizontal, ৳270 for Vertical, ৳88 for Pana Change.
- **Auto-Incrementing Bill Numbers**: Automatically detects bill numbers or increments starting from user input / highest existing bill in the stack.
- **Real-Time Visual Queue**: Live status badges (`Pending` ⏳, `AI Scanning...` 🔍, `Stacked! ✅`) and progress bar tracking.

### 2. 📚 Project Stack & Batch Management
- **Project Groups**: Organize bills under dedicated projects (e.g. `October 2026 Batch`, `Uttara Zone Showrooms`).
- **Clean "Add to Stack" Workflow**: Form submission immediately adds bills to the active project stack without annoying browser file download popups.
- **Live Counter & Slide-Over Drawer**: Visual counter badge (`📚 Stack (5)`) and side drawer to review, inspect, download individual bills, or remove items.

### 3. 📊 Excel Top-Sheet Generator (`topsheet.py`)
- Generates official formatted Walton Excel Top-Sheets (`.xlsx`) using `openpyxl`.
- Automatically populates SL No, Particulars (Showroom), Bill No, Amount, Remarks, formatted borders, and Total Amount formula/value.
- Computes bill range headers (e.g. `Bill No 1042-1048`).

### 4. 📑 DOCX Document Merger (`merge.py`)
- Concatenates all `.docx` invoices in a project stack into a single master Word document using `docxcompose`.
- Maintains page breaks and template formatting across all pages.
- 1-click **Download Full Project ZIP** packaging all individual bills, merged DOCX, and Top-Sheet Excel.

### 5. ✏️ Interactive A4 DOCX Live Sheet & Bidirectional Sync
- Realistic life-sized document sheet that renders an exact visual match of the official Word invoice.
- **Viewport Height Auto-Fitting**: Automatically fits within viewport height without page scrolling.
- **Two-Way Bidirectional Sync ("Both Ways Are Open")**:
  - Edit via the form $\longrightarrow$ Document sheet updates live in real-time.
  - Click directly into any text or table cell on the document sheet $\longrightarrow$ Syncs back to the form and recalculates totals immediately.

### 6. ⌨️ Fast Keyboard Navigation
- Jump between input fields using `↑`, `↓`, `←`, and `→` arrow keys.
- Press `↵ Enter` from any field to immediately add the bill to the project stack.

### 7. 🧮 Automated Tax & Amount Calculations
- Sub-Total 1 (Sum of all board items)
- (+) VAT @ 15%
- Sub-Total 2 (Sub1 + VAT)
- (-) VAT @ 15%
- Sub-Total 1
- (-) TDS @ 5%
- Grand Total (Sub1 - TDS)
- Amount in words (*e.g. Twenty-Two Thousand Nine Hundred Ninety Taka Only*)

---

## 🔑 Environment Configuration

Copy `.env.example` to `.env` in the project root:

```bash
cp .env.example .env
```

Add your Gemini API Key in `.env`:

```env
GEMINI_API_KEY=your_gemini_api_key_here
```

The server automatically loads the key using `python-dotenv` and initializes the official `google-genai` SDK with smart fallback handling across Gemini models (`gemini-3.8-flash`, `gemini-3.5-flash`, etc.).

---

## 🚀 Running the Project

### 1. Run with `uv` (Recommended)

```bash
uv run uvicorn main:app --reload --port 8000
```

### 2. Access the Web App

Open your browser at:
```
http://127.0.0.1:8000
```

---

## 📡 API Endpoints

- `GET /`: Serves the web studio with live document viewer, batch scanner, and project drawer.
- `POST /api/scan`: Single image scan returning extracted fields.
- `POST /api/scan-and-stack`: Scans an image, creates DOCX, and directly stacks it into a project.
- `POST /api/scan-batch`: Batch scans multiple images sequentially and stacks them.
- `POST /api/calculate`: Live calculation breakdown (amounts, VAT, TDS, words).
- `POST /api/generate`: Generates DOCX invoice and adds to project stack.
- `GET /api/projects`: Lists all projects with bill counts and cumulative amounts.
- `POST /api/projects`: Creates a new project.
- `GET /api/projects/{project_id}`: Retrieves project details and bills stack.
- `GET /api/projects/{project_id}/topsheet`: Generates and downloads formatted Excel Top-Sheet via `topsheet.py`.
- `GET /api/projects/{project_id}/merged`: Merges all project DOCX files via `merge.py` and downloads.
- `GET /api/projects/{project_id}/zip`: Downloads full project ZIP archive.
- `DELETE /api/projects/{project_id}/bills/{bill_no}`: Removes a bill from the stack.
- `POST /api/projects/{project_id}/clear`: Clears all bills from the stack.
