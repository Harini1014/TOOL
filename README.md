# Publishing QA Validation Tool

A full-stack QA tool to compare a Word source document against a final PDF output.
**No API key required** — all analysis runs locally using Python libraries.

---

## Tech Stack

| Layer    | Technology                                  |
|----------|---------------------------------------------|
| Frontend | React 18 + Vite + Tailwind CSS + AG Grid    |
| Backend  | Python FastAPI + PyMuPDF + python-docx      |
| PDF      | PyMuPDF (fitz) — text extraction            |
| Word     | python-docx — paragraph/heading/table parse |

---

## Project Structure

```
qc-tool/
├── backend/
│   ├── main.py           ← FastAPI server (all validation logic)
│   └── requirements.txt  ← Python dependencies
└── frontend/
    ├── index.html
    ├── package.json
    ├── vite.config.js
    ├── tailwind.config.js
    └── src/
        ├── App.jsx
        ├── main.jsx
        ├── index.css
        ├── components/
        │   ├── UploadZone.jsx
        │   └── CheckSelector.jsx
        └── pages/
            ├── ValidatePage.jsx   ← Upload + run
            └── ReportPage.jsx     ← AG Grid error table
```

---

## Prerequisites

- **Python 3.9+** — https://python.org
- **Node.js 18+** — https://nodejs.org

---

## Setup & Run

### Step 1 — Backend

```bash
cd backend
pip install -r requirements.txt
python main.py
```

Backend runs at: https://tool-2-3w1t.onrender.com/

### Step 2 — Frontend (new terminal)

```bash
cd frontend
npm install
npm run dev
```

Frontend runs at: http://localhost:5173

### Step 3 — Open in browser

Go to **http://localhost:5173**

---

## How to Use

1. Upload your **Word source** (.docx)
2. Upload your **final PDF** (.pdf — must be text-extractable, not scanned)
3. Select the checks you want to run (all 29 selected by default)
4. Click **Run QA Validation**
5. View the error report — sortable, filterable AG Grid table
6. Export as **CSV** or **TXT**

---

## Validation Checks (29 total)

- Page Number Sequence & Folio
- Running Head Style & Position
- Slug Line Page Range & File Name
- Word-to-Word Comparison
- Typos
- Missing Content
- Content Order
- Heading Levels & Numbering
- Mini TOC
- Equations
- Special Characters & Symbols
- Footnote Citation & Placement
- List Spacing
- Double Digit Alignment
- Facing Page Alignment
- Global Instructions
- Quotations
- Citations & Placement
- Image / Figure Cutoffs
- Image Size
- Line Art Readability
- Credit Lines
- Heading Style Consistency
- Box Style Consistency
- Table Style Consistency
- Font Consistency
- Key Term Page Numbers
- FPO / Placeholder Images
- Unwanted Characters

---

## Notes

- PDFs must be **text-based** (exported from InDesign/Word), not scanned images
- For very large documents, only the first ~7,000 characters are sampled per document section
- The backend runs fully offline — no internet connection or API key needed
