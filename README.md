```markdown
# Financial Statement Reconciliation Tool

The Financial Statement Reconciliation Tool is a prototype developed as part of a thesis project. It supports the reconciliation of financial statement figures from text-based PDF documents.

The tool compares figures from an older financial statement (current period) with the corresponding comparative period figures in a newer financial statement. It is designed to reduce manual work and highlight items that require further review, not to replace professional judgment.

---

## Features

- Processes text-based PDF financial statements
- Detects income statement and balance sheet sections
- Extracts line items and numeric values
- Normalizes row labels using a predefined term dictionary
- Performs rule-based reconciliation across periods
- Flags uncertain or unmatched items for review
- Generates a structured Excel report
- Includes a Streamlit-based user interface

---

## Limitations

- Designed primarily for Finnish micro-company financial statements
- Supports only text-based PDFs (no scanned PDFs without OCR)
- Does not perform accounting, auditing, or legal conclusions
- Intended as a support tool for professionals

---

## Project Structure

All files must be located in the **same folder** for the tool to work correctly.

---

#project-folder/
#│
#├── FSRT.py            # Main application
#├── terms.xlsx         # Term dictionary (required)
#├── requirements.txt   # Python dependencies
#├── README.md          # Documentation
#└── LICENSE            # License file

---

---

## Installation

1. Clone or download the project.

2. Create a virtual environment:

```bash
python -m venv .venv
````

Activate the environment:

**Windows**

```bash
.venv\Scripts\activate
```

**Mac/Linux**

```bash
source .venv/bin/activate
```

Install dependencies:

```bash
pip install -r requirements.txt
```

---

## Dependencies

```text
pandas==2.2.2
pdfplumber==0.11.0
streamlit==1.35.0
openpyxl==3.1.2
```

---

## Running the Application

Start the application with:

```bash
python -m streamlit run FSRT.py
```

The Streamlit interface will open automatically in your browser.

---

## Usage

1. Upload the **older financial statement PDF** (left side)
2. Upload the **newer financial statement PDF** (right side)
3. Click **"Reconcile Financial Statements"**
4. Review the summary in the interface
5. Download the Excel report for detailed analysis

---

## Excel Report

The generated Excel file includes multiple sheets:

* Summary
* Structured financial statement data
* Reconciliation results
* Items requiring review
* Additional details

### Color coding:

* **Green** = matched
* **Yellow** = requires review
* **Red** = mismatch or missing item
* **Gray** = structural/header row

---

## Notes

* Works best with clean, text-based PDFs
* If documents are scanned or poorly structured, results may be unreliable
* Comparative period data must be present in the financial statements

---

## Author

Elina Malkki

---

## Version

1.6.0

