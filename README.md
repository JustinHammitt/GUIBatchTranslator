# GUIBatchTranslator

Offline **batch document translator** with a simple PyQt5 GUI.  
Powered by **Argos Translate** (`argostranslate` + `ctranslate2` + `sentencepiece`) and **argos-translate-files**.

> **Note**  
> • **Language models are not included** in the repo to keep it lean.  
> • **Built executables are not checked in** (build locally or via CI).  
> • Main app file: **`GUIBatchTranslate.py`**.

---

## ✨ Features

- **Offline** translations (no network calls).
- **Batch** translate many files and folders at once.
- **Formats**
  - Via `argos-translate-files`: `docx`, `odt`, `pptx`, `odp`, `epub`, `html`, `htm`, `srt`, `pdf` (text-based), `txt`.
  - **Excel native**:
    - `.xlsx` — translates **string cells only**; preserves formulas, numbers, dates, styles.
    - `.xls` — read legacy file and produce `*_translated.xlsx` (values preserved; legacy formatting/styles/formulas not).
- **Pivot via English** (EN) if a direct pair isn’t installed (e.g., ES ↔ JA).
- Optional **first-run model install** when models are placed in a local `models/` folder.

> ⚠️ **Scanned PDFs** require OCR first (e.g., Tesseract/OCRmyPDF). Text PDFs work.

---

## 🖥 Requirements

- **Windows 10/11 (x64)**
- **Python 3.11 (python.org installer)**  
  Using 3.12/3.13 often forces native builds of deps; stick to **3.11** for stable wheels.

---

## 🚀 Quick Start (Developers)

```powershell
# 1) Create & activate a 3.11 venv
py -3.11 -m venv .venv
. .\.venv\Scripts\Activate.ps1

# 2) Install dependencies
python -m pip install --upgrade pip setuptools wheel
python -m pip install PyQt5 argostranslate==1.9.6 argos-translate-files `
  "sentencepiece==0.2.0" "ctranslate2>=4,<5" `
  openpyxl "xlrd==1.2.0" et_xmlfile

# 3) Run the app
python GUIBatchTranslate.py
