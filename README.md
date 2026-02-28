# NIPT Report Generator

A professional, standalone reporting suite for Non-Invasive Prenatal Screening (NIPS), featuring a side-by-side live preview, bulk processing, and clinical demographic comparison.

## 🚀 Quick Start (Windows)

1. **Install Python**: Ensure you have [Python 3.10+](https://www.python.org/downloads/) installed (check "Add Python to PATH" during installation).
2. **Launch**: Double-click **`launch_nipt.bat`**.
   - *First run*: It will automatically create a virtual environment and install all dependencies.
   - *Subsequent runs*: It will launch the application immediately.

## 🐧 Quick Start (Linux)

1. Ensure Python 3.10+ and `python3-venv` are installed.
2. Run via terminal:
   ```bash
   bash launch_nipt.sh
   ```

## ✨ Key Features

- **Manual Entry**: Professional form with 1-second debounced live PDF preview.
- **Batch Upload**: Process multiple patients via Excel (`.xlsx`). Edit demographics and Z-scores directly in the app.
- **Report Comparison**: Compare Manual (signed) vs Automated (pipeline) reports using clinical demographic text extraction.
- **Draft System**: Save/Load JSON drafts for individual patients or entire batches.
- **Multi-Format Export**: Generate high-fidelity PDFs and editable DOCX reports.
- **Branding Toggle**: Switch between Letterhead (With Logo) and Internal (Without Logo) modes.

## 📂 Folder Structure

- `nipt_report_generator.py`: Main GUI application logic.
- `nipt_template.py`: PDF generation engine (ReportLab).
- `nipt_docx_generator.py`: DOCX generation engine.
- `nipt_assets.py`: Base64 encoded logos and signatures.
- `NIPT_Template.xlsx`: Sample Excel format for batch uploads.

## 🛠 Troubleshooting

If the application fails to start:
1. Delete the `.venv` folder and run `launch_nipt.bat` again.
2. Ensure no other instance of the app is running.
3. Check `run_error.log` (if created) for crash details.

---
© 2026 Anderson Diagnostics & Labs. All rights reserved.
