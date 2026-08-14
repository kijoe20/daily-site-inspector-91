# 📸 Daily OCR Reporter

An automated web application built with **Streamlit** that streamlines the photo reporting process for site inspections. The app automatically scans uploaded photo floor/unit labels using OCR, standardizes the file naming, orders the photos by batch hierarchy, and embeds them directly into a downloadable Word report.

---

## ✨ Features

* **Automated Label Recognition:** Uses EasyOCR to detect unit labels (e.g., `36/F C1`, `28/F D2`) directly from photos.
* **Smart File Naming:** Automatically normalizes names and handles duplicate tags gracefully.
* **Batch Hierarchy Sorting:** Automatically orders images according to predefined floor batch specs.
* **One-Click Report Generation:** Outputs a ready-to-use Word table report (`.docx`) with embedded photos and standardized text descriptions.
* **User-Friendly Web Interface:** Simple drag-and-drop interface for non-technical team members.

---

## 🛠️ Project Structure

```text
├── app.py              # Main Streamlit application
├── requirements.txt    # Python dependencies
├── packages.txt        # System-level dependencies (for Streamlit Cloud deployment)
└── README.md           # Project documentation
