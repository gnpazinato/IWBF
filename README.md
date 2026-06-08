# IWBF Classification Tools

---

## Overview

This web app is a **hub of IWBF classification tools** — a single Streamlit deployment (one URL, https://classificationiwbf.streamlit.app/) where users pick a tool from the sidebar menu. It currently includes:

* **Player Assessment Forms Generator** — fills the IWBF assessment forms (`Worksheet-Stages-2C-and-3.pdf` and `Assessment-Form-Stages-2AB.pdf`) from an Excel spreadsheet (`Players.xlsx`), generating multiple personalized PDF forms at once.
* **Player Card Merger** — merges multiple player card PDFs into a single, print-ready sheet (A4 layout or a business-card template).
* **Final Results Generator** — fills the Classification Results forms (Stage 2 and Final) from a spreadsheet with one tab per team, producing one PDF per team in a ZIP.

New tools can be added without creating new repositories — each tool lives in its own folder under `tools/` (see [Project Structure](#project-structure) below).

---

## How to Use

1.  **Access the Application:** Open the app in your browser:
    https://classificationiwbf.streamlit.app/

2.  **Prepare Your Excel File:**
    * Ensure your `Players.xlsx` file contains the following columns for each player:
        * `number`
        * `proposed-class`
        * `name`
        * `country`
        * `date`
        * `competition`
        * `dob`

3.  **Upload the File:**
    * On the application interface, click the "Select your Players.xlsx file" button.
    * Choose the Excel file from your computer.

4.  **Generate Forms:**
    * After uploading the file, click the "Generate Player Forms" button.
    * A progress bar and status messages will indicate the generation progress.

5.  **Download Forms:**
    * Once the process is complete, a "Click to Download Generated Forms (ZIP)" button will appear.
    * Click it to download a `.zip` file containing all the personalized PDF forms. The forms will be organized into folders named "Stages 2C and 3" and "Stages 2AB" within the ZIP archive.

---

## Technologies

This project is built using:

* **Python**
* **Streamlit:** For the interactive web interface and multipage navigation (`st.navigation`).
* **pandas:** For efficient Excel data reading and manipulation.
* **PyPDF2:** For PDF form filling and manipulation.
* **PyMuPDF (fitz) & reportlab:** For the Player Card Merger's PDF layout and rendering.

---

## Project Structure

This repository is a **multi-tool hub**: a single Streamlit app that bundles several IWBF tools, each in its own folder, wired together with Streamlit's native multipage navigation (`st.navigation` / `st.Page`).

```
IWBF/
├── app.py                      # Entry/router: page config + st.navigation menu
├── requirements.txt            # Combined dependencies for all tools
├── home/
│   └── home.py                 # Landing page (welcome + tool descriptions)
└── tools/
    ├── assessment_forms/
    │   ├── assessment_forms.py # Tool 1 — Player Assessment Forms Generator
    │   └── assets/             # PDF templates + Players.xlsx for this tool
    ├── card_merger/
    │   └── card_merger.py      # Tool 2 — Player Card Merger (upload-only)
    └── final_results/
        ├── final_results.py    # Tool 3 — Final Results Generator
        └── assets/             # 2 PDF templates + template spreadsheet
```

### Adding a new tool
1. Create `tools/<name>/<name>.py` — a normal Streamlit script with **no** `st.set_page_config`. Load any bundled files via `Path(__file__).resolve().parent / "assets"`.
2. Register it in `app.py`: add one `st.Page("tools/<name>/<name>.py", title=..., icon=...)` and include it in the `st.navigation([...])` list.

---

## 📝 License

This project is licensed under the MIT License. Please refer to the `LICENSE` file in the repository for more details.

---
