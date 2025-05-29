# 📄 IWBF PDF Generator

---

## Overview

This web application automates the process of filling two specific PDF forms (`Worksheet-Stages-2C-and-3.pdf` and `Assessment-Form-Stages-2AB.pdf`) using data from an Excel spreadsheet (`Players.xlsx`). It simplifies repetitive tasks by quickly generating multiple personalized PDFs.

---

## How to Use

1.  **Access the Application:** Open the app in your browser:
    [Link to your Streamlit Community Cloud app (add this after deployment)]

2.  **Prepare Your Excel File:**
    * Ensure your `Players.xlsx` file contains the following columns:
        * `number`
        * `proposed-class`
        * `name`
        * `country`
        * `date`
        * `competition`
        * `dob`

3.  **Upload the File:**
    * Click the "Select your Players.xlsx file" button within the app.
    * Choose your Excel file from your computer.

4.  **Generate PDFs:**
    * Click the "Generate Worksheets" button.
    * A progress bar and status messages will show the generation progress.

5.  **Download Forms:**
    * Once complete, click the "Download Generated PDFs (ZIP)" button.
    * A `.zip` file containing all personalized PDFs, organized by sheet from your Excel, will be downloaded.

---

## Technologies

* **Python**
* **Streamlit**
* **pandas**
* **pypdf**

---

## Project Structure

* `app.py`: Main Streamlit application code.
* `Worksheet-Stages-2C-and-3.pdf`: Worksheet PDF template.
* `Assessment-Form-Stages-2AB.pdf`: Assessment Form PDF template.
* `requirements.txt`: Python dependencies.
* `README.md`: This file.
* `.gitignore`: Specifies files/folders to be ignored by Git.

---

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.

---
