# IWBF Player Assessment Forms Generator

---

## Overview

This web application is designed to automate the process of filling two specific IWBF player assessment forms: `Worksheet-Stages-2C-and-3.pdf` and `Assessment-Form-Stages-2AB.pdf`. It uses data from an Excel spreadsheet (`Players.xlsx`) to quickly generate multiple personalized PDF forms.

This tool streamlines the manual form-filling process, making it efficient for professionals involved in player assessment.

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
* **Streamlit:** For the interactive web interface.
* **pandas:** For efficient Excel data reading and manipulation.
* **PyPDF2:** For PDF form filling and manipulation.

---

## Project Structure

The repository contains the following essential files:

* `app.py`: The main Streamlit application code.
* `Worksheet-Stages-2C-and-3.pdf`: The template for the Worksheet form.
* `Assessment-Form-Stages-2AB.pdf`: The template for the Assessment form.
* `requirements.txt`: Lists all Python dependencies required for the application.
* `README.md`: This file, providing project overview and usage instructions.
* `.gitignore`: Specifies files and directories to be ignored by Git.

---

## 📝 License

This project is licensed under the MIT License. Please refer to the `LICENSE` file in the repository for more details.

---
