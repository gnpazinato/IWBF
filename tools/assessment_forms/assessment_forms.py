import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.generic import NameObject, BooleanObject, DictionaryObject
import io
import re
import zipfile
from pathlib import Path

# Assets bundled with this tool, resolved relative to THIS file (not the CWD),
# so the templates keep loading no matter how the page is launched.
ASSETS = Path(__file__).resolve().parent / "assets"

# --- Helper Functions ---

def format_date(date):
    """
    Formats a date to 'dd-mm-yyyy' or returns the value as a string.
    Handles different input types for dates.
    """
    try:
        return pd.to_datetime(date).strftime("%d-%m-%Y")
    except Exception:
        return str(date)

def format_class(value):
    """
    Formats a sport class as 'X.0'/'X.5' with a dot decimal
    (e.g. 1.0, 1.5, 2.0, 2.5) — never '1,0' nor a bare '1'.
    Non-numeric values are returned trimmed, as-is.
    """
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip().replace(",", ".")
    try:
        return f"{float(text):.1f}"
    except ValueError:
        return text

def is_review_status(status):
    """
    Decide whether a player's classification 'status' is a *review*.

    A review athlete has already been classified and only needs the
    "Worksheet Stages 2C and 3" form — they SKIP the full Stages 2AB
    assessment. A *new* athlete (status N / New) goes through the whole
    process and gets BOTH forms. A blank/missing/unrecognized status
    defaults to NEW (both forms) so nothing is ever silently dropped.

    IWBF/IPC sport-class statuses spell "review" many ways; anything tied
    to a review counts here, e.g.:
        R, Review, R-FRD, RFRD, R-FD, RFD, FRD, R-NAO, "Review (FRD)" ...
    """
    if status is None or pd.isna(status):
        return False
    s = str(status).strip().upper()
    if not s:
        return False
    # An explicit "new" never counts as a review.
    if s in {"N", "NEW"}:
        return False
    # Collapse separators so "R-FRD", "R FD", "R.F.D", "R(FRD)" all look the same.
    compact = re.sub(r"[\s\-_./()]", "", s)
    # Any review marker -> review (only the worksheet is generated):
    #   - starts with "R"        -> R, REVIEW, RFRD, RFD, RNAO, RNAT, ...
    #   - contains "FRD"/"FDR"   -> a Fixed Review Date written without a leading
    #                               R. IWBF's own master-list legend transposes
    #                               the IPC "FRD" code as "FDR", so match both.
    return compact.startswith("R") or "FRD" in compact or "FDR" in compact

@st.cache_resource # Using st.cache_resource to load the PDF only once
def load_pdf_template(template_name_with_extension):
    """
    Loads a PDF template using PyPDF2.PdfReader from the local repository.
    """
    try:
        # Path resolved relative to this tool's bundled assets folder
        path = ASSETS / template_name_with_extension
        if not path.exists():
            st.error(f"Error: PDF template '{template_name_with_extension}' not found at: {path}")
            st.stop() # Stops app execution if template is not found
        # Load the PDF directly from the local path
        return PdfReader(str(path))
    except Exception as e:
        st.error(f"Error loading PDF template '{template_name_with_extension}': {e}")
        st.stop() # Stops app execution in case of a loading error

def fill_and_get_pdf_bytes(pdf_reader_obj, field_values):
    """
    Fills PDF fields from a PdfReader object and returns the filled PDF as bytes.
    Ensures form fields remain interactive.
    """
    try:
        pdf_writer = PdfWriter()

        # Explicitly ensure /AcroForm dictionary exists in PdfWriter
        if "/AcroForm" not in pdf_writer._root_object:
            pdf_writer._root_object[NameObject("/AcroForm")] = DictionaryObject()

        # Add all pages from the template to the writer
        for page in pdf_reader_obj.pages:
            pdf_writer.add_page(page)

        # Fill form fields on the pages
        # update_page_form_field_values applies values to existing fields.
        # Fields not in field_values will not be altered, preserving their interactivity.
        for i, page in enumerate(pdf_writer.pages):
            pdf_writer.update_page_form_field_values(page, field_values)

        # Ensure the PDF displays the filled values (NeedAppearances)
        # This is crucial for text fields to appear correctly.
        # For untouched checkboxes, it helps maintain the form structure.
        if "/AcroForm" in pdf_reader_obj.trailer["/Root"]:
            acroform = pdf_reader_obj.trailer["/Root"]["/AcroForm"]
            acroform.update({NameObject("/NeedAppearances"): BooleanObject(True)})
            pdf_writer._root_object.update({NameObject("/AcroForm"): acroform})
        else:
            pdf_writer._root_object.update({
                NameObject("/AcroForm"): DictionaryObject({
                    NameObject("/NeedAppearances"): BooleanObject(True)
                })
            })

        # Save the filled PDF to a memory buffer
        buffer = io.BytesIO()
        pdf_writer.write(buffer)
        buffer.seek(0) # Rewind the buffer to the beginning
        return buffer.getvalue()
    except Exception as e:
        # Re-raise the exception for the calling function to handle
        raise Exception(f"Failed to fill PDF: {e}")

# --- Load PDF Templates ---
# Ensure these files are in the root of your GitHub repository
worksheet_template_reader = load_pdf_template("Worksheet-Stages-2C-and-3.pdf")
assessment_template_reader = load_pdf_template("Assessment-Form-Stages-2AB.pdf")

# --- App Title and Instruction ---
st.title("📄 IWBF Player Assessment Forms Generator")

# Download button for template
with open(ASSETS / "Players.xlsx", "rb") as f:
    st.download_button(
        label="📥 Click here to download the template file Players.xlsx",
        data=f,
        file_name="Players.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Brief tutorial in English
st.markdown("""
**Step 1** – Download the `Players.xlsx` template file to your computer by clicking the button above.\\
**Step 2** – Fill out the `Players.xlsx` spreadsheet with the player data, then save and close the file.\\
Note: The data in the spreadsheet is just an example; you can replace it with your own data.\\
**Step 3** – Upload the `Players.xlsx` file below and click the "Generate Player Forms" button.\\
**Step 4** – Download the generated forms by clicking the "Click to Download Generated Forms" button.

**About the `status` column:**\\
• **New** players (`N` or `New`) get **both** forms — the *Assessment Form Stages 2AB* and the *Worksheet Stages 2C and 3*.\\
• **Review** players (`R`, `Review`, `R-FRD`, `RFD`, `FRD`, … any review status) only need to be re-classified, so they get **only** the *Worksheet Stages 2C and 3*; the *Stages 2AB* form is **not** created for them.\\
• A blank/unknown status is treated as **New** (both forms).
""")

st.markdown("---")

# --- File Uploader Component ---
uploaded_file = st.file_uploader(
    "Select your Players.xlsx file",
    type=["xlsx"],
    help="The Excel file must contain the following columns: 'number', 'name', 'proposed-class', 'status', 'dob', 'date', 'competition'. The output folders are named after each sheet/tab."
)

# --- Processing Logic ---
if uploaded_file:
    st.success(f"File selected: **{uploaded_file.name}**")

    # Button to start generation
    if st.button("Generate Player Forms"):
        st.info("Starting PDF generation. This might take a few minutes...")

        # Feedback elements for the user
        progress_text = st.empty()
        progress_bar = st.progress(0)

        total_pdfs_to_generate = 0
        generated_pdfs_count = 0
        failed_items = [] # List to store information about failed PDFs

        try:
            # Load all sheets from the Excel file
            excel_data = io.BytesIO(uploaded_file.getvalue())
            # Read the jersey 'number' and the 'proposed-class' as text so that
            # values like "00", "01" keep their leading zeros and the class is
            # not coerced to a float (e.g. 12.0). Date columns stay as dates.
            planilhas = pd.read_excel(
                excel_data,
                sheet_name=None,
                dtype={"number": str, "proposed-class": str},
            )

            # Calculate total PDFs for the progress bar. A "new" player yields
            # 2 PDFs (Worksheet + Assessment), a "review" player yields only 1
            # (Worksheet), so we count per-row based on the 'status' column.
            for sheet_name, df in planilhas.items():
                for _, row in df.iterrows():
                    total_pdfs_to_generate += 1 if is_review_status(row.get("status")) else 2

            # In-memory buffer for the output ZIP file
            zip_buffer = io.BytesIO()
            
            # Use zipfile to create the ZIP archive in memory
            with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED, False) as zip_file:
                for sheet_name, df in planilhas.items():
                    # Validate required columns. The 'status' column is optional:
                    # if it is missing, every player is treated as "new" (both
                    # forms), which keeps older spreadsheets working unchanged.
                    required_columns = ["number", "name", "proposed-class", "dob", "date", "competition"]
                    if not all(col in df.columns for col in required_columns):
                        st.error(f"Error: Missing required columns in sheet **'{sheet_name}'**. Required: `{', '.join(required_columns)}`")
                        st.stop() # Stops execution if columns are missing

                    for index, row in df.iterrows():
                        player_name = str(row.get("name", "N/A"))
                        player_number = str(row.get("number", "N/A"))

                        # A "review" player skips the Stages 2AB assessment, so
                        # this row produces only 1 PDF instead of 2. Use this
                        # count to keep the progress bar accurate on skips/errors.
                        review = is_review_status(row.get("status"))
                        expected_pdfs = 1 if review else 2

                        # Basic validation for essential data
                        if pd.isna(row["name"]) or pd.isna(row["number"]):
                            error_msg = f"Skipping row {index+2} (name: '{player_name}') in sheet '{sheet_name}') due to missing 'name' or 'number'."
                            failed_items.append(error_msg)
                            generated_pdfs_count += expected_pdfs
                            progress_bar.progress(generated_pdfs_count / total_pdfs_to_generate)
                            progress_text.text(f"Progress: {generated_pdfs_count}/{total_pdfs_to_generate} PDFs generated. (Skipped: {player_name})")
                            continue

                        try:
                            # --- Fill Worksheet (Form 1) — generated for EVERY player ---
                            field_values_worksheet = {
                                "number": player_number,
                                "proposed-class": format_class(row.get("proposed-class", "")),
                                "name": player_name,
                                "date": format_date(row.get("date", "")),
                                "competition": str(row.get("competition", "")),
                                "xnumber": player_number,
                                "xproposed-class": format_class(row.get("proposed-class", "")),
                                "xname": player_name,
                                "xdate": format_date(row.get("date", "")),
                                "xcompetition": str(row.get("competition", "")),
                            }
                            # Pass the template reader object to the filling function
                            worksheet_bytes = fill_and_get_pdf_bytes(worksheet_template_reader, field_values_worksheet)

                            # Renomeia a pasta de saída para "Stages 2C and 3"
                            zip_file.writestr(f"{sheet_name}/Stages 2C and 3/{player_name}-Worksheet-Stages-2C-and-3.pdf", worksheet_bytes)
                            generated_pdfs_count += 1
                            progress_bar.progress(generated_pdfs_count / total_pdfs_to_generate)
                            progress_text.text(f"Progress: {generated_pdfs_count}/{total_pdfs_to_generate} PDFs generated. (Processing: {player_name})")

                            # --- Fill Assessment Form (Form 2) — only for NEW players ---
                            # A "review" player was already classified and skips
                            # the Stages 2AB assessment, so this form is not created
                            # for them (the Stages 2AB folder only appears if at
                            # least one new player exists in the sheet).
                            if not review:
                                field_values_assessment = {
                                    "name": player_name,
                                    "dob": format_date(row.get("dob", "")),
                                }
                                # Pass the template reader object to the filling function
                                assessment_bytes = fill_and_get_pdf_bytes(assessment_template_reader, field_values_assessment)

                                # Renomeia a pasta de saída para "Stages 2AB"
                                zip_file.writestr(f"{sheet_name}/Stages 2AB/{player_name}-Assessment-Form-Stages-2AB.pdf", assessment_bytes)
                                generated_pdfs_count += 1
                                progress_bar.progress(generated_pdfs_count / total_pdfs_to_generate)
                                progress_text.text(f"Progress: {generated_pdfs_count}/{total_pdfs_to_generate} PDFs generated. (Processing: {player_name})")

                        except Exception as e:
                            error_msg = f"Error processing '{player_name}' from sheet '{sheet_name}': {e}"
                            failed_items.append(error_msg)
                            generated_pdfs_count += expected_pdfs
                            progress_bar.progress(min(1.0, generated_pdfs_count / total_pdfs_to_generate))
                            progress_text.text(f"Error with {player_name} (Sheet: {sheet_name}). Continuing...")

            progress_bar.progress(1.0)
            progress_text.text("PDF Generation Complete!")

            zip_buffer.seek(0)

            if not failed_items:
                st.success("All forms generated successfully!")
            else:
                st.warning(f"Generation completed with **{len(failed_items)}** errors or skips. Check the logs for details.")
                for i, msg in enumerate(failed_items[:5]):
                    st.error(f"Error {i+1}: {msg}")
                if len(failed_items) > 5:
                    st.info(f"...and {len(failed_items) - 5} more errors. Check the console for full details.")

            st.download_button(
                label="Click to Download Generated Forms (ZIP)",
                data=zip_buffer,
                file_name="Generated_Forms.zip",
                mime="application/zip",
                help="Download a ZIP file containing all filled PDFs."
            )

        except Exception as e:
            st.error(f"An unexpected error occurred during generation: {e}")
            st.exception(e)

    st.markdown("---")
    st.caption("IWBF Player Assessment Forms Generator.")
