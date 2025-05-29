import streamlit as st
import pandas as pd
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, BooleanObject, DictionaryObject
import io
import os
import zipfile

# --- Streamlit Page Configuration (MUST BE THE FIRST STREAMLIT COMMAND) ---
# Removed 'icon' argument as it caused TypeErrors in certain Streamlit/Python versions.
st.set_page_config(page_title="Automated PDF Forms Generator", layout="centered")

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

@st.cache_resource
def load_pdf_template(template_name):
    """
    Loads a PDF template using pypdf.PdfReader.
    Uses st.cache_resource to load the PDF only once, optimizing app performance.
    """
    try:
        # The Streamlit app runs from the repository root, so path is direct
        path = os.path.join(os.path.dirname(__file__), f"{template_name}.pdf")
        if not os.path.exists(path):
            st.error(f"Error: PDF template '{template_name}.pdf' not found at: {path}")
            st.stop() # Stops app execution if template is not found
        return PdfReader(path)
    except Exception as e:
        st.error(f"Error loading PDF template '{template_name}.pdf': {e}")
        st.stop() # Stops app execution in case of a loading error

def fill_and_get_pdf_bytes(pdf_reader_obj, field_values):
    """
    Fills PDF fields from a PdfReader object and returns the filled PDF as bytes.
    Ensures form fields remain interactive.
    """
    try:
        pdf_writer = PdfWriter()

        # Explicitly ensure /AcroForm dictionary exists in PdfWriter
        # This is a robust workaround for "No /AcroForm dictionary" errors.
        if "/AcroForm" not in pdf_writer._root_object:
            pdf_writer._root_object[NameObject("/AcroForm")] = DictionaryObject()

        # Add all pages from the template to the writer
        for page in pdf_reader_obj.pages:
            pdf_writer.add_page(page)

        # Fill form fields on the pages
        # update_page_form_field_values applies values to existing fields.
        # Fields not in field_values will not be altered, preserving their interactivity (e.g., empty checkboxes).
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

# --- Load PDF Templates (after st.set_page_config) ---
# These calls will now execute correctly as st.set_page_config has been called.
worksheet_template_reader = load_pdf_template("Worksheet-Stages-2C-and-3")
assessment_template_reader = load_pdf_template("Assessment-Form-Stages-2AB")

# --- App Title and Description ---
st.title("📄 IWBF PDF Forms Generator")
st.markdown("Upload your Excel file (`Players.xlsx`) to generate player assessment forms.")
st.markdown("---")

# --- File Uploader Component ---
uploaded_file = st.file_uploader(
    "Select your Players.xlsx file",
    type=["xlsx"],
    help="The Excel file must contain the following columns: 'number', 'proposed-class', 'name', 'country', 'date', 'competition', 'dob'."
)

# --- Processing Logic ---
if uploaded_file:
    st.success(f"File selected: **{uploaded_file.name}**")

    # Button to start generation
    if st.button("Generate Worksheets"):
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
            planilhas = pd.read_excel(excel_data, sheet_name=None)

            # Calculate total PDFs for the progress bar
            for sheet_name, df in planilhas.items():
                total_pdfs_to_generate += len(df) * 2 # Each row generates 2 PDFs

            # In-memory buffer for the output ZIP file
            zip_buffer = io.BytesIO()

            # Use zipfile to create the ZIP archive in memory
            with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED, False) as zip_file:
                for sheet_name, df in planilhas.items():
                    # Validate required columns
                    required_columns = ["number", "proposed-class", "name", "country", "date", "competition", "dob"]
                    if not all(col in df.columns for col in required_columns):
                        st.error(f"Error: Missing required columns in sheet **'{sheet_name}'**. Required: `{', '.join(required_columns)}`")
                        st.stop() # Stops execution if columns are missing

                    for index, row in df.iterrows():
                        player_name = str(row.get("name", "N/A")) # Use .get for robustness
                        player_number = str(row.get("number", "N/A"))

                        # Basic validation for essential data
                        if pd.isna(row["name"]) or pd.isna(row["number"]):
                            error_msg = f"Skipping row {index+2} (name: '{player_name}') in sheet '{sheet_name}' due to missing 'name' or 'number'."
                            failed_items.append(error_msg)
                            # Still increment for progress to keep the bar moving
                            generated_pdfs_count += 2
                            progress_bar.progress(generated_pdfs_count / total_pdfs_to_generate)
                            progress_text.text(f"Progress: {generated_pdfs_count}/{total_pdfs_to_generate} PDFs generated. (Skipped: {player_name})")
                            continue

                        try:
                            # --- Fill Worksheet (Form 1) ---
                            # Only text fields are included here.
                            # Checkboxes are untouched to preserve interactivity.
                            field_values_worksheet = {
                                "number": player_number,
                                "proposed-class": str(row.get("proposed-class", "")),
                                "name": player_name,
                                "country": str(row.get("country", "")),
                                "date": format_date(row.get("date", "")),
                                "competition": str(row.get("competition", "")),
                                # Page 2 fields (with 'x' prefix)
                                "xnumber": player_number,
                                "xproposed-class": str(row.get("proposed-class", "")),
                                "xname": player_name,
                                "xcountry": str(row.get("country", "")),
                                "xdate": format_date(row.get("date", "")),
                                "xcompetition": str(row.get("competition", "")),
                            }
                            worksheet_bytes = fill_and_get_pdf_bytes(worksheet_template_reader, field_values_worksheet)

                            # Add the generated PDF to the ZIP file
                            zip_file.writestr(f"{sheet_name}/Worksheet/{player_name}-Worksheet-Stages-2C-and-3.pdf", worksheet_bytes)
                            generated_pdfs_count += 1
                            progress_bar.progress(generated_pdfs_count / total_pdfs_to_generate)
                            progress_text.text(f"Progress: {generated_pdfs_count}/{total_pdfs_to_generate} PDFs generated. (Processing: {player_name})")

                            # --- Fill Assessment Form (Form 2) ---
                            # Only text fields are included here.
                            # Checkboxes are untouched to preserve interactivity.
                            field_values_assessment = {
                                "name": player_name,
                                "country": str(row.get("country", "")),
                                "dob": format_date(row.get("dob", "")),
                            }
                            assessment_bytes = fill_and_get_pdf_bytes(assessment_template_reader, field_values_assessment)

                            # Add the generated PDF to the ZIP file
                            zip_file.writestr(f"{sheet_name}/Assessment/{player_name}-Assessment-Form-Stages-2AB.pdf", assessment_bytes)
                            generated_pdfs_count += 1
                            progress_bar.progress(generated_pdfs_count / total_pdfs_to_generate)
                            progress_text.text(f"Progress: {generated_pdfs_count}/{total_pdfs_to_generate} PDFs generated. (Processing: {player_name})")

                        except Exception as e:
                            error_msg = f"Error processing '{player_name}' from sheet '{sheet_name}': {e}"
                            failed_items.append(error_msg)
                            # Still increment for progress, but with an error
                            generated_pdfs_count += 2
                            progress_bar.progress(generated_pdfs_count / total_pdfs_to_generate)
                            progress_text.text(f"Error with {player_name} (Sheet: {sheet_name}). Continuing...")

            # Finalize progress bar
            progress_bar.progress(1.0)
            progress_text.text("PDF Generation Complete!")

            # Rewind the ZIP buffer to the beginning for download
            zip_buffer.seek(0)

            # --- Final Message and Download Button ---
            if not failed_items:
                st.success("All PDFs generated successfully!")
            else:
                st.warning(f"Generation completed with **{len(failed_items)}** errors or skips. Check the logs for details.")
                # Display the first 5 errors to the user
                for i, msg in enumerate(failed_items[:5]):
                    st.error(f"Error {i+1}: {msg}")
                if len(failed_items) > 5:
                    st.info(f"...and {len(failed_items) - 5} more errors. Check the console for full details.")

            # Download button for the ZIP file
            st.download_button(
                label="Click to Download Generated PDFs (ZIP)",
                data=zip_buffer,
                file_name="Generated_PDFs.zip",
                mime="application/zip",
                help="Download a ZIP file containing all filled PDFs."
            )

        except Exception as e:
            st.error(f"An unexpected error occurred during generation: {e}")
            st.exception(e) # Displays the full traceback for debugging

    st.markdown("---")
    st.caption("Developed to simplify PDF form filling.")
