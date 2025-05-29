import streamlit as st
import pandas as pd
import io
import os
import zipfile
import pikepdf

# --- Streamlit Page Configuration ---
st.set_page_config(page_title="IWBF Player Assessment Forms Generator", layout="centered")

# --- Helper Functions ---
def format_date(date):
    try:
        return pd.to_datetime(date, dayfirst=True).strftime("%d-%m-%Y")
    except Exception:
        return str(date)

@st.cache_resource
def load_pdf_template(template_name_with_extension):
    try:
        path = os.path.join(os.path.dirname(__file__), template_name_with_extension)
        if not os.path.exists(path):
            st.error(f"Error: PDF template '{template_name_with_extension}' not found at: {path}")
            st.stop()
        pdf = pikepdf.Pdf.open(path)
        return pdf
    except Exception as e:
        st.error(f"Error loading PDF template '{template_name_with_extension}': {e}")
        st.stop()

def fill_and_get_pdf_bytes(pdf_template_obj, field_values):
    import io
    import pikepdf

    try:
        temp_buffer = io.BytesIO()
        pdf_template_obj.save(temp_buffer)
        temp_buffer.seek(0)
        pdf = pikepdf.Pdf.open(temp_buffer)

        acroform = pdf.Root.get("/AcroForm", None)
        if not acroform:
            raise Exception("PDF does not contain an AcroForm.")

        fields = acroform.get("/Fields", [])
        if not fields:
            raise Exception("No fields found in the AcroForm.")

        for field_ref in fields:
            field = field_ref.get_object() if hasattr(field_ref, "get_object") else field_ref
            field_name = field.get("/T", None)
            if field_name and field_name in field_values:
                field[pikepdf.Name("/V")] = pikepdf.String(str(field_values[field_name]))
                field[pikepdf.Name("/Ff")] = 1
                if pikepdf.Name("/AP") in field:
                    del field[pikepdf.Name("/AP")]

        acroform[pikepdf.Name("/NeedAppearances")] = pikepdf.objects.BooleanObject(True)

        output_buffer = io.BytesIO()
        pdf.save(output_buffer)
        output_buffer.seek(0)
        return output_buffer.getvalue()

    except Exception as e:
        raise Exception(f"Failed to fill PDF: {e}")

# --- Load PDF Templates ---
worksheet_template_obj = load_pdf_template("Worksheet-Stages-2C-and-3.pdf")
assessment_template_obj = load_pdf_template("Assessment-Form-Stages-2AB.pdf")

# --- App UI ---
st.title("📄 IWBF Player Assessment Forms Generator")
st.markdown("Upload your Excel file (`Players.xlsx`) to generate player forms.")
st.markdown("---")

uploaded_file = st.file_uploader(
    "Select your Players.xlsx file",
    type=["xlsx"],
    help="The Excel file must contain the following columns: 'number', 'proposed-class', 'name', 'country', 'date', 'competition', 'dob'."
)

if uploaded_file:
    st.success(f"File selected: **{uploaded_file.name}**")

    if st.button("Generate Player Forms"):
        st.info("Starting PDF generation. This might take a few minutes...")

        progress_text = st.empty()
        progress_bar = st.progress(0)

        total_pdfs_to_generate = 0
        generated_pdfs_count = 0
        failed_items = []

        try:
            excel_data = io.BytesIO(uploaded_file.getvalue())
            planilhas = pd.read_excel(excel_data, sheet_name=None)

            for sheet_name, df in planilhas.items():
                total_pdfs_to_generate += len(df) * 2

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED, False) as zip_file:
                for sheet_name, df in planilhas.items():
                    required_columns = ["number", "proposed-class", "name", "country", "date", "competition", "dob"]
                    if not all(col in df.columns for col in required_columns):
                        st.error(f"Error: Missing required columns in sheet **'{sheet_name}'**.")
                        st.stop()

                    for index, row in df.iterrows():
                        player_name = str(row.get("name", "N/A"))
                        player_number = str(row.get("number", "N/A"))

                        if pd.isna(row["name"]) or pd.isna(row["number"]):
                            error_msg = f"Skipping row {index+2} (name: '{player_name}') in sheet '{sheet_name}' due to missing 'name' or 'number'."
                            failed_items.append(error_msg)
                            generated_pdfs_count += 2
                            progress_bar.progress(generated_pdfs_count / total_pdfs_to_generate)
                            progress_text.text(f"Progress: {generated_pdfs_count}/{total_pdfs_to_generate} PDFs generated. (Skipped: {player_name})")
                            continue

                        try:
                            field_values_worksheet = {
                                "number": player_number,
                                "proposed-class": str(row.get("proposed-class", "")),
                                "name": player_name,
                                "country": str(row.get("country", "")),
                                "date": format_date(row.get("date", "")),
                                "competition": str(row.get("competition", "")),
                                "xnumber": player_number,
                                "xproposed-class": str(row.get("proposed-class", "")),
                                "xname": player_name,
                                "xcountry": str(row.get("country", "")),
                                "xdate": format_date(row.get("date", "")),
                                "xcompetition": str(row.get("competition", "")),
                            }
                            worksheet_bytes = fill_and_get_pdf_bytes(worksheet_template_obj, field_values_worksheet)
                            zip_file.writestr(f"{sheet_name}/Stages 2C and 3/{player_name}-Worksheet-Stages-2C-and-3.pdf", worksheet_bytes)

                            generated_pdfs_count += 1
                            progress_bar.progress(generated_pdfs_count / total_pdfs_to_generate)
                            progress_text.text(f"Progress: {generated_pdfs_count}/{total_pdfs_to_generate} PDFs generated. (Processing: {player_name})")

                            field_values_assessment = {
                                "name": player_name,
                                "country": str(row.get("country", "")),
                                "dob": format_date(row.get("dob", "")),
                            }
                            assessment_bytes = fill_and_get_pdf_bytes(assessment_template_obj, field_values_assessment)
                            zip_file.writestr(f"{sheet_name}/Stages 2AB/{player_name}-Assessment-Form-Stages-2AB.pdf", assessment_bytes)

                            generated_pdfs_count += 1
                            progress_bar.progress(generated_pdfs_count / total_pdfs_to_generate)
                            progress_text.text(f"Progress: {generated_pdfs_count}/{total_pdfs_to_generate} PDFs generated. (Processing: {player_name})")

                        except Exception as e:
                            error_msg = f"Error processing '{player_name}' from sheet '{sheet_name}': {e}"
                            failed_items.append(error_msg)
                            generated_pdfs_count += 2
                            progress_bar.progress(generated_pdfs_count / total_pdfs_to_generate)
                            progress_text.text(f"Error with {player_name} (Sheet: {sheet_name}). Continuing...")

            progress_bar.progress(1.0)
            progress_text.text("PDF Generation Complete!")

            zip_buffer.seek(0)

            if not failed_items:
                st.success("All forms generated successfully!")
            else:
                st.warning(f"Generation completed with **{len(failed_items)}** errors or skips.")
                for i, msg in enumerate(failed_items[:5]):
                    st.error(f"Error {i+1}: {msg}")
                if len(failed_items) > 5:
                    st.info(f"...and {len(failed_items) - 5} more errors. Check the console for full details.")

            st.download_button(
                label="Click to Download Generated Forms (ZIP)",
                data=zip_buffer,
                file_name="Generated_Forms.zip",
                mime="application/zip"
            )

        except Exception as e:
            st.error(f"An unexpected error occurred during generation: {e}")
            st.exception(e)

    st.markdown("---")
    st.caption("IWBF Player Assessment Forms Generator.")
