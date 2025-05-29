import streamlit as st
import pandas as pd
# Removidas as importações de PyPDF2.generic, pois pikepdf tem seus próprios tipos
import io
import os
import zipfile
import pikepdf # Importa pikepdf

# --- Streamlit Page Configuration (MUST BE THE FIRST STREAMLIT COMMAND) ---
st.set_page_config(page_title="IWBF Player Assessment Forms Generator", layout="centered")

# --- Helper Functions ---

def format_date(date):
    """
    Formata uma data para o formato 'dd-mm-yyyy' ou retorna o valor como string.
    Lida com diferentes tipos de entrada para datas.
    """
    try:
        return pd.to_datetime(date).strftime("%d-%m-%Y")
    except Exception:
        return str(date)

@st.cache_resource # Using st.cache_resource to load the PDF only once
def load_pdf_template(template_name_with_extension):
    """
    Loads a PDF template using pikepdf.Pdf.open() from the local repository.
    Returns the pikepdf.Pdf object.
    """
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

# Função fill_and_get_pdf_bytes REESCRITA para usar pikepdf
def fill_and_get_pdf_bytes(pdf_template_obj, field_values): # Recebe pikepdf.Pdf object
    """
    Fills PDF fields in a pikepdf.Pdf object and returns the filled PDF as bytes.
    Ensures form fields remain interactive.
    """
    try:
        # Cria uma nova cópia do PDF do template para evitar modificar o objeto em cache.
        temp_buffer = io.BytesIO()
        pdf_template_obj.save(temp_buffer)
        temp_buffer.seek(0)
        pdf = pikepdf.Pdf.open(temp_buffer)

        # Acessa os campos do formulário
        # ALTERADO: Usa pdf.forms em vez de pdf.get_form()
        form_fields = pdf.forms # <--- MUDANÇA AQUI

        # Preenche os campos
        for field_name, value in field_values.items():
            if field_name in form_fields.get_fields(): # get_fields() é um método de pdf.forms
                field = form_fields.get_fields()[field_name]
                field.V = pikepdf.String(str(value))
                field.AP = None # Clear appearance stream to force viewer to redraw
            else:
                st.warning(f"Warning: Field '{field_name}' not found in PDF form. Skipping.")

        # Garante que NeedAppearances seja definido no nível do AcroForm
        if '/AcroForm' in pdf.root:
            pdf.root.AcroForm.NeedAppearances = pikepdf.Boolean(True)
        else:
            pdf.root['/AcroForm'] = pikepdf.Dictionary()
            pdf.root.AcroForm.NeedAppearances = pikepdf.Boolean(True)

        # Salva o PDF preenchido em um buffer de memória
        output_buffer = io.BytesIO()
        pdf.save(output_buffer)
        output_buffer.seek(0)
        return output_buffer.getvalue()
    except Exception as e:
        raise Exception(f"Failed to fill PDF: {e}")

# --- Load PDF Templates (após st.set_page_config) ---
worksheet_template_obj = load_pdf_template("Worksheet-Stages-2C-and-3.pdf")
assessment_template_obj = load_pdf_template("Assessment-Form-Stages-2AB.pdf")

# --- App Title and Description ---
st.title("📄 IWBF Player Assessment Forms Generator")
st.markdown("Upload your Excel file (`Players.xlsx`) to generate player forms.")
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
                        st.error(f"Error: Missing required columns in sheet **'{sheet_name}'**. Required: `{', '.join(required_columns)}`")
                        st.stop()

                    for index, row in df.iterrows():
                        player_name = str(row.get("name", "N/A"))
                        player_number = str(row.get("number", "N/A"))

                        if pd.isna(row["name"]) or pd.isna(row["number"]):
                            error_msg = f"Skipping row {index+2} (name: '{player_name}') in sheet '{sheet_name}') due to missing 'name' or 'number'."
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
                            generated_pdfs_count
