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
    Formats a date to 'dd-mm-yyyy' or returns the value as a string.
    Handles different input types for dates.
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
        # Path is relative to the app.py file in the repository
        path = os.path.join(os.path.dirname(__file__), template_name_with_extension)
        if not os.path.exists(path):
            st.error(f"Error: PDF template '{template_name_with_extension}' not found at: {path}")
            st.stop() # Stops app execution if template is not found
        
        # Open the PDF with pikepdf
        pdf = pikepdf.Pdf.open(path)
        return pdf
    except Exception as e:
        st.error(f"Error loading PDF template '{template_name_with_extension}': {e}")
        st.stop() # Stops app execution in case of a loading error

# Função fill_and_get_pdf_bytes REESCRITA para usar pikepdf
def fill_and_get_pdf_bytes(pdf_template_obj, field_values): # Recebe pikepdf.Pdf object
    """
    Fills PDF fields in a pikepdf.Pdf object and returns the filled PDF as bytes.
    Ensures form fields remain interactive.
    """
    try:
        # Cria uma nova cópia do PDF do template para evitar modificar o objeto em cache.
        # Isso é crucial porque st.cache_resource armazena em cache o objeto Pdf.
        # Precisamos de uma cópia mutável para cada preenchimento.
        temp_buffer = io.BytesIO()
        pdf_template_obj.save(temp_buffer)
        temp_buffer.seek(0)
        pdf = pikepdf.Pdf.open(temp_buffer)

        # Acessa os campos do formulário
        # pikepdf manipula o AcroForm através da propriedade 'form'
        # Não é necessário verificar '/AcroForm' in pdf.root explicitamente aqui
        # pois get_form() lida com isso.
        form_fields = pdf.get_form()

        # Preenche os campos
        for field_name, value in field_values.items():
            if field_name in form_fields.get_fields(): # Verifica se o campo existe
                field = form_fields.get_fields()[field_name]
                field.V = pikepdf.String(str(value)) # Define o valor. pikepdf precisa de um objeto String
                
                # Força a aparência a ser gerada pelo visualizador.
                # AP (Appearance Stream) nulo força o redesenho.
                field.AP = None
                
                # O 'Ff' (Field Flags) pode ser usado para setar flags como Print, NoExport, etc.
                # Para forçar a aparência, setamos NeedAppearances no AcroForm,
                # e pikepdf geralmente lida com a geração de aparência para campos preenchidos.
                
            else: # Adicionado aviso para campos não encontrados, útil para depuração
                st.warning(f"Warning: Field '{field_name}' not found in PDF form. Skipping.")

        # Garante que NeedAppearances seja definido no nível do AcroForm (dicionário de formulário raiz)
        # Isso é crucial para que os visualizadores de PDF renderizem os campos preenchidos corretamente.
        # pikepdf acessa o AcroForm via pdf.root.AcroForm e o cria se não existir.
        if '/AcroForm' in pdf.root:
            pdf.root.AcroForm.NeedAppearances = pikepdf.Boolean(True)
        else:
            # Se não houver AcroForm, cria um (isso é mais um fallback, o template deve ter um)
            pdf.root['/AcroForm'] = pikepdf.Dictionary()
            pdf.root.AcroForm.NeedAppearances = pikepdf.Boolean(True)

        # Salva o PDF preenchido em um buffer de memória
        output_buffer = io.BytesIO()
        pdf.save(output_buffer) # pikepdf otimiza e salva para o buffer
        output_buffer.seek(0)
        return output_buffer.getvalue()
    except Exception as e:
        # Relança a exceção para que a função chamadora possa tratá-la
        raise Exception(f"Failed to fill PDF: {e}")

# --- Load PDF Templates (após st.set_page_config) ---
# Estes serão objetos pikepdf.Pdf, armazenados em cache.
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
                        player_name = str(row.get("name", "N/A"))
                        player_number = str(row.get("number", "N/A"))

                        # Basic validation for essential data
                        if pd.isna(row["name"]) or pd.isna(row["number"]):
                            error_msg = f"Skipping row {index+2} (name: '{player_name}') in sheet '{sheet_name}') due to missing 'name' or 'number'."
                            failed_items.append(error_msg)
                            generated_pdfs_count += 2
                            progress_bar.progress(generated_pdfs_count / total_pdfs_to_generate)
                            progress_text.text(f"Progress: {generated_pdfs_count}/{total_pdfs_to_generate} PDFs generated. (Skipped: {player_name})")
                            continue

                        try:
                            # --- Fill Worksheet (Form 1) ---
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
                            # Pass the pikepdf object to the filling function
                            worksheet_bytes = fill_and_get_pdf_bytes(worksheet_template_obj, field_values_worksheet)
                            
                            # Renomeia a pasta de saída para "Stages 2C and 3"
                            zip_file.writestr(f"{sheet_name}/Stages 2C and 3/{player_name}-Worksheet-Stages-2C-and-3.pdf", worksheet_bytes)
                            generated_pdfs_count += 1
                            progress_bar.progress(generated_pdfs_count / total_pdfs_to_generate)
                            progress_text.text(f"Progress: {generated_pdfs_count}/{total_pdfs_to_generate} PDFs generated. (Processing: {player_name})")

                            # --- Fill Assessment Form (Form 2) ---
                            field_values_assessment = {
                                "name": player_name,
                                "country": str(row.get("country", "")),
                                "dob": format_date(row.get("dob", "")),
                            }
                            # Pass the pikepdf object to the filling function
                            assessment_bytes = fill_and_get_pdf_bytes(assessment_template_obj, field_values_assessment)

                            # Renomeia a pasta de saída para "Stages 2AB"
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
