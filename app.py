import streamlit as st
import pandas as pd
# Removidas as importações de PyPDF2.generic, pois pikepdf tem seus próprios tipos
import io
import os
import zipfile
import pikepdf # Importa pikepdf

# --- Configuração da Página do Streamlit (DEVE SER O PRIMEIRO COMANDO STREAMLIT) ---
st.set_page_config(page_title="IWBF Player Assessment Forms Generator", layout="centered")

# --- Funções Auxiliares ---

def format_date(date):
    """
    Formata uma data para o formato 'dd-mm-yyyy' ou retorna o valor como string.
    Lida com diferentes tipos de entrada para datas.
    """
    try:
        return pd.to_datetime(date).strftime("%d-%m-%Y")
    except Exception:
        return str(date)

@st.cache_resource # Usando st.cache_resource para carregar o PDF apenas uma vez
def load_pdf_template(template_name_with_extension):
    """
    Carrega um template PDF usando pikepdf.Pdf.open() a partir do repositório local.
    Retorna o pikepdf.Pdf object.
    """
    try:
        # O caminho é relativo ao arquivo app.py no repositório
        path = os.path.join(os.path.dirname(__file__), template_name_with_extension)
        if not os.path.exists(path):
            st.error(f"Error: PDF template '{template_name_with_extension}' not found at: {path}")
            st.stop() # Interrompe a execução do app se o template não for encontrado
        
        # Abre o PDF com pikepdf
        pdf = pikepdf.Pdf.open(path)
        return pdf
    except Exception as e:
        st.error(f"Error loading PDF template '{template_name_with_extension}': {e}")
        st.stop() # Interrompe a execução do app em caso de erro de carregamento

# Função fill_and_get_pdf_bytes REESCRITA para usar pikepdf
def fill_and_get_pdf_bytes(pdf_template_obj, field_values): # Recebe pikepdf.Pdf object
    """
    Preenche os campos de um pikepdf.Pdf object e retorna os bytes do PDF preenchido.
    Garante que os campos de formulário permaneçam interativos.
    """
    try:
        # Cria uma nova cópia do PDF do template para evitar modificar o objeto em cache.
        temp_buffer = io.BytesIO()
        pdf_template_obj.save(temp_buffer)
        temp_buffer.seek(0)
        pdf = pikepdf.Pdf.open(temp_buffer)

        # Acessa os campos do formulário
        form_fields = pdf.forms 

        # Preenche os campos
        for field_name, value in field_values.items():
            if field_name in form_fields.get_fields(): # Verifica se o campo existe
                field = form_fields.get_fields()[field_name]
                field.V = pikepdf.String(str(value)) # Define o valor. pikepdf precisa de um objeto String
                
                field.AP = None # Limpa o stream de aparência para forçar o visualizador a redesenhar
                
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
        # Relança a exceção para que a função chamadora possa tratar
        raise Exception(f"Failed to fill PDF: {e}")

# --- Carregamento dos Templates PDF (após st.set_page_config) ---
# Estes serão objetos pikepdf.Pdf, armazenados em cache.
worksheet_template_obj = load_pdf_template("Worksheet-Stages-2C-and-3.pdf")
assessment_template_obj = load_pdf_template("Assessment-Form-Stages-2AB.pdf")

# --- Título e Descrição do Aplicativo ---
st.title("📄 IWBF Player Assessment Forms Generator")
st.markdown("Upload your Excel file (`Players.xlsx`) to generate player forms.")
st.markdown("---")

# --- File Uploader Component ---
uploaded_file = st.file_uploader(
    "Select your Players.xlsx file",
    type=["xlsx"],
    help="The Excel file must contain the following columns: 'number', 'proposed-class', 'name', 'country', 'date', 'competition', 'dob'."
)

# --- Lógica de Processamento ---
if uploaded_file:
    st.success(f"File selected: **{uploaded_file.name}**")

    # Botão para iniciar a geração
    if st.button("Generate Player Forms"):
        st.info("Starting PDF generation. This might take a few minutes...")

        # Feedback elements for the user
        progress_text = st.empty()
        progress_bar = st.progress(0)

        total_pdfs_to_generate = 0
        generated_pdfs_count = 0
        failed_items = [] # Lista para armazenar informações sobre PDFs que falharam

        # --- BLOCO try EXTERNO ---
        try:
            # Carrega todas as abas do Excel
            excel_data = io.BytesIO(uploaded_file.getvalue())
            planilhas = pd.read_excel(excel_data, sheet_name=None)

            # Calcula o total de PDFs a serem gerados para a barra de progresso
            for sheet_name, df in planilhas.items():
                total_pdfs_to_generate += len(df) * 2 # Cada linha gera 2 PDFs

            # In-memory buffer for the output ZIP file
            zip_buffer = io.BytesIO()
            
            # Use zipfile para criar o ZIP archive em memória
            with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED, False) as zip_file:
                for sheet_name, df in planilhas.items():
                    # Validação de colunas obrigatórias
                    required_columns = ["number", "proposed-class", "name", "country", "date", "competition", "dob"]
                    if not all(col in df.columns for col in required_columns):
                        st.error(f"Error: Missing required columns in sheet **'{sheet_name}'**. Required: `{', '.join(required_columns)}`")
                        st.stop() # Interrompe a execução se as colunas estiverem faltando

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

            # --- BLOCO FINAL DE PROGRESSO E DOWNLOAD (ALINHADO COM O try EXTERNO) ---
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

        # --- BLOCO except EXTERNO (ALINHADO COM O try) ---
        except Exception as e:
            st.error(f"An unexpected error occurred during generation: {e}")
            st.exception(e)

    # --- BLOCO st.markdown e st.caption (ALINHADO COM O if uploaded_file) ---
    st.markdown("---")
    st.caption("IWBF Player Assessment Forms Generator.")
