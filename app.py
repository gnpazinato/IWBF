import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter # Usando PyPDF2 conforme solicitado
from PyPDF2.generic import NameObject, BooleanObject, DictionaryObject
import io
import os
import zipfile
# requests não é mais necessário, pois os arquivos são carregados localmente

# --- Configuração da Página do Streamlit (DEVE SER O PRIMEIRO COMANDO STREAMLIT) ---
st.set_page_config(page_title="Gerador de Formulários PDF Automatizado", layout="centered")

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
    Carrega um template PDF usando PyPDF2.PdfReader a partir do repositório local.
    """
    try:
        # O caminho é relativo ao arquivo app.py no repositório
        path = os.path.join(os.path.dirname(__file__), template_name_with_extension)
        if not os.path.exists(path):
            st.error(f"Erro: O template PDF '{template_name_with_extension}' não foi encontrado em: {path}")
            st.stop() # Interrompe a execução do app se o template não for encontrado
        # Carrega o PDF diretamente do caminho local
        return PdfReader(path)
    except Exception as e:
        st.error(f"Erro ao carregar o template PDF '{template_name_with_extension}': {e}")
        st.stop() # Interrompe a execução do app em caso de erro de carregamento

def fill_and_get_pdf_bytes(pdf_reader_obj, field_values):
    """
    Preenche os campos de um PDF a partir de um objeto PdfReader e retorna os bytes do PDF preenchido.
    Garante que os campos de formulário permaneçam interativos.
    """
    try:
        pdf_writer = PdfWriter()

        # Garante explicitamente que o dicionário /AcroForm exista no PdfWriter
        # Isso é uma solução robusta para erros de "No /AcroForm dictionary".
        if "/AcroForm" not in pdf_writer._root_object:
            pdf_writer._root_object[NameObject("/AcroForm")] = DictionaryObject()

        # Adiciona todas as páginas do template ao escritor
        for page in pdf_reader_obj.pages:
            pdf_writer.add_page(page)

        # Preenche os campos do formulário nas páginas
        # update_page_form_field_values aplica os valores aos campos existentes.
        # Campos que não estão em field_values não serão alterados, preservando sua interatividade.
        for i, page in enumerate(pdf_writer.pages):
            pdf_writer.update_page_form_field_values(page, field_values)

        # Garante que o PDF exiba os valores preenchidos (NeedAppearances)
        # Isso é crucial para que os campos de texto apareçam corretamente.
        # Para checkboxes não preenchidos, ajuda a manter a estrutura do formulário.
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

        # Salva o PDF preenchido em um buffer de memória
        buffer = io.BytesIO()
        pdf_writer.write(buffer)
        buffer.seek(0) # Retorna o ponteiro para o início do buffer
        return buffer.getvalue()
    except Exception as e:
        # Relança a exceção para que a função chamadora possa tratá-la
        raise Exception(f"Falha ao preencher PDF: {e}")

# --- Carregamento dos Templates PDF (após st.set_page_config) ---
# Garanta que esses arquivos estejam na raiz do seu repositório GitHub
worksheet_template_reader = load_pdf_template("Worksheet-Stages-2C-and-3.pdf")
assessment_template_reader = load_pdf_template("Assessment-Form-Stages-2AB.pdf")

# --- Título e Descrição do Aplicativo ---
st.title("📄 Gerador de Formulários PDF Automatizado")
st.markdown("Faça o upload do seu arquivo Excel (`Players.xlsx`) para gerar formulários PDF personalizados.")
st.markdown("---")

# --- Componente de Upload de Arquivo ---
uploaded_file = st.file_uploader(
    "Selecione seu arquivo Players.xlsx",
    type=["xlsx"],
    help="O arquivo Excel deve conter as colunas: 'number', 'proposed-class', 'name', 'country', 'date', 'competition', 'dob'."
)

# --- Lógica de Processamento ---
if uploaded_file:
    st.success(f"Arquivo selecionado: **{uploaded_file.name}**")

    # Botão para iniciar a geração
    if st.button("Gerar Worksheets"):
        st.info("Iniciando a geração dos PDFs. Isso pode levar alguns minutos...")

        # Elementos de feedback para o usuário
        progress_text = st.empty()
        progress_bar = st.progress(0)

        total_pdfs_to_generate = 0
        generated_pdfs_count = 0
        failed_items = [] # Lista para armazenar informações sobre PDFs que falharam

        try:
            # Carrega todas as abas do Excel
            excel_data = io.BytesIO(uploaded_file.getvalue())
            planilhas = pd.read_excel(excel_data, sheet_name=None)

            # Calcula o total de PDFs a serem gerados para a barra de progresso
            for sheet_name, df in planilhas.items():
                total_pdfs_to_generate += len(df) * 2 # Cada linha gera 2 PDFs

            # Buffer em memória para o arquivo ZIP de saída
            zip_buffer = io.BytesIO()
            
            # Usa zipfile para criar o arquivo ZIP em memória
            with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED, False) as zip_file:
                for sheet_name, df in planilhas.items():
                    # Validação de colunas obrigatórias
                    required_columns = ["number", "proposed-class", "name", "country", "date", "competition", "dob"]
                    if not all(col in df.columns for col in required_columns):
                        st.error(f"Erro: Colunas obrigatórias faltando na aba **'{sheet_name}'**. Necessário: `{', '.join(required_columns)}`")
                        st.stop() # Interrompe a execução se as colunas estiverem faltando

                    for index, row in df.iterrows():
                        player_name = str(row.get("name", "N/A"))
                        player_number = str(row.get("number", "N/A"))

                        # Validação básica de dados essenciais
                        if pd.isna(row["name"]) or pd.isna(row["number"]):
                            error_msg = f"Pulando linha {index+2} (nome: '{player_name}') na aba '{sheet_name}' devido a 'name' ou 'number' ausente."
                            failed_items.append(error_msg)
                            generated_pdfs_count += 2
                            progress_bar.progress(generated_pdfs_count / total_pdfs_to_generate)
                            progress_text.text(f"Progresso: {generated_pdfs_count}/{total_pdfs_to_generate} PDFs gerados. (Pulado: {player_name})")
                            continue

                        try:
                            # --- Preenchimento do Worksheet (Formulário 1) ---
                            # Apenas campos de texto são incluídos aqui.
                            # Checkboxes não são tocados para preservar a interatividade.
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
                            worksheet_bytes = fill_and_get_pdf_bytes(worksheet_template_reader, field_values_worksheet)
                            
                            zip_file.writestr(f"{sheet_name}/Worksheet/{player_name}-Worksheet-Stages-2C-and-3.pdf", worksheet_bytes)
                            generated_pdfs_count += 1
                            progress_bar.progress(generated_pdfs_count / total_pdfs_to_generate)
                            progress_text.text(f"Progresso: {generated_pdfs_count}/{total_pdfs_to_generate} PDFs gerados. (Processando: {player_name})")

                            # --- Preenchimento do Assessment Form (Formulário 2) ---
                            # Apenas campos de texto são incluídos aqui.
                            # Checkboxes não são tocados para preservar a interatividade.
                            field_values_assessment = {
                                "name": player_name,
                                "country": str(row.get("country", "")),
                                "dob": format_date(row.get("dob", "")),
                            }
                            assessment_bytes = fill_and_get_pdf_bytes(assessment_template_reader, field_values_assessment)

                            zip_file.writestr(f"{sheet_name}/Assessment/{player_name}-Assessment-Form-Stages-2AB.pdf", assessment_bytes)
                            generated_pdfs_count += 1
                            progress_bar.progress(generated_pdfs_count / total_pdfs_to_generate)
                            progress_text.text(f"Progresso: {generated_pdfs_count}/{total_pdfs_to_generate} PDFs gerados. (Processando: {player_name})")

                        except Exception as e:
                            error_msg = f"Erro ao processar '{player_name}' da aba '{sheet_name}': {e}"
                            failed_items.append(error_msg)
                            generated_pdfs_count += 2
                            progress_bar.progress(generated_pdfs_count / total_pdfs_to_generate)
                            progress_text.text(f"Erro com {player_name} (Aba: {sheet_name}). Continuando...")

            progress_bar.progress(1.0)
            progress_text.text("Geração de PDFs Concluída!")

            zip_buffer.seek(0)

            if not failed_items:
                st.success("Todos os PDFs foram gerados com sucesso!")
            else:
                st.warning(f"Geração concluída com **{len(failed_items)}** erros ou pulos. Verifique os logs para detalhes.")
                for i, msg in enumerate(failed_items[:5]):
                    st.error(f"Erro {i+1}: {msg}")
                if len(failed_items) > 5:
                    st.info(f"...e mais {len(failed_items) - 5} erros. Verifique o console para detalhes completos.")

            st.download_button(
                label="Clique para Baixar PDFs Gerados (ZIP)",
                data=zip_buffer,
                file_name="Generated_PDFs.zip",
                mime="application/zip",
                help="Baixe um arquivo ZIP contendo todos os PDFs preenchidos."
            )

        except Exception as e:
            st.error(f"Ocorreu um erro inesperado durante a geração: {e}")
            st.exception(e)

    st.markdown("---")
    st.caption("Desenvolvido para simplificar o preenchimento de formulários PDF.")
