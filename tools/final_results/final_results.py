import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.generic import NameObject, BooleanObject, DictionaryObject
import io
import zipfile
from pathlib import Path

# Assets bundled with this tool, resolved relative to THIS file (not the CWD).
ASSETS = Path(__file__).resolve().parent / "assets"

TEMPLATE_XLSX = "classification-results-spreadsheet-template.xlsx"

# The two output variants: the label shown on the button, the PDF template to
# fill, and the suffix used in each generated file name.
VARIANTS = {
    "Stage 2": {
        "template": "Classification-Results-Stage-2.pdf",
        "suffix": "Classification Results - Stage 2",
    },
    "Final": {
        "template": "Classification-Results-Final.pdf",
        "suffix": "Classification Results - Final",
    },
}

MAX_PLAYERS = 12  # the PDF forms have 12 player rows


# --- Helper Functions ---

def clean(value):
    """Returns a trimmed string for a cell value, or '' for empty/NaN."""
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def format_date(date):
    """Formats a date to 'dd-mm-yyyy' (day-first), or returns it as a string."""
    if pd.isna(date) or str(date).strip() == "":
        return ""
    try:
        return pd.to_datetime(date, dayfirst=True).strftime("%d-%m-%Y")
    except Exception:
        return str(date).strip()


def format_class(value):
    """Formats a sport class as 'X.0'/'X.5' with a dot decimal
    (e.g. 1.0, 1.5, 2.0, 2.5) — never '1,0' nor a bare '1'.
    Non-numeric values are returned trimmed, as-is.
    """
    text = clean(value).replace(",", ".")
    if text == "":
        return ""
    try:
        return f"{float(text):.1f}"
    except ValueError:
        return text


@st.cache_resource  # load each PDF template only once
def load_pdf_template(template_name):
    """Loads a fillable PDF template from this tool's assets folder."""
    path = ASSETS / template_name
    if not path.exists():
        st.error(f"Error: PDF template '{template_name}' not found at: {path}")
        st.stop()
    return PdfReader(str(path))


def fill_and_get_pdf_bytes(pdf_reader_obj, field_values):
    """Fills a PdfReader's AcroForm fields and returns the filled PDF as bytes.

    Fields not present in ``field_values`` are left untouched, keeping the form
    interactive; ``/NeedAppearances`` ensures the values render in all viewers.
    """
    pdf_writer = PdfWriter()

    if "/AcroForm" not in pdf_writer._root_object:
        pdf_writer._root_object[NameObject("/AcroForm")] = DictionaryObject()

    for page in pdf_reader_obj.pages:
        pdf_writer.add_page(page)

    for page in pdf_writer.pages:
        pdf_writer.update_page_form_field_values(page, field_values)

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

    buffer = io.BytesIO()
    pdf_writer.write(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def first_value(df, column):
    """First non-empty value of a column (header info repeated across rows)."""
    if column not in df.columns:
        return ""
    for v in df[column]:
        text = clean(v)
        if text:
            return text
    return ""


def build_field_values(df):
    """Maps one team's sheet to PDF field values.

    Returns (field_values, number_of_players, overflow_count).
    """
    gender = first_value(df, "gender").upper()
    field_values = {
        "Event": first_value(df, "competition"),
        "Location": first_value(df, "location-of-the-competition"),
        "COUNTRY": first_value(df, "country"),
        "Male": "X" if gender.startswith("M") else "",
        "Female": "X" if gender.startswith("F") else "",
    }

    # Keep only rows that actually have a player name, capped at 12.
    players = [row for _, row in df.iterrows() if clean(row.get("name"))]
    overflow = max(0, len(players) - MAX_PLAYERS)
    players = players[:MAX_PLAYERS]

    for i, row in enumerate(players, start=1):
        field_values[f"Row{i}"] = clean(row.get("number"))
        field_values[f"PLAYER FAMILY NAME Given NameRow{i}"] = clean(row.get("name"))
        field_values[f"Date of birth ddmmyyyyRow{i}"] = format_date(row.get("dob"))
        field_values[f"Sport classRow{i}"] = format_class(row.get("sport-class"))
        field_values[f"SCSRow{i}"] = clean(row.get("sport-class-status"))
        field_values[f"RemarkRow{i}"] = clean(row.get("remark"))

    return field_values, len(players), overflow


def generate_zip(variant_key, sheets):
    """Generates one filled PDF per team sheet; returns (zip_bytes, count, notes)."""
    variant = VARIANTS[variant_key]
    template_reader = load_pdf_template(variant["template"])
    suffix = variant["suffix"]

    zip_buffer = io.BytesIO()
    notes = []
    generated = 0

    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for sheet_name, df in sheets.items():
            field_values, n_players, overflow = build_field_values(df)
            if n_players == 0:
                notes.append(f"Sheet '{sheet_name}' skipped (no players found).")
                continue
            try:
                pdf_bytes = fill_and_get_pdf_bytes(template_reader, field_values)
            except Exception as e:
                notes.append(f"Sheet '{sheet_name}' failed: {e}")
                continue
            zip_file.writestr(f"{sheet_name} {suffix}.pdf", pdf_bytes)
            generated += 1
            if overflow:
                notes.append(
                    f"Sheet '{sheet_name}': only the first {MAX_PLAYERS} players were used "
                    f"({overflow} extra ignored)."
                )

    zip_buffer.seek(0)
    return zip_buffer.getvalue(), generated, notes


# --- App UI ---

st.title("🏆 Final Results Generator")

with open(ASSETS / TEMPLATE_XLSX, "rb") as f:
    st.download_button(
        label="📥 Click here to download the template file "
              "(classification-results-spreadsheet-template.xlsx)",
        data=f,
        file_name=TEMPLATE_XLSX,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.markdown("""
**Step 1** – Download the template spreadsheet above.\\
**Step 2** – Fill **one sheet (tab) per team**, one player per row (up to 12 players per team), using the provided column headers.\\
**Step 3** – Upload the filled spreadsheet below.\\
**Step 4** – Click one of the two buttons to generate the PDFs (one per team) and download the ZIP.
""")

st.markdown("---")

uploaded_file = st.file_uploader(
    "Select your filled classification-results-spreadsheet-template.xlsx file",
    type=["xlsx"],
    help="Each sheet (tab) is one team. Columns: competition, "
         "location-of-the-competition, country, gender, number, name, dob, "
         "sport-class, sport-class-status, remark.",
)

if uploaded_file:
    st.success(f"File selected: **{uploaded_file.name}**")

    # Forget any previous result when a different file is uploaded.
    file_sig = (uploaded_file.name, len(uploaded_file.getvalue()))
    if st.session_state.get("fr_file_sig") != file_sig:
        st.session_state["fr_file_sig"] = file_sig
        st.session_state.pop("fr_result", None)

    try:
        sheets = pd.read_excel(
            io.BytesIO(uploaded_file.getvalue()), sheet_name=None, dtype=str
        )
    except Exception as e:
        st.error(f"Could not read the spreadsheet: {e}")
        st.stop()

    st.caption(f"{len(sheets)} team sheet(s) detected: {', '.join(sheets.keys())}")

    col1, col2 = st.columns(2)
    with col1:
        stage2_clicked = st.button("Classification Results - Stage 2", use_container_width=True)
    with col2:
        final_clicked = st.button("Classification Results - Final", use_container_width=True)

    if stage2_clicked or final_clicked:
        chosen = "Stage 2" if stage2_clicked else "Final"
        with st.spinner(f"Generating '{chosen}' PDFs..."):
            zip_bytes, generated, notes = generate_zip(chosen, sheets)
        st.session_state["fr_result"] = {
            "variant": chosen,
            "zip": zip_bytes,
            "generated": generated,
            "notes": notes,
        }

    result = st.session_state.get("fr_result")
    if result:
        chosen = result["variant"]
        if result["generated"] == 0:
            st.error("No PDFs were generated — no players found in any sheet.")
        else:
            st.success(f"{result['generated']} PDF(s) generated for '{chosen}'.")
        for note in result["notes"]:
            st.warning(note)
        if result["generated"] > 0:
            st.download_button(
                label=f"Click to Download '{chosen}' PDFs (ZIP)",
                data=result["zip"],
                file_name=f"Classification Results - {chosen}.zip",
                mime="application/zip",
            )

st.markdown("---")
st.caption("Final Results Generator.")
