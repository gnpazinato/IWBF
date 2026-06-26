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

# The two spreadsheet-driven (batch) variants: the PDF template to fill and the
# suffix used in each generated file name.
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

# The browser-filled (individual) variant uses its own template.
INDIVIDUAL_TEMPLATE = "Classification-Individual-Result.pdf"
INDIVIDUAL_FILENAME = "Classification-Individual-Result.pdf"

MAX_PLAYERS = 12  # the PDF forms have 12 player rows

# Column labels for the in-browser individual player table.
COL_NUM = "#"
COL_NAME = "PLAYER (FAMILY NAME, Given Name)"
COL_DOB = "Date of birth (dd/mm/yyyy)"
COL_CLASS = "Sport class"
COL_SCS = "SCS"
COL_REMARK = "Remark"


# --- Helper Functions ---

def clean(value):
    """Returns a trimmed string for a cell value, or '' for empty/NaN."""
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def format_date(date):
    """Formats a date to 'dd/mm/yyyy' (day-first) to match the form's own
    'Date of birth dd/mm/yyyy' column header, or returns it as a string."""
    if pd.isna(date) or str(date).strip() == "":
        return ""
    try:
        return pd.to_datetime(date, dayfirst=True).strftime("%d/%m/%Y")
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


def header_field_values(event, location, country, gender):
    """The five form-header fields, shared by every variant."""
    g = clean(gender).upper()
    return {
        "Event": clean(event),
        "Location": clean(location),
        "COUNTRY": clean(country),
        "Male": "X" if g.startswith("M") else "",
        "Female": "X" if g.startswith("F") else "",
    }


def set_player_row(field_values, i, number, name, dob, sport_class, scs, remark):
    """Writes one player's six fields for row ``i`` (1-based), with the
    consistent formatting used across every variant."""
    field_values[f"Row{i}"] = clean(number)
    field_values[f"PLAYER FAMILY NAME Given NameRow{i}"] = clean(name)
    field_values[f"Date of birth ddmmyyyyRow{i}"] = format_date(dob)
    field_values[f"Sport classRow{i}"] = format_class(sport_class)
    field_values[f"SCSRow{i}"] = clean(scs)
    field_values[f"RemarkRow{i}"] = clean(remark)


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
    """Maps one team's sheet to PDF field values (spreadsheet / batch flow).

    Returns (field_values, number_of_players, overflow_count).
    """
    field_values = header_field_values(
        first_value(df, "competition"),
        first_value(df, "location-of-the-competition"),
        first_value(df, "country"),
        first_value(df, "gender"),
    )

    # Keep only rows that actually have a player name, capped at 12.
    players = [row for _, row in df.iterrows() if clean(row.get("name"))]
    overflow = max(0, len(players) - MAX_PLAYERS)
    players = players[:MAX_PLAYERS]

    for i, row in enumerate(players, start=1):
        set_player_row(
            field_values, i,
            row.get("number"), row.get("name"), row.get("dob"),
            row.get("sport-class"), row.get("sport-class-status"), row.get("remark"),
        )

    return field_values, len(players), overflow


def build_individual_field_values(event, location, country, gender, players):
    """Maps the in-browser individual form to PDF field values.

    ``players`` is a list of dicts (already filtered to those with a name),
    each with keys: number, name, dob, sport_class, scs, remark.
    """
    field_values = header_field_values(event, location, country, gender)
    for i, p in enumerate(players[:MAX_PLAYERS], start=1):
        set_player_row(
            field_values, i,
            p.get("number"), p.get("name"), p.get("dob"),
            p.get("sport_class"), p.get("scs"), p.get("remark"),
        )
    return field_values


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


# --- Navigation ---

def go(view):
    """Switches the visible view and clears any stale generated output."""
    st.session_state["fr_view"] = view
    st.session_state.pop("fr_result", None)
    st.session_state.pop("fr_individual_pdf", None)


# --- App UI ---

st.title("📝 IWBF Results Forms Generator")

view = st.session_state.get("fr_view", "menu")


# ============================ MENU ============================
if view == "menu":
    st.markdown(
        "Choose which classification results form you want to generate. "
        "There are **three options**, grouped into **two ways of working**:"
    )

    # --- Option group 1: spreadsheet-driven (Stage 2 + Final) ---
    with st.container(border=True):
        st.markdown("#### 📊 From a spreadsheet — one PDF per team")
        st.markdown(
            "Use this for **Classification Results - Stage 2** and "
            "**Classification Results - Final**. They work the same way:"
        )
        st.markdown("""
**Step 1** – Download the template spreadsheet (one tab per team) using the button below.\\
**Step 2** – Fill **one sheet (tab) per team**, one player per row (up to 12 players).\\
**Step 3** – Pick **Stage 2** or **Final** below, then upload your filled spreadsheet.\\
**Step 4** – Download one ready-to-sign PDF per team (as a ZIP).
""")
        with open(ASSETS / TEMPLATE_XLSX, "rb") as f:
            st.download_button(
                label="📥 Download the template spreadsheet "
                      "(classification-results-spreadsheet-template.xlsx)",
                data=f,
                file_name=TEMPLATE_XLSX,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        c1, c2 = st.columns(2)
        with c1:
            st.button("Classification Results - Stage 2", use_container_width=True,
                      on_click=go, args=("stage2",))
        with c2:
            st.button("Classification Results - Final", use_container_width=True,
                      on_click=go, args=("final",))

    st.markdown("")

    # --- Option group 2: in-browser single form (Individual) ---
    with st.container(border=True):
        st.markdown("#### ✍️ In the browser — a single form")
        st.markdown(
            "Use this for **Classification Results - Individual**. "
            "Fill **one form right here** — event, location, country, team and up to "
            "12 players — then download the filled PDF. **No spreadsheet needed.**"
        )
        st.button("Classification Results - Individual", use_container_width=True,
                  on_click=go, args=("individual",))


# ===================== BATCH (Stage 2 / Final) =====================
elif view in ("stage2", "final"):
    variant_key = "Stage 2" if view == "stage2" else "Final"

    st.button("← Back to menu", on_click=go, args=("menu",))
    st.subheader(f"Classification Results - {variant_key}")

    st.markdown(
        "Upload your filled spreadsheet below. One ready-to-sign PDF is generated "
        "**per team** (one tab = one team) and bundled into a single ZIP."
    )

    with st.expander("Need the template spreadsheet again?"):
        with open(ASSETS / TEMPLATE_XLSX, "rb") as f:
            st.download_button(
                label="📥 Download the template spreadsheet",
                data=f,
                file_name=TEMPLATE_XLSX,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    uploaded_file = st.file_uploader(
        "Select your filled classification-results-spreadsheet-template.xlsx file",
        type=["xlsx"],
        help="Each sheet (tab) is one team. Columns: competition, "
             "location-of-the-competition, country, gender, number, name, "
             "sport-class, sport-class-status, dob, remark.",
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

        if st.button(f"Generate '{variant_key}' PDFs", use_container_width=True,
                     type="primary"):
            with st.spinner(f"Generating '{variant_key}' PDFs..."):
                zip_bytes, generated, notes = generate_zip(variant_key, sheets)
            st.session_state["fr_result"] = {
                "variant": variant_key,
                "zip": zip_bytes,
                "generated": generated,
                "notes": notes,
            }

        result = st.session_state.get("fr_result")
        if result and result["variant"] == variant_key:
            if result["generated"] == 0:
                st.error("No PDFs were generated — no players found in any sheet.")
            else:
                st.success(f"{result['generated']} PDF(s) generated for '{variant_key}'.")
            for note in result["notes"]:
                st.warning(note)
            if result["generated"] > 0:
                st.download_button(
                    label=f"⬇️ Click to Download '{variant_key}' PDFs (ZIP)",
                    data=result["zip"],
                    file_name=f"Classification Results - {variant_key}.zip",
                    mime="application/zip",
                )


# ===================== INDIVIDUAL (in-browser) =====================
elif view == "individual":
    st.button("← Back to menu", on_click=go, args=("menu",))
    st.subheader("Classification Results - Individual")

    st.markdown("""
Fill the form below and click **Generate PDF**. Fill **at least one player** (a player
row is only used if it has a name). Up to **12 players** can be added.
""")

    col1, col2 = st.columns(2)
    with col1:
        event = st.text_input("Event")
        country = st.text_input("Country")
    with col2:
        location = st.text_input("Location")
        gender = st.selectbox("Team", ["Male", "Female"])

    st.markdown(
        "**Players** — a row is included on the form only if it has a player **name**. "
        "Fill in as many of the 12 rows as you need (at least one). `Remark` is optional."
    )

    empty_players = pd.DataFrame(
        [{COL_NUM: "", COL_NAME: "", COL_DOB: "", COL_CLASS: "", COL_SCS: "", COL_REMARK: ""}
         for _ in range(MAX_PLAYERS)]
    )

    edited = st.data_editor(
        empty_players,
        num_rows="fixed",
        hide_index=True,
        use_container_width=True,
        key="fr_individual_players",
        column_config={
            COL_NUM: st.column_config.TextColumn(
                COL_NUM, width="small",
                help="Jersey number (leading zeros are preserved)."),
            COL_NAME: st.column_config.TextColumn(COL_NAME, width="large"),
            COL_DOB: st.column_config.TextColumn(COL_DOB, width="medium"),
            COL_CLASS: st.column_config.TextColumn(
                COL_CLASS, width="small",
                help="e.g. 1.0, 1.5, 2.0 ... 4.5"),
            COL_SCS: st.column_config.TextColumn(COL_SCS, width="small"),
            COL_REMARK: st.column_config.TextColumn(
                COL_REMARK, width="medium", help="Optional."),
        },
    )

    if st.button("📄 Generate PDF", type="primary"):
        players = []
        for rec in edited.to_dict("records"):
            name = clean(rec.get(COL_NAME))
            if not name:
                continue
            players.append({
                "number": rec.get(COL_NUM),
                "name": name,
                "dob": rec.get(COL_DOB),
                "sport_class": rec.get(COL_CLASS),
                "scs": rec.get(COL_SCS),
                "remark": rec.get(COL_REMARK),
            })

        if not players:
            st.session_state.pop("fr_individual_pdf", None)
            st.error("Please fill in at least one player — a player row needs a name.")
        else:
            field_values = build_individual_field_values(
                event, location, country, gender, players
            )
            reader = load_pdf_template(INDIVIDUAL_TEMPLATE)
            try:
                pdf_bytes = fill_and_get_pdf_bytes(reader, field_values)
                st.session_state["fr_individual_pdf"] = pdf_bytes
                st.session_state["fr_individual_count"] = len(players)
            except Exception as e:
                st.session_state.pop("fr_individual_pdf", None)
                st.error(f"Could not generate the PDF: {e}")

    pdf_bytes = st.session_state.get("fr_individual_pdf")
    if pdf_bytes:
        st.success(
            f"PDF generated with {st.session_state.get('fr_individual_count', 0)} player(s)."
        )
        st.download_button(
            label=f"⬇️ Download {INDIVIDUAL_FILENAME}",
            data=pdf_bytes,
            file_name=INDIVIDUAL_FILENAME,
            mime="application/pdf",
        )


st.markdown("---")
st.caption("Results Forms Generator.")
