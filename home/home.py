import streamlit as st

st.title("🏀 IWBF Classification Tools")

st.markdown(
    """
Welcome! This hub brings together the IWBF classification tools in one place.
Use the **sidebar menu** on the left to open a tool, or click a button below.
"""
)

st.markdown("---")

st.subheader("📄 Player Assessment Forms Generator")
st.write(
    "Generate player assessment forms (PDF) automatically from a `Players.xlsx` spreadsheet."
)
st.page_link(
    "tools/assessment_forms/assessment_forms.py",
    label="Open the Assessment Forms Generator",
    icon="📄",
)

st.markdown("")

st.subheader("🪪 Player Card Merger")
st.write(
    "Merge multiple player card PDFs into a single, print-ready sheet "
    "(A4 layout or a business-card template)."
)
st.page_link(
    "tools/card_merger/card_merger.py",
    label="Open the Player Card Merger",
    icon="🪪",
)

st.markdown("")

st.subheader("📝 Results Forms Generator")
st.write(
    "Generate the Classification Results forms — **Stage 2** and **Final** from a "
    "spreadsheet (one tab per team, one PDF per team in a ZIP), or fill an "
    "**Individual** form right in the browser."
)
st.page_link(
    "tools/final_results/final_results.py",
    label="Open the Results Forms Generator",
    icon="📝",
)

st.markdown("---")
st.caption("More tools coming soon.")
