import streamlit as st

# Single, app-wide page configuration.
# With st.navigation, set_page_config must be called exactly ONCE, here in the
# router — never inside the individual tool/page scripts (it would raise
# StreamlitSetPageConfigMustBeFirstCommandError).
st.set_page_config(page_title="IWBF Classification Tools", page_icon="🏀", layout="centered")

# --- Pages (one entry per tool) ---
home = st.Page(
    "home/home.py",
    title="Home",
    icon="🏠",
    default=True,  # shown at the app root on first load
)
assessment_forms = st.Page(
    "tools/assessment_forms/assessment_forms.py",
    title="Assessment Forms Generator",
    icon="📄",
    url_path="assessment-forms",
)
card_merger = st.Page(
    "tools/card_merger/card_merger.py",
    title="Player Card Merger",
    icon="🪪",
    url_path="card-merger",
)
final_results = st.Page(
    "tools/final_results/final_results.py",
    title="Results Forms Generator",
    icon="📝",
    url_path="final-results",
)

# The order of this list is the order shown in the sidebar menu.
#
# To add a new tool:
#   1) create tools/<name>/<name>.py (a normal Streamlit script, NO set_page_config;
#      load any bundled files via Path(__file__).resolve().parent / "assets")
#   2) add one st.Page(...) line above and include it in the list below.
pg = st.navigation([home, assessment_forms, card_merger, final_results])
pg.run()
