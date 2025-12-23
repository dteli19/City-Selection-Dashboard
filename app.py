import os
import streamlit as st
import pandas as pd
from PIL import Image

# -------------------------------------------------
# Page configuration
# -------------------------------------------------
st.set_page_config(
    page_title="Excel Dashboard Viewer",
    page_icon="📊",
    layout="wide"
)

# -------------------------------------------------
# File names (must be in same directory)
# -------------------------------------------------
EXCEL_FILE = "Group 3 Dashboard.xlsm"

IMG_DASHBOARD = "Cost_Of_Living_1.png"
IMG_COST = "Cost_Of_Living_2.png"
IMG_CRIME = "Crime_Rate.png"
IMG_WEATHER = "Climate.png"

# Sheet names for preview only
SHEET_COST = "Cost_of_living"
SHEET_CRIME = "US_violent_crime"
SHEET_WEATHER = "Weather_Dataset"

# -------------------------------------------------
# Helper functions
# -------------------------------------------------
def safe_image(path):
    if os.path.exists(path):
        return Image.open(path)
    return None


def excel_download_button(file_path, key):
    if not os.path.exists(file_path):
        st.warning("Excel dashboard file not found.")
        return

    with open(file_path, "rb") as f:
        st.download_button(
            label="⬇️ Download Excel Dashboard (.xlsm)",
            data=f,
            file_name=os.path.basename(file_path),
            mime="application/vnd.ms-excel",
            use_container_width=True,
            key=key
        )


@st.cache_data(show_spinner=False)
def load_preview_sheet(xlsm_path, sheet_name, nrows=40):
    return pd.read_excel(
        xlsm_path,
        sheet_name=sheet_name,
        engine="openpyxl",
        nrows=nrows
    )

# -------------------------------------------------
# Header
# -------------------------------------------------
st.title("📊 Excel Dashboard Viewer")
st.caption(
    "This page displays the Excel dashboard visuals and provides access "
    "to the original macro-enabled workbook."
)

# Top download button (unique key)
excel_download_button(EXCEL_FILE, key="download_top")

with st.expander("How to use this dashboard"):
    st.markdown(
        """
- Use the tabs below to view screenshots of the Excel dashboard.
- To interact with slicers, dropdowns, and VBA macros, download and open the `.xlsm` file in Microsoft Excel.
"""
    )

# -------------------------------------------------
# Tabs
# -------------------------------------------------
tab1, tab2, tab3, tab4, tab5 = st.tabs(
    ["🏠 Dashboard", "💸 Cost of Living", "🚨 Crime Rate", "🌤 Climate", "📄 Data Preview"]
)

# -------------------------------------------------
# Dashboard tab
# -------------------------------------------------
with tab1:
    st.subheader("Main Dashboard View")
    img = safe_image(IMG_DASHBOARD)
    if img:
        st.image(img, use_container_width=True)
    else:
        st.info(f"Add `{IMG_DASHBOARD}` to display the dashboard image.")

# -------------------------------------------------
# Cost of Living tab
# -------------------------------------------------
with tab2:
    st.subheader("Cost of Living View")
    img = safe_image(IMG_COST)
    if img:
        st.image(img, use_container_width=True)
    else:
        st.info(f"Add `{IMG_COST}` to display the cost of living image.")

# -------------------------------------------------
# Crime Rate tab
# -------------------------------------------------
with tab3:
    st.subheader("Crime Rate View")
    img = safe_image(IMG_CRIME)
    if img:
        st.image(img, use_container_width=True)
    else:
        st.info(f"Add `{IMG_CRIME}` to display the crime rate image.")

# -------------------------------------------------
# Climate tab
# -------------------------------------------------
with tab4:
    st.subheader("Climate / Weather View")
    img = safe_image(IMG_WEATHER)
    if img:
        st.image(img, use_container_width=True)
    else:
        st.info(f"Add `{IMG_WEATHER}` to display the climate image.")


# -------------------------------------------------
# Footer
# -------------------------------------------------
st.markdown("## 💡 Key Takeaways")

st.markdown(
    """
- This Excel dashboard functions as a **decision-support tool** for evaluating post-job relocation options.
- It integrates **cost of living, crime rate, and climate** datasets to simplify complex lifestyle trade-offs.
- **Data validation and predefined expense brackets** guide user inputs and ensure consistency.
- **VBA macros** automate navigation and recalculation, improving usability and reducing manual effort.
"""
)

