import os
import streamlit as st
import pandas as pd
from PIL import Image

# ----------------------------
# Config
# ----------------------------
st.set_page_config(page_title="Excel Dashboard Viewer", page_icon="📊", layout="wide")

EXCEL_FILE = "Group 3 Dashboard.xlsm"

# Screenshots (put these files in the same folder as app.py)
IMG_DASHBOARD = "Cost_Of_Living_1.png"   # your main dashboard screenshot
IMG_COST = "Cost_Of_Living_2.png"
IMG_CRIME = "Crime_Rate.png"
IMG_WEATHER = "Climate.png"

# Sheet names (only used for optional preview)
SHEET_COST = "Cost_of_living"
SHEET_CRIME = "US_violent_crime"
SHEET_WEATHER = "Weather_Dataset"


def safe_image(path: str):
    if os.path.exists(path):
        return Image.open(path)
    return None


def excel_download_button(file_path: str):
    if not os.path.exists(file_path):
        st.warning(f"Excel file not found: {file_path}")
        return
    with open(file_path, "rb") as f:
        st.download_button(
            label="⬇️ Download Excel Dashboard (.xlsm)",
            data=f,
            file_name=os.path.basename(file_path),
            mime="application/vnd.ms-excel",
            use_container_width=True
        )


@st.cache_data(show_spinner=False)
def load_sheet_preview(xlsm_path: str, sheet_name: str, nrows: int = 50):
    return pd.read_excel(xlsm_path, sheet_name=sheet_name, engine="openpyxl", nrows=nrows)


# ----------------------------
# Header
# ----------------------------
st.title("📊 Excel Dashboard Viewer")
st.caption("This web page displays the Excel dashboard screenshots and provides the original .xlsm for download.")

# Download button at top
excel_download_button(EXCEL_FILE)

# Quick note (kept honest)
with st.expander("How to use"):
    st.markdown(
        """
- Use the tabs below to view the Excel dashboard screens.
- To interact with slicers, buttons, and VBA macros, download and open the `.xlsm` in Microsoft Excel.
"""
    )

# ----------------------------
# Tabs (viewer only)
# ----------------------------
tab1, tab2, tab3, tab4, tab5 = st.tabs(
    ["🏠 Dashboard", "💸 Cost of Living", "🚨 Crime Rate", "🌤 Weather", "📄 Data Preview (Optional)"]
)

with tab1:
    st.subheader("Dashboard")
    img = safe_image(IMG_DASHBOARD)
    if img:
        st.image(img, use_container_width=True)
    else:
        st.info(f"Add `{IMG_DASHBOARD}` in the same folder to display the dashboard screenshot aware.")

with tab2:
    st.subheader("Cost of Living")
    img = safe_image(IMG_COST)
    if img:
        st.image(img, use_container_width=True)
    else:
        st.info(f"Add `{IMG_COST}` in the same folder to display the cost screenshot.")

with tab3:
    st.subheader("Crime Rate")
    img = safe_image(IMG_CRIME)
    if img:
        st.image(img, use_container_width=True)
    else:
        st.info(f"Add `{IMG_CRIME}` in the same folder to display the crime screenshot.")

with tab4:
    st.subheader("Weather")
    img = safe_image(IMG_WEATHER)
    if img:
        st.image(img, use_container_width=True)
    else:
        st.info(f"Add `{IMG_WEATHER}` in the same folder to display the weather screenshot.")

with tab5:
    st.subheader("Data Preview (Optional)")
    st.caption("This section previews a few rows from the Excel sheets. It does not replicate macros or slicers.")

    if not os.path.exists(EXCEL_FILE):
        st.warning("Excel file not found, cannot show previews.")
    else:
        c1, c2 = st.columns(2)

        with c1:
            st.markdown("#### Cost_of_living (sample)")
            try:
                df_cost = load_sheet_preview(EXCEL_FILE, SHEET_COST, nrows=40)
                st.dataframe(df_cost, use_container_width=True)
            except Exception as e:
                st.error(f"Could not load sheet `{SHEET_COST}`: {e}")

        with c2:
            st.markdown("#### US_violent_crime (sample)")
            try:
                df_crime = load_sheet_preview(EXCEL_FILE, SHEET_CRIME, nrows=40)
                st.dataframe(df_crime, use_container_width=True)
            except Exception as e:
                st.error(f"Could not load sheet `{SHEET_CRIME}`: {e}")

        st.markdown("#### Weather_Dataset (sample)")
        try:
            df_weather = load_sheet_preview(EXCEL_FILE, SHEET_WEATHER, nrows=40)
            st.dataframe(df_weather, use_container_width=True)
        except Exception as e:
            st.error(f"Could not load sheet `{SHEET_WEATHER}`: {e}")

# Footer download button
st.markdown("---")
excel_download_button(EXCEL_FILE)
st.caption("For full interactivity, open the downloaded .xlsm in Microsoft Excel.")
