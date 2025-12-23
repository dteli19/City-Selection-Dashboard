import os
import streamlit as st
import pandas as pd
from PIL import Image

# ----------------------------
# Page config
# ----------------------------
st.set_page_config(
    page_title="City Selection Dashboard (Excel to App)",
    page_icon="📊",
    layout="wide"
)

# ----------------------------
# File names (must exist in same folder)
# ----------------------------
DATA_FILE = "Group 3 Dashboard.xlsm"

IMG_CLIMATE = "Climate.png"
IMG_COL1 = "Cost_Of_Living_1.png"
IMG_COL2 = "Cost_Of_Living_2.png"
IMG_CRIME = "Crime_Rate.png"

SHEET_COST = "Cost_of_living"
SHEET_CRIME = "US_violent_crime"
SHEET_WEATHER = "Weather_Dataset"
SHEET_CONCISE = "Concised Table"


# ----------------------------
# Helpers
# ----------------------------
def safe_open_image(path: str):
    if os.path.exists(path):
        return Image.open(path)
    return None


def excel_download_button(file_path: str, label: str):
    """Adds a Streamlit download button for the Excel file."""
    if not os.path.exists(file_path):
        st.warning(f"Excel file not found: {file_path}")
        return

    with open(file_path, "rb") as f:
        st.download_button(
            label=label,
            data=f,
            file_name=os.path.basename(file_path),
            mime="application/vnd.ms-excel"
        )


@st.cache_data(show_spinner=False)
def load_excel_data(xlsm_path: str):
    cost_df = pd.read_excel(xlsm_path, sheet_name=SHEET_COST, engine="openpyxl")
    crime_df = pd.read_excel(xlsm_path, sheet_name=SHEET_CRIME, engine="openpyxl")
    weather_df = pd.read_excel(xlsm_path, sheet_name=SHEET_WEATHER, engine="openpyxl")
    concise_df_raw = pd.read_excel(xlsm_path, sheet_name=SHEET_CONCISE, engine="openpyxl")

    # Concised Table: keep the first block of columns used for bracket filtering
    keep_cols = ["Rank", "State Name", "Grocery", "Housing", "Utilities",
                 "Transportation", "Health", "Misc.", "Total"]
    concise_df = concise_df_raw.loc[:, keep_cols].copy()

    # Cleanup
    concise_df["State Name"] = concise_df["State Name"].astype(str).str.strip()
    cost_df["State Name"] = cost_df["State Name"].astype(str).str.strip()
    crime_df["State"] = crime_df["State"].astype(str).str.strip()
    weather_df["State"] = weather_df["State"].astype(str).str.strip()

    # Rename for clarity in the app output
    cost_df = cost_df.rename(columns={"Rank": "Cost Rank", "Total": "Estimated Expense"})
    crime_df = crime_df.rename(columns={"Rank": "Crime Rank", "Total": "Total Arrests"})
    weather_df = weather_df.rename(columns={
        "Yearly Avg Temp": "Yearly Avg Temp (°F)",
        "Climate Type": "Climate Condition"
    })

    return cost_df, crime_df, weather_df, concise_df


def get_bracket_options(concise_df: pd.DataFrame, col: str):
    opts = sorted([x for x in concise_df[col].dropna().unique().tolist() if str(x).strip() != ""])
    return [str(x) for x in opts]


def filter_by_brackets(concise_df: pd.DataFrame, selections: dict):
    df = concise_df.copy()
    for k, v in selections.items():
        if v != "Any":
            df = df[df[k].astype(str) == str(v)]
    return df


# ----------------------------
# Validate Excel presence
# ----------------------------
if not os.path.exists(DATA_FILE):
    st.error(
        f"Could not find '{DATA_FILE}' in the current folder.\n\n"
        "Place `Group 3 Dashboard.xlsm` in the same directory as `app.py`."
    )
    st.stop()

# Load data
cost_df, crime_df, weather_df, concise_df = load_excel_data(DATA_FILE)

# Load images (optional)
img_climate = safe_open_image(IMG_CLIMATE)
img_col1 = safe_open_image(IMG_COL1)
img_col2 = safe_open_image(IMG_COL2)
img_crime = safe_open_image(IMG_CRIME)

# ----------------------------
# Sidebar
# ----------------------------
st.sidebar.header("Navigation")
page = st.sidebar.radio(
    "Go to",
    ["Home (Final Recommendation)", "Cost of Living", "Crime Rate", "Weather", "Data (Raw Tables)"]
)

# Sidebar download (always visible)
st.sidebar.markdown("---")
st.sidebar.subheader("📥 Download")
excel_download_button(DATA_FILE, "Download Excel Dashboard (.xlsm)")

# ----------------------------
# Main Title
# ----------------------------
st.title("📊 City Selection Decision Dashboard")

# ----------------------------
# Home
# ----------------------------
if page == "Home (Final Recommendation)":
    st.subheader("Where to go next? Let the data decide.")

    left, right = st.columns([1.2, 1])

    with left:
        st.markdown(
            """
This app replicates the Excel dashboard workflow:
- Select **expense brackets** you are comfortable with (Grocery, Housing, Utilities, etc.)
- The app filters states using the **Concised Table** bracket logic
- Results are enriched with **Cost of Living**, **Crime**, and **Weather** datasets
"""
        )

        bracket_cols = ["Grocery", "Housing", "Utilities", "Transportation", "Health", "Misc.", "Total"]
        selections = {}

        cols_grid = st.columns(3)
        for i, col in enumerate(bracket_cols):
            opts = ["Any"] + get_bracket_options(concise_df, col)
            with cols_grid[i % 3]:
                selections[col] = st.selectbox(f"{col} bracket", options=opts, index=0, key=f"sel_{col}")

        st.markdown("### Optional preferences")
        p1, p2, p3 = st.columns(3)

        climate_opts = ["Any"] + sorted(weather_df["Climate Condition"].dropna().unique().tolist())
        with p1:
            chosen_climate = st.selectbox("Preferred Climate Condition", climate_opts, index=0)

        with p2:
            max_crime_rank = st.slider("Max Crime Rank (lower is better)", min_value=1, max_value=50, value=50)

        with p3:
            max_cost_rank = st.slider("Max Cost Rank (lower is cheaper)", min_value=1, max_value=50, value=50)

        filtered_brackets = filter_by_brackets(concise_df, selections)

        merged = (
            filtered_brackets
            .merge(cost_df[["State Name", "Cost Rank", "Estimated Expense"]], on="State Name", how="left")
            .merge(crime_df[["State", "Crime Rank", "Total Arrests"]], left_on="State Name", right_on="State", how="left")
            .merge(weather_df[["State", "Climate Condition", "Yearly Avg Temp (°F)"]], left_on="State Name", right_on="State", how="left")
        )

        # Apply optional filters
        if chosen_climate != "Any":
            merged = merged[merged["Climate Condition"].astype(str) == str(chosen_climate)]

        merged["Crime Rank"] = pd.to_numeric(merged["Crime Rank"], errors="coerce")
        merged["Cost Rank"] = pd.to_numeric(merged["Cost Rank"], errors="coerce")
        merged = merged[merged["Crime Rank"] <= max_crime_rank]
        merged = merged[merged["Cost Rank"] <= max_cost_rank]

        show_cols = [
            "State Name",
            "Cost Rank",
            "Estimated Expense",
            "Crime Rank",
            "Total Arrests",
            "Climate Condition",
            "Yearly Avg Temp (°F)",
        ]

        st.markdown("### Suggested states based on your selections")
        if merged.empty:
            st.warning("No states match your selections. Try setting some fields to 'Any'.")
        else:
            merged_out = merged.sort_values(by=["Cost Rank", "Crime Rank"], ascending=True)

            best_state = merged_out["State Name"].iloc[0]
            k1, k2, k3 = st.columns(3)
            with k1:
                st.metric("Top Recommendation", best_state)
            with k2:
                st.metric("Matching States", int(merged_out.shape[0]))
            with k3:
                st.metric("Selected Climate", chosen_climate if chosen_climate != "Any" else "No preference")

            st.dataframe(merged_out[show_cols], use_container_width=True)

        # Download section on Home page too
        st.markdown("---")
        st.subheader("⬇️ Download the Excel Dashboard")
        st.write("Download the original macro-enabled Excel file to view VBA macros, data validation, and sheet navigation.")
        excel_download_button(DATA_FILE, "Download Excel Dashboard (.xlsm)")

    with right:
        st.markdown("### Excel Screenshots")
        if img_col1:
            st.image(img_col1, caption="Excel Dashboard Home View", use_container_width=True)
        if img_col2:
            st.image(img_col2, caption="Cost of Living View (Excel)", use_container_width=True)

# ----------------------------
# Cost of Living
# ----------------------------
elif page == "Cost of Living":
    st.subheader("💸 Cost of Living")
    c1, c2 = st.columns([1, 1])
    with c1:
        if img_col2:
            st.image(img_col2, caption="Cost of Living Index View (Excel)", use_container_width=True)
    with c2:
        cols = ["State Name", "Cost Rank", "Estimated Expense", "Grocery", "Housing", "Utilities",
                "Transportation", "Health", "Misc."]
        available = [c for c in cols if c in cost_df.columns]
        st.dataframe(cost_df[available].sort_values("Cost Rank"), use_container_width=True)

# ----------------------------
# Crime Rate
# ----------------------------
elif page == "Crime Rate":
    st.subheader("🚨 Crime Rate")
    c1, c2 = st.columns([1, 1])
    with c1:
        if img_crime:
            st.image(img_crime, caption="Crime Rate View (Excel)", use_container_width=True)
    with c2:
        cols = ["State", "Crime Rank", "Total Arrests", "Murder", "Assault", "Rape"]
        available = [c for c in cols if c in crime_df.columns]
        st.dataframe(crime_df[available].sort_values("Crime Rank"), use_container_width=True)

# ----------------------------
# Weather
# ----------------------------
elif page == "Weather":
    st.subheader("🌤 Weather and Climate")
    c1, c2 = st.columns([1, 1])
    with c1:
        if img_climate:
            st.image(img_climate, caption="Weather View (Excel)", use_container_width=True)
    with c2:
        cols = ["State", "Spring (Avg ° F)", "Summer (Avg ° F)", "Fall (Avg ° F)", "Winter (Avg ° F)",
                "Yearly Avg Temp (°F)", "Climate Condition"]
        available = [c for c in cols if c in weather_df.columns]
        st.dataframe(weather_df[available].sort_values("State"), use_container_width=True)

# ----------------------------
# Raw Tables
# ----------------------------
else:
    st.subheader("📄 Raw Tables")
    tab1, tab2, tab3, tab4 = st.tabs(
        ["Concised Table (Brackets)", "Cost_of_living", "US_violent_crime", "Weather_Dataset"]
    )
    with tab1:
        st.dataframe(concise_df.sort_values("Rank"), use_container_width=True)
    with tab2:
        st.dataframe(cost_df, use_container_width=True)
    with tab3:
        st.dataframe(crime_df, use_container_width=True)
    with tab4:
        st.dataframe(weather_df, use_container_width=True)
