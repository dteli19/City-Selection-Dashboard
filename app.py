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
# Helpers
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


def safe_open_image(path: str):
    if os.path.exists(path):
        return Image.open(path)
    return None


@st.cache_data(show_spinner=False)
def load_excel_data(xlsm_path: str):
    # Read sheets with pandas (openpyxl engine supports .xlsm reading)
    cost_df = pd.read_excel(xlsm_path, sheet_name=SHEET_COST, engine="openpyxl")
    crime_df = pd.read_excel(xlsm_path, sheet_name=SHEET_CRIME, engine="openpyxl")
    weather_df = pd.read_excel(xlsm_path, sheet_name=SHEET_WEATHER, engine="openpyxl")
    concise_df_raw = pd.read_excel(xlsm_path, sheet_name=SHEET_CONCISE, engine="openpyxl")

    # The "Concised Table" has duplicated columns in your workbook.
    # Keep the first set: Rank, State Name, Grocery, Housing, Utilities, Transportation, Health, Misc., Total
    keep_cols = ["Rank", "State Name", "Grocery", "Housing", "Utilities", "Transportation", "Health", "Misc.", "Total"]
    concise_df = concise_df_raw.loc[:, keep_cols].copy()

    # Clean types
    for c in ["Rank"]:
        if c in concise_df.columns:
            concise_df[c] = pd.to_numeric(concise_df[c], errors="coerce")

    # Normalize state name fields for merges
    cost_df["State Name"] = cost_df["State Name"].astype(str).str.strip()
    crime_df["State"] = crime_df["State"].astype(str).str.strip()
    weather_df["State"] = weather_df["State"].astype(str).str.strip()
    concise_df["State Name"] = concise_df["State Name"].astype(str).str.strip()

    # Rename some columns for clarity in the merged output
    cost_df = cost_df.rename(columns={
        "Rank": "Cost Rank",
        "Total": "Estimated Expense"
    })

    crime_df = crime_df.rename(columns={
        "Total": "Total Arrests",
        "Rank": "Crime Rank"
    })

    weather_df = weather_df.rename(columns={
        "Yearly Avg Temp": "Yearly Avg Temp (°F)",
        "Climate Type": "Climate Condition"
    })

    return cost_df, crime_df, weather_df, concise_df


def get_bracket_options(concise_df: pd.DataFrame, col: str):
    opts = sorted([x for x in concise_df[col].dropna().unique().tolist() if str(x).strip() != ""])
    # Keep as strings
    return [str(x) for x in opts]


def filter_by_brackets(concise_df: pd.DataFrame, selections: dict):
    df = concise_df.copy()
    for k, v in selections.items():
        if v != "Any":
            df = df[df[k].astype(str) == str(v)]
    return df


# ----------------------------
# Load data
# ----------------------------
st.title("📊 City Selection Decision Dashboard (Excel VBA to Streamlit)")

if not os.path.exists(DATA_FILE):
    st.error(
        f"Could not find '{DATA_FILE}' in the current folder.\n\n"
        "Place `Group 3 Dashboard.xlsm` in the same directory as `app.py`."
    )
    st.stop()

cost_df, crime_df, weather_df, concise_df = load_excel_data(DATA_FILE)

# Images
img_climate = safe_open_image(IMG_CLIMATE)
img_col1 = safe_open_image(IMG_COL1)
img_col2 = safe_open_image(IMG_COL2)
img_crime = safe_open_image(IMG_CRIME)

# ----------------------------
# Sidebar Navigation
# ----------------------------
st.sidebar.header("Navigation")
page = st.sidebar.radio(
    "Go to",
    ["Home (Final Recommendation)", "Cost of Living", "Crime Rate", "Weather", "Data (Raw Tables)"]
)

# ----------------------------
# Home: Decision Workflow
# ----------------------------
if page == "Home (Final Recommendation)":
    st.subheader("Where to go next? Let the data decide.")

    left, right = st.columns([1.2, 1])

    with left:
        st.markdown(
            """
This app mirrors the Excel dashboard workflow:
- You select **expense brackets** you are comfortable with (Grocery, Housing, Utilities, etc.)
- The model filters states using the **Concised Table** brackets
- Results are enriched with **Cost of Living**, **Crime**, and **Weather** data
"""
        )

        # Bracket dropdowns (same idea as Excel validation lists)
        bracket_cols = ["Grocery", "Housing", "Utilities", "Transportation", "Health", "Misc.", "Total"]

        selections = {}
        cols1 = st.columns(3)
        for i, col in enumerate(bracket_cols):
            opts = ["Any"] + get_bracket_options(concise_df, col)
            with cols1[i % 3]:
                selections[col] = st.selectbox(f"{col} bracket", options=opts, index=0, key=f"sel_{col}")

        # Additional preference filters (optional)
        st.markdown("### Optional preferences")
        pref1, pref2, pref3 = st.columns(3)

        climate_opts = ["Any"] + sorted(weather_df["Climate Condition"].dropna().unique().tolist())
        with pref1:
            chosen_climate = st.selectbox("Preferred Climate Condition", climate_opts, index=0)

        with pref2:
            max_crime_rank = st.slider("Max Crime Rank (lower is better)", min_value=1, max_value=50, value=50)

        with pref3:
            max_cost_rank = st.slider("Max Cost Rank (lower is cheaper)", min_value=1, max_value=50, value=50)

        # Filter by bracket table
        filtered_brackets = filter_by_brackets(concise_df, selections)

        # Merge with other datasets
        merged = (
            filtered_brackets
            .merge(cost_df[["State Name", "Cost Rank", "Estimated Expense"]], on="State Name", how="left")
            .merge(crime_df[["State", "Crime Rank", "Total Arrests"]], left_on="State Name", right_on="State", how="left")
            .merge(weather_df[["State", "Climate Condition", "Yearly Avg Temp (°F)"]], left_on="State Name", right_on="State", how="left")
        )

        # Apply optional filters
        if chosen_climate != "Any":
            merged = merged[merged["Climate Condition"].astype(str) == str(chosen_climate)]

        merged = merged[pd.to_numeric(merged["Crime Rank"], errors="coerce") <= max_crime_rank]
        merged = merged[pd.to_numeric(merged["Cost Rank"], errors="coerce") <= max_cost_rank]

        # Clean columns
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
        st.caption("Tip: Adjust brackets and filters to see updated recommendations.")

        if merged.empty:
            st.warning("No states match your current bracket selections and filters. Try setting some fields to 'Any'.")
        else:
            # Sort for readability: cheaper and safer first
            merged_out = merged.copy()
            merged_out["Cost Rank"] = pd.to_numeric(merged_out["Cost Rank"], errors="coerce")
            merged_out["Crime Rank"] = pd.to_numeric(merged_out["Crime Rank"], errors="coerce")
            merged_out = merged_out.sort_values(by=["Cost Rank", "Crime Rank"], ascending=True)

            # KPI style
            best_state = merged_out["State Name"].iloc[0]
            k1, k2, k3 = st.columns(3)
            with k1:
                st.metric("Top Recommendation", best_state)
            with k2:
                st.metric("Matching States", int(merged_out.shape[0]))
            with k3:
                st.metric("Selected Climate", chosen_climate if chosen_climate != "Any" else "No preference")

            st.dataframe(merged_out[show_cols], use_container_width=True)

    with right:
        st.markdown("### Dashboard Screens (from Excel)")
        if img_col1:
            st.image(img_col1, caption="Excel Dashboard Home View", use_container_width=True)
        if img_col2:
            st.image(img_col2, caption="Cost of Living View (Excel)", use_container_width=True)

        st.info(
            "Note: This Streamlit app replicates the Excel decision flow using the same underlying datasets and bracket logic."
        )

# ----------------------------
# Cost of Living page
# ----------------------------
elif page == "Cost of Living":
    st.subheader("💸 Cost of Living")
    st.write("This section displays the Cost of Living dataset used by the dashboard.")

    c1, c2 = st.columns([1, 1])
    with c1:
        if img_col2:
            st.image(img_col2, caption="Cost of Living Index View (Excel)", use_container_width=True)
    with c2:
        st.dataframe(
            cost_df[["State Name", "Cost Rank", "Estimated Expense", "Grocery", "Housing", "Utilities",
                     "Transportation", "Health", "Misc.", ]].sort_values("Cost Rank"),
            use_container_width=True
        )

# ----------------------------
# Crime page
# ----------------------------
elif page == "Crime Rate":
    st.subheader("🚨 Crime Rate")
    st.write("This section displays violent crime metrics used as decision inputs.")

    c1, c2 = st.columns([1, 1])
    with c1:
        if img_crime:
            st.image(img_crime, caption="Crime Rate View (Excel)", use_container_width=True)
    with c2:
        st.dataframe(
            crime_df[["State", "Crime Rank", "Total Arrests", "Murder", "Assault", "Rape"]]
            .sort_values("Crime Rank"),
            use_container_width=True
        )

# ----------------------------
# Weather page
# ----------------------------
elif page == "Weather":
    st.subheader("🌤 Weather and Climate")
    st.write("This section displays seasonal temperature breakdown and climate condition by state.")

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
# Raw Data
# ----------------------------
else:
    st.subheader("📄 Raw Tables")
    st.caption("These are the source tables powering the dashboard logic.")

    tab1, tab2, tab3, tab4 = st.tabs(["Concised Table (Brackets)", "Cost_of_living", "US_violent_crime", "Weather_Dataset"])

    with tab1:
        st.dataframe(concise_df.sort_values("Rank"), use_container_width=True)
    with tab2:
        st.dataframe(cost_df, use_container_width=True)
    with tab3:
        st.dataframe(crime_df, use_container_width=True)
    with tab4:
        st.dataframe(weather_df, use_container_width=True)
