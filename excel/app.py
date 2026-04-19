import streamlit as st
import pandas as pd
from fuzzywuzzy import process
from io import BytesIO

st.set_page_config(page_title="AI Excel Data Cleaner", layout="wide")

# Custom CSS Styling
st.markdown("""
<style>
.main {
    background: linear-gradient(to right, #eef2ff, #f8fafc);
}

h1 {
    color: #1e3a8a;
    text-align: center;
    font-size: 42px !important;
    font-weight: bold;
}

.stMarkdown p {
    font-size: 18px;
    color: #374151;
}

section[data-testid="stSidebar"] {
    background-color: #1e293b;
}

.stFileUploader {
    background-color: white;
    border: 2px dashed #3b82f6;
    border-radius: 15px;
    padding: 20px;
}

.stButton > button {
    background: linear-gradient(90deg, #2563eb, #7c3aed);
    color: white;
    border: none;
    border-radius: 12px;
    padding: 12px 28px;
    font-size: 16px;
    font-weight: 600;
    transition: 0.3s;
}

.stButton > button:hover {
    transform: scale(1.05);
    background: linear-gradient(90deg, #1d4ed8, #6d28d9);
}

.stDownloadButton > button {
    background: linear-gradient(90deg, #10b981, #059669);
    color: white;
    border-radius: 12px;
    padding: 12px 28px;
    font-size: 16px;
    border: none;
}

.stDataFrame {
    background: white;
    border-radius: 15px;
    padding: 10px;
    box-shadow: 0px 4px 15px rgba(0,0,0,0.08);
}

div[data-testid="stMetric"] {
    background-color: white;
    border-radius: 14px;
    padding: 15px;
    box-shadow: 0px 4px 10px rgba(0,0,0,0.08);
}

.block-container {
    padding-top: 2rem;
    padding-bottom: 2rem;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div style='text-align: center; padding: 20px 0;'>
    <h1 style='font-size: 52px; font-weight: 800; 
               background: linear-gradient(90deg, #2563eb, #7c3aed, #ec4899); 
               -webkit-background-clip: text; 
               -webkit-text-fill-color: transparent;
               margin-bottom: 10px;'>
        🤖 AI-Assisted Excel Data Cleaning Tool
    </h1>
    <p style='font-size: 20px; color: #475569; margin-top: 0;'>
        Upload messy Excel files and let AI clean, standardize, and organize your data instantly.
    </p>
</div>
""", unsafe_allow_html=True)

st.write(
    "Upload a messy Excel dataset. The tool will detect duplicates, standardize names, "
    "clean dates, correct formatting, and generate a data cleaning report."
)

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

# ------------------------------------------------
# Helper Functions
# ------------------------------------------------

def standardize_column_names(df):
    df.columns = (
        df.columns
        .str.strip()
        .str.lower()
        .str.replace(" ", "_")
    )
    return df


def clean_text_columns(df):
    for col in df.select_dtypes(include="object"):
        df[col] = df[col].astype(str).str.strip().str.title()
    return df


def clean_date_columns(df):
    date_columns = []

    for col in df.columns:
        if "date" in col.lower():
            df[col] = pd.to_datetime(
                df[col],
                errors="coerce"
            )

            df[col] = df[col].dt.strftime("%Y-%m-%d")
            date_columns.append(col)

    return df, date_columns


def handle_missing_values(df):
    missing_before = df.isna().sum().sum()

    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].fillna("Unknown")
        else:
            df[col] = df[col].fillna(df[col].median())

    missing_after = df.isna().sum().sum()
    fixed = missing_before - missing_after

    return df, fixed


def detect_duplicates(df):
    return df[df.duplicated()]


def remove_duplicates(df):
    before = len(df)
    df = df.drop_duplicates()
    after = len(df)

    removed = before - after
    return df, removed


def standardize_names(df, column, threshold=85):
    unique_names = df[column].dropna().unique().tolist()

    standardized = {}

    for name in unique_names:
        if not standardized:
            standardized[name] = name
            continue

        match, score = process.extractOne(
            name,
            list(standardized.values())
        )

        if score >= threshold:
            standardized[name] = match
        else:
            standardized[name] = name

    df[column] = df[column].map(standardized)

    return df, len(unique_names)


def convert_df_to_excel(df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)

    return output.getvalue()


# ------------------------------------------------
# Main Processing
# ------------------------------------------------

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.subheader("📊 Raw Data Preview")
    st.dataframe(df)

    original_rows = len(df)

    # Step 1: Standardize column names
    df = standardize_column_names(df)

    # Step 2: Clean text columns
    df = clean_text_columns(df)

    # Step 3: Standardize date columns
    df, date_columns = clean_date_columns(df)

    if date_columns:
        st.success(f"Date columns standardized: {date_columns}")

    # Step 4: Handle missing values
    df, missing_fixed = handle_missing_values(df)

    # Step 5: Detect duplicates
    duplicates = detect_duplicates(df)

    if not duplicates.empty:
        st.warning(f"{len(duplicates)} duplicate rows detected")
        st.dataframe(duplicates)

    # Step 6: Remove duplicates
    df, removed = remove_duplicates(df)
    st.success(f"Removed {removed} duplicate rows")

    # Step 7: Name standardization
    name_columns = df.select_dtypes(include="object").columns.tolist()

    names_standardized = 0

    if name_columns:
        selected_col = st.selectbox(
            "Select column for name standardization",
            name_columns
        )

        if st.button("Standardize Names"):
            df, names_standardized = standardize_names(df, selected_col)
            st.success("Names standardized successfully")

    # Step 8: Show cleaned data
    st.subheader("✅ Cleaned Data Preview")
    st.dataframe(df)

    # Step 9: Cleaning report
    st.subheader("📑 Data Cleaning Report")

    report = {
        "Original Rows": original_rows,
        "Final Rows": len(df),
        "Duplicates Removed": removed,
        "Missing Values Fixed": missing_fixed,
        "Unique Names Processed": names_standardized,
        "Date Columns Standardized": date_columns
    }

    st.json(report)

    # Step 10: Download cleaned file
    excel_file = convert_df_to_excel(df)

    st.download_button(
        label="📥 Download Cleaned Excel",
        data=excel_file,
        file_name="cleaned_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
