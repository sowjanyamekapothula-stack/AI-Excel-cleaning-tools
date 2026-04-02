import streamlit as st
import pandas as pd
from io import BytesIO

# Optional fuzzy matching
try:
    from fuzzywuzzy import process
    FUZZY_AVAILABLE = True
except:
    FUZZY_AVAILABLE = False

st.set_page_config(page_title="AI Excel Data Cleaner", layout="wide")

st.title("🤖 AI-Assisted Excel Data Cleaning Tool")

uploaded_file = st.file_uploader("📂 Upload Excel File", type=["xlsx", "xls"])

if uploaded_file is None:
    st.info("👆 Please upload a file to start")
    st.stop()

# -----------------------------
# READ FILE
# -----------------------------
df = pd.read_excel(uploaded_file)

st.subheader("📊 Raw Data Preview")
st.dataframe(df)

original_rows = len(df)

# -----------------------------
# CLEAN COLUMN NAMES
# -----------------------------
df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")

# -----------------------------
# STEP 1: DETECT NUMERIC COLUMNS
# -----------------------------
numeric_columns = []

for col in df.columns:
    converted = pd.to_numeric(df[col], errors="coerce")
    if converted.notna().sum() / len(df) > 0.8:
        df[col] = converted
        numeric_columns.append(col)

# -----------------------------
# 🔥 STEP 2: DETECT & FIX DATES FIRST (IMPORTANT)
# -----------------------------
date_columns = []

for col in df.columns:

    if col in numeric_columns:
        continue

    temp = df[col].astype(str).str.strip()

    # Try multiple parsing strategies
    dt1 = pd.to_datetime(temp, errors="coerce", dayfirst=True)
    dt2 = pd.to_datetime(temp, errors="coerce", format="mixed")

    # Combine results
    final = dt1.fillna(dt2)

    # If it's mostly dates → confirm column
    if final.notna().sum() > len(df) * 0.5:

        # ✅ FORCE FORMAT
        df[col] = final.dt.strftime("%Y-%m-%d")

        date_columns.append(col)

# -----------------------------
# STEP 3: CLEAN TEXT (AFTER DATE FIX)
# -----------------------------
for col in df.select_dtypes(include="object"):
    if col not in numeric_columns and col not in date_columns:
        df[col] = df[col].astype(str).str.strip().str.title()

# -----------------------------
# STEP 4: HANDLE MISSING VALUES
# -----------------------------
missing_before = df.isna().sum().sum()

for col in df.columns:
    if col in numeric_columns:
        df[col] = df[col].fillna(df[col].median())
    else:
        df[col] = df[col].fillna("Unknown")

missing_after = df.isna().sum().sum()
missing_fixed = missing_before - missing_after

# -----------------------------
# STEP 5: REMOVE DUPLICATES
# -----------------------------
duplicates = df[df.duplicated()]

if not duplicates.empty:
    st.warning(f"⚠️ {len(duplicates)} duplicate rows found")
    st.dataframe(duplicates)

before = len(df)
df = df.drop_duplicates()
removed = before - len(df)

st.success(f"✅ Removed {removed} duplicate rows")

# -----------------------------
# STEP 6: NAME STANDARDIZATION
# -----------------------------
names_standardized = 0

if FUZZY_AVAILABLE:
    text_cols = df.select_dtypes(include="object").columns.tolist()

    if text_cols:
        selected_col = st.selectbox("Select column for name standardization", text_cols)

        if st.button("Standardize Names"):
            unique_vals = df[selected_col].dropna().unique().tolist()
            mapping = {}

            for val in unique_vals:
                if not mapping:
                    mapping[val] = val
                    continue

                match, score = process.extractOne(val, list(mapping.values()))
                mapping[val] = match if score > 85 else val

            df[selected_col] = df[selected_col].map(mapping)
            names_standardized = len(unique_vals)

            st.success("✅ Names standardized")

# -----------------------------
# OUTPUT
# -----------------------------
st.subheader("✅ Cleaned Data Preview")
st.dataframe(df)

st.subheader("📑 Data Cleaning Report")

st.json({
    "Original Rows": original_rows,
    "Final Rows": len(df),
    "Duplicates Removed": removed,
    "Missing Values Fixed": missing_fixed,
    "Date Columns Standardized": date_columns,
    "Numeric Columns": numeric_columns,
    "Names Standardized": names_standardized
})

# -----------------------------
# DOWNLOAD
# -----------------------------
output = BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df.to_excel(writer, index=False)

st.download_button(
    "📥 Download Cleaned Excel",
    data=output.getvalue(),
    file_name="cleaned_data.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)