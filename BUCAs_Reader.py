import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(page_title="Business Case Comparison", layout="wide")

st.title("📊 Business Case Financial Comparison")

# Relevant sheets to analyze
relevant_sheets = [
    "Business Case",
    "1. Kosten",
    "2. Inv.& Afschrijvingen",
    "3. Centrale Middelen",
    "4. Opbrengsten",
    "4.b Directe inkoop Marge",
    "Opbrengsten per instelling",
    "Budget SAM",
    "Budget GL"
]

def read_clean_sheet(excel_data, sheet_name):
    """Read Excel sheet, auto-detect header row, drop empty rows/cols."""
    raw_df = excel_data.parse(sheet_name, header=None)

    # Find the first row that has at least 2 non-empty cells (likely header)
    header_row = raw_df.notna().sum(axis=1).idxmax()

    df = excel_data.parse(sheet_name, header=header_row)
    df = df.dropna(how="all").dropna(axis=1, how="all")
    return df

def split_by_keywords(df, keywords, col_idx=1):
    """Split dataframe into multiple sub-dataframes based on keywords in a column (default: column B)."""
    split_dfs = {}
    if col_idx >= len(df.columns):
        return split_dfs

    col = df.columns[col_idx]
    for keyword in keywords:
        mask = df[col].astype(str).str.contains(keyword, case=False, na=False)
        sub_df = df[mask].copy()
        if not sub_df.empty:
            split_dfs[keyword] = sub_df
    return split_dfs

def extract_years(excel_data):
    """Extract the 10-year period based on Uitgangspunten!15C."""
    try:
        uitgangspunten = excel_data.parse("Uitgangspunten", header=None)
        start_year = int(uitgangspunten.iloc[14, 2])  # row 15C -> (14,2) zero-indexed
        years = list(range(start_year - 2, start_year + 8))  # 10 years
        return years
    except Exception:
        return None


def read_business_case(excel_data):
    """Read Business Case sheet, align years, return DataFrame."""
    df = excel_data.parse("Business Case", header=None)
    years = extract_years(excel_data)

    if not years:
        return None, None

    # Look for year blocks in row 4 and 24
    header_rows = [3, 23]
    dataframes = []
    for row in header_rows:
        year_row = df.iloc[row, :].dropna().tolist()
        if any(str(y) in str(v) for v in year_row for y in years):
            data = df.iloc[row+1:row+15, :]  # take ~15 rows under each block
            data.columns = df.iloc[row, :]
            dataframes.append(data)

    if not dataframes:
        return None, None

    st.write(dataframes)
    combined = pd.concat(dataframes, axis=0)
    year_cols = [c for c in combined.columns if str(c).isdigit() and int(c) in years]
    combined = combined[year_cols]
    combined = combined.dropna(axis=1, how="all")
    return combined, years


uploaded_files = st.file_uploader(
    "Upload Excel files",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:
    sheet_name = st.selectbox("Select a financial sheet:", relevant_sheets)

    dfs = []
    for uploaded_file in uploaded_files:
        try:
            excel_data = pd.ExcelFile(uploaded_file)
            if sheet_name not in excel_data.sheet_names:
                st.warning(f"Sheet '{sheet_name}' not found in {uploaded_file.name}")
                continue

            df = read_clean_sheet(excel_data, sheet_name)
            dfs.append((uploaded_file.name, df, excel_data))
        except Exception as e:
            st.error(f"Error reading {uploaded_file.name}: {e}")

    # Show individual dataframes
    st.subheader("📄 Individual Files")
    for name, df, _ in dfs:

        #st.markdown(f"**File: {name}**")
        with st.expander(name):

            st.dataframe(df, use_container_width=True)

    # Special handling for Kosten
    if sheet_name == "1. Kosten" and dfs:
        st.subheader("👥 Aggregated FTEs per Function Type")

        # Initialize an empty list to store the new DataFrames
        dfs = []

        # Initialize an empty list to store the current DataFrame
        current_df = []
        
        for index, row in df.iterrows():
            # Check if the row contains the sequence of years
            print(row)
            if row.iloc[4] == 2022 and row.iloc[5] == 2023:
            # If the current_df is not empty, append it to the dfs list
                st.write("i found it")

                if current_df:
                    dfs.append(pd.DataFrame(current_df))
                    current_df = []
                    # Start a new current_df
                    current_df.append(row)
                else:
                    # If the row does not contain the sequence of years, append it to the current_df
                    current_df.append(row)

            
        # Append the last current_df to the dfs list
        if current_df:
            dfs.append(pd.DataFrame(current_df))

        st.write(dfs)

        fte_dfs = []
        for name, df, _ in dfs:
            possible_fte_cols = [col for col in df.columns if "FTE" in str(col)]
            possible_role_cols = [col for col in df.columns if "Functie" in str(col) or "Function" in str(col)]

            if possible_fte_cols and possible_role_cols:
                role_col = possible_role_cols[0]
                fte_col = possible_fte_cols[0]

                fte_summary = df.groupby(role_col)[fte_col].sum().reset_index()
                fte_summary["File"] = name
                fte_dfs.append(fte_summary)

                st.markdown(f"**FTEs in {name}:**")
                st.dataframe(fte_summary, use_container_width=True)

        if fte_dfs:
            all_fte = pd.concat(fte_dfs)
            total_fte = all_fte.groupby(all_fte.columns[0])[all_fte.columns[1]].sum().reset_index()
            total_fte.rename(columns={total_fte.columns[0]: "Function Type", total_fte.columns[1]: "Total FTE"}, inplace=True)

            st.markdown("**Aggregated FTEs across all files:**")
            st.dataframe(total_fte, use_container_width=True)

            fig, ax = plt.subplots(figsize=(8, 5))
            for name, group in all_fte.groupby("File"):
                ax.bar(group.iloc[:, 0], group.iloc[:, 1], alpha=0.5, label=name)

            ax.bar(total_fte["Function Type"], total_fte["Total FTE"], color="black", alpha=0.7, label="Total")
            ax.set_ylabel("FTEs")
            ax.set_title("FTEs per Function Type")
            ax.legend()
            st.pyplot(fig)

        # Kosten split by categories
        st.subheader("💰 Kosten Breakdown by Category (Column B)")
        keywords = [
            "Operationele Personeelskosten",
            "Ontwikkeling Personeelskosten",
            "Materiële kosten",
            "Overige kosten"
        ]
        for name, df, _ in dfs:
            st.markdown(f"**Split categories for {name}:**")
            split_dfs = split_by_keywords(df, keywords, col_idx=1)
            for keyword, sub_df in split_dfs.items():
                st.markdown(f"🔹 {keyword}")
                st.dataframe(sub_df, use_container_width=True)

    # Special handling for Business Case
    elif sheet_name == "Business Case" and dfs:
        st.subheader("📈 Business Case Comparison by Year")

        all_cases = []
        years_ref = None

        for name, _, excel_data in dfs:
            try:
                case_df, years = read_business_case(excel_data)
                if case_df is not None:
                    st.markdown(f"**{name} – Extracted Years:** {years}")
                    st.dataframe(case_df, use_container_width=True)

                    totals = case_df.sum()
                    totals.name = name
                    all_cases.append(totals)
                    years_ref = years
            except Exception as e:
                st.error(f"Error extracting Business Case from {name}: {e}")

        if all_cases:
            result = pd.DataFrame(all_cases).T
            st.markdown("**Comparison across files (summed per year):**")
            st.dataframe(result, use_container_width=True)

            # Plot
            fig, ax = plt.subplots(figsize=(10, 5))
            for col in result.columns:
                ax.plot(result.index, result[col], marker="o", label=col)
            ax.set_xlabel("Year")
            ax.set_ylabel("Values")
            ax.set_title("Business Case Comparison")
            ax.legend()
            st.pyplot(fig)

    # General handling for other sheets
    elif dfs:
        st.subheader("➕ Aggregated Numeric Summary")

        try:
            numeric_dfs = [df.select_dtypes(include="number") for _, df, _ in dfs]
            df_sum = sum(numeric_dfs)

            st.write("**Summed numeric values across all files:**")
            st.dataframe(df_sum, use_container_width=True)

            fig, ax = plt.subplots(figsize=(8, 5))
            for name, df, _ in dfs:
                df_numeric = df.select_dtypes(include="number")
                if not df_numeric.empty:
                    df_numeric.sum().plot(kind="bar", ax=ax, alpha=0.5, label=name)

            df_sum.sum().plot(kind="bar", ax=ax, color="black", label="Total", linewidth=2)
            ax.set_ylabel("Values")
            ax.set_title(f"Comparison of '{sheet_name}'")
            ax.legend()
            st.pyplot(fig)

        except Exception as e:
            st.error(f"Could not aggregate data: {e}")
