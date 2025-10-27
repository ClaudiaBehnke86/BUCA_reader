import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="SURF BUCA Aggregator", layout="wide")

st.title("📊 SURF BUCA Aggregator")
st.markdown("""
Upload **multiple SURF Business Case Excel files** to compare and aggregate their
costs, FTEs, and revenues from the **'1. Kosten'** and **'Business Case'** sheets.
""")

uploaded_files = st.file_uploader(
    "📂 Upload one or more BUCA Excel files",
    type=["xlsx"],
    accept_multiple_files=True
)


# -----------------------------------------------------------
# Helper functions
# -----------------------------------------------------------

def flatten_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Flatten multi-index columns into readable strings."""
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [
            "_".join([str(level) for level in col if pd.notna(level)]).strip()
            for col in df.columns
        ]
    return df


def read_kosten(file, name):
    """Read the '1. Kosten' sheet and tag with BUCA name."""
    try:
        df = pd.read_excel(file, sheet_name="1. Kosten", header=[1, 3])
        df = flatten_columns(df)
        df["BUCA"] = name
        df = df.dropna(how="all")
        return df
    except Exception as e:
        st.error(f"Error reading '1. Kosten' from {name}: {e}")
        return pd.DataFrame()


def read_business_case(file, name):
    """Read the 'Business Case' sheet and tag with BUCA name."""
    try:
        df = pd.read_excel(file, sheet_name="Business Case", header=[1, 2])
        df = flatten_columns(df)
        df["BUCA"] = name
        df = df.dropna(how="all")
        return df
    except Exception as e:
        st.error(f"Error reading 'Business Case' from {name}: {e}")
        return pd.DataFrame()


def summarize_yearly_values(df, pattern="€", id_col="BUCA"):
    """Extract numeric columns per year and sum across all files."""
    numeric_cols = [c for c in df.columns if any(str(y) in c for y in ["2022","2023","2024","2025","2026"])]
    sub = df[[id_col] + numeric_cols].copy()
    sub = sub.apply(pd.to_numeric, errors="coerce").fillna(0)
    summary = sub.groupby(id_col).sum().reset_index()
    total = summary[numeric_cols].sum().to_frame(name="Total").T
    return summary, total


# -----------------------------------------------------------
# Main App Logic
# -----------------------------------------------------------

if uploaded_files:
    all_kosten = []
    all_business = []

    for f in uploaded_files:
        name = f.name.split(".xlsx")[0]
        st.markdown(f"### 📁 Processing **{name}** ...")
        kosten_df = read_kosten(f, name)
        business_df = read_business_case(f, name)

        if not kosten_df.empty:
            st.success(f"✅ Loaded '1. Kosten' for {name}")
            all_kosten.append(kosten_df)

        if not business_df.empty:
            st.success(f"✅ Loaded 'Business Case' for {name}")
            all_business.append(business_df)

    if all_kosten:
        kosten_all = pd.concat(all_kosten, ignore_index=True)
        st.subheader("📊 Aggregated '1. Kosten' (All Files Combined)")
        st.dataframe(kosten_all.head(30), use_container_width=True)

        # Yearly sum plot
        kosten_summary, kosten_total = summarize_yearly_values(kosten_all)
        st.markdown("### 💰 Total Kosten per BUCA (by Year)")
        st.dataframe(kosten_summary)

        melted = kosten_summary.melt(id_vars="BUCA", var_name="Year", value_name="Costs")
        fig = px.bar(melted, x="Year", y="Costs", color="BUCA", barmode="group", title="Kosten per Year and BUCA")
        st.plotly_chart(fig, use_container_width=True)

        total_fig = px.line(kosten_total.T, markers=True, title="Total Kosten (All BUCAs Combined)")
        st.plotly_chart(total_fig, use_container_width=True)

    if all_business:
        business_all = pd.concat(all_business, ignore_index=True)
        st.subheader("📈 Aggregated 'Business Case' (All Files Combined)")
        st.dataframe(business_all.head(30), use_container_width=True)

        business_summary, business_total = summarize_yearly_values(business_all)
        st.markdown("### 📦 Total Business Case per BUCA (by Year)")
        st.dataframe(business_summary)

        melted2 = business_summary.melt(id_vars="BUCA", var_name="Year", value_name="Value")
        fig2 = px.bar(melted2, x="Year", y="Value", color="BUCA", barmode="group", title="Business Case per Year and BUCA")
        st.plotly_chart(fig2, use_container_width=True)

        total_fig2 = px.line(business_total.T, markers=True, title="Total Business Case (All BUCAs Combined)")
        st.plotly_chart(total_fig2, use_container_width=True)

else:
    st.info("⬆️ Upload multiple Excel (.xlsx) files to start comparing BUCAs.")
