import pandas as pd
import streamlit as st
import plotly.express as px

# === Setup ===
st.set_page_config(layout="wide", page_title="AT&C Dashboard — Highlighted Visuals")
xls = pd.ExcelFile("atcmay.xlsx")
sheet_names = xls.sheet_names
selected_sheet = st.sidebar.selectbox("📁 Select Sheet", sheet_names)
st.title(f"📊 COMPARISON AT&C PROGRESS REPORT — {selected_sheet.upper()} — MAY 2025")

# === Read Excel with 2-level header
df_raw = pd.read_excel(xls, sheet_name=selected_sheet, header=[0, 1])

# === Header cleaner
def clean_col(col):
    if isinstance(col, tuple) and len(col) == 2:
        group, sub = col
        group = str(group).replace("Unnamed", "").replace("COMPARISION AT&C PROG REPORT (OVERALL) MAY-25", "").strip()
        sub = str(sub).replace("2025-05-01", "May-25").replace("2024-05-01", "May-24").replace("00:00:00", "").strip()
        sub = sub.replace("01-05-2025", "May-25").replace("01-05-2024", "May-24").strip()
        return (group.strip(), sub.strip())
    else:
        return ("NAME OF CIR./DIV./ZONE", "")

df_raw.columns = pd.MultiIndex.from_tuples([clean_col(col) for col in df_raw.columns])
df_raw.columns = pd.MultiIndex.from_tuples([
    ("NAME OF CIR./DIV./ZONE", "") if "NAME OF" in col[0].upper() else col
    for col in df_raw.columns
])

# === Numeric conversion and rounding
df = df_raw.copy()
for col in df.columns:
    df[col] = pd.to_numeric(df[col], errors="ignore")
    if pd.api.types.is_numeric_dtype(df[col]):
        df[col] = df[col].round(2)

# === Flatten columns
flat_df = df.copy()
flat_df.columns = [
    "NAME OF CIR./DIV./ZONE" if col[0] == "NAME OF CIR./DIV./ZONE"
    else f"{col[0]} {col[1]}"
    for col in df.columns
]
flat_df.columns = [col.replace("  ", " ").strip() for col in flat_df.columns]

# === Identify key columns
zone_col = next((c for c in flat_df.columns if "NAME OF CIR./DIV./ZONE" in c.upper()), None)

# === Find metrics with May-25 and May-24
comparison_metrics = sorted(set([
    c.split(" ")[0].strip() for c in flat_df.columns
    if ("May-25" in c or "May-24" in c) and "% INC/DEC" not in c and c != zone_col
]))

# === Custom metric color map
metric_colors = {
    "ATC LOSS": ("#d62728", "#1f77b4"),
    "INPUT ENERGY PROG(MU)": ("#9467bd", "#8c564b"),
    "UNIT SOLD PROG(MU)": ("#2ca02c", "#ff7f0e"),
    "REALISATION PROG(LAKH)": ("#17becf", "#bcbd22"),
    "COLLECTION EFFICIENCY": ("#e377c2", "#7f7f7f"),
    "ABR": ("#8c564b", "#e377c2"),
}

st.subheader("📊 Metric Comparison Charts — May-25 vs May-24")

for metric in comparison_metrics:
    col_25 = next((c for c in flat_df.columns if metric in c and "May-25" in c), None)
    col_24 = next((c for c in flat_df.columns if metric in c and "May-24" in c), None)

    if col_25 and col_24 and zone_col:
        comp_df = flat_df[[zone_col, col_25, col_24]].copy()
        comp_df = comp_df.melt(id_vars=zone_col, value_vars=[col_25, col_24],
                               var_name="Month", value_name="Value")
        comp_df["Month"] = comp_df["Month"].apply(lambda x: "May-25" if "May-25" in x else "May-24")

        color_pair = metric_colors.get(metric.upper(), ("#d62728", "#1f77b4"))

        fig = px.bar(
            comp_df,
            x=zone_col,
            y="Value",
            color="Month",
            text="Value",
            barmode="group",
            title=f"{metric} — May-25 vs May-24",
            labels={zone_col: "Division", "Value": metric},
            height=500,
            color_discrete_map={"May-25": color_pair[0], "May-24": color_pair[1]}
        )

        fig.update_traces(width=0.35, textposition="outside")
        fig.update_layout(
            xaxis_title="<b>📍 NAME OF CIR./DIV./ZONE</b>",
            yaxis_title=f"<b>📊 {metric}</b>",
            xaxis_tickangle=-45,
            font=dict(size=13),
            xaxis=dict(tickfont=dict(size=13, family="Arial", color="#333", weight="bold")),
            legend_title_text="<b>Month</b>"
        )

        st.plotly_chart(fig, use_container_width=True)

# === Final Summary Table
st.subheader("📋 Structured Summary Table")
st.dataframe(flat_df)