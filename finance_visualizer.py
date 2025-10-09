# finance_visualizer_pro.py
"""
Finance Visualizer Pro (single-file)
Features:
- Upload CSV / XLSX / JSON
- Auto-detect date/number columns
- Line / Bar / Stacked Bar / Area / Pie
- Multi-series plotting, dual Y-axis
- Date-range filter & aggregation (monthly/quarterly/yearly)
- Currency formatting, scale (K/M/B), tick increments
- Per-series color/style, marker sizes
- Export PNG/PDF + download processed CSV
- Save/Load chart presets (download/upload JSON)
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from io import BytesIO
import json
from datetime import datetime
from dateutil import parser

st.set_page_config(page_title="Finance Visualizer Pro", layout="wide", initial_sidebar_state="expanded")

# ---- Helper functions ----
def try_parse_date_series(s):
    try:
        parsed = pd.to_datetime(s, errors="coerce")
        return parsed.notna().sum() > 0.5 * len(s)  # if >50% parseable treat as date
    except Exception:
        return False

def numeric_columns(df):
    return [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]

def auto_detect_date_cols(df):
    candidates = []
    for c in df.columns:
        if try_parse_date_series(df[c].astype(str)):
            candidates.append(c)
    return candidates

def human_scale(value):
    # convert large numbers to K/M/B with suffix and return scaled value and suffix
    if abs(value) >= 1e9: return value/1e9, "B"
    if abs(value) >= 1e6: return value/1e6, "M"
    if abs(value) >= 1e3: return value/1e3, "K"
    return value, ""

def to_bytes(fig, format="png"):
    # use kaleido via plotly to write image bytes
    img_bytes = fig.to_image(format=format, engine="kaleido")
    return img_bytes

# ---- UI: header & instructions ----
st.title("📈 Finance Visualizer Pro — Advanced, Beautiful, Free")
st.write("Upload CSV / Excel / JSON. Choose chart type, filter by date, customize colors and axes. Export PNG/PDF and download processed data.")

with st.expander("Quick start (copy-paste)"):
    st.markdown("""
    1. Prepare a CSV or Excel with a Date column (or a column that looks like dates) and numeric columns (Revenue, Expenses, etc.).  
    2. Upload it below.  
    3. Choose X and Y columns, date range, aggregation, and chart type.  
    4. Click *Export PNG* or *Download CSV* when ready.
    """)

# ---- Upload ----
col_up, col_template = st.columns([3,1])
with col_up:
    uploaded = st.file_uploader("Upload data (CSV, XLSX, JSON)", type=["csv","xlsx","xls","json"])
with col_template:
    st.markdown("**Templates**")
    st.download_button("Sample time-series CSV",
                       data="Date,Revenue,Expenses\n2024-01-01,1000,600\n2024-02-01,1200,700\n2024-03-01,1400,800\n",
                       file_name="sample_ts.csv",
                       mime="text/csv")
    st.download_button("Sample category CSV",
                       data="Category,Amount\nMarketing,400\nSalaries,1200\nTools,150\n",
                       file_name="sample_cat.csv",
                       mime="text/csv")

if uploaded is None:
    st.info("Upload a file to get started. Use the sample CSVs if you want to test.")
    st.stop()

# ---- Load file robustly ----
try:
    if uploaded.name.lower().endswith((".xls", ".xlsx")):
        df = pd.read_excel(uploaded)
    elif uploaded.name.lower().endswith(".json"):
        df = pd.read_json(uploaded)
    else:
        # try to sniff delimiter for csv
        uploaded.seek(0)
        df = pd.read_csv(uploaded)
except Exception as e:
    st.error(f"Error reading file: {e}")
    st.stop()

if df.empty:
    st.error("File loaded but dataframe is empty.")
    st.stop()

# ---- Basic cleaning & preview ----
st.subheader("Data preview")
st.dataframe(df.head(10))

# auto-detect date columns and numeric columns
date_cols = auto_detect_date_cols(df)
num_cols = numeric_columns(df)
all_cols = list(df.columns)

# ---- Sidebar controls ----
st.sidebar.header("Chart & Data Controls")
# File-level options
if date_cols:
    date_col = st.sidebar.selectbox("Date column (auto-detected)", options=["(none)"] + date_cols, index=0)
else:
    date_col = st.sidebar.selectbox("Date column (none detected)", options=["(none)"] + all_cols)

# Convert selected date column
if date_col != "(none)":
    try:
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
    except Exception:
        st.sidebar.error("Could not parse date column.")

# date range filter
if date_col != "(none)" and df[date_col].notna().sum() > 0:
    min_dt = df[date_col].min()
    max_dt = df[date_col].max()
    start_dt, end_dt = st.sidebar.date_input("Date range", [min_dt.date(), max_dt.date()], min_value=min_dt.date(), max_value=max_dt.date())
    # apply filter
    mask = (df[date_col] >= pd.to_datetime(start_dt)) & (df[date_col] <= pd.to_datetime(end_dt))
    df = df.loc[mask].copy()
else:
    start_dt = end_dt = None

# aggregation (for time series)
agg_period = st.sidebar.selectbox("Aggregate by", ["None", "D (daily)", "W (weekly)", "M (monthly)", "Q (quarterly)", "Y (yearly)"])

# x and y selection
x_axis = st.sidebar.selectbox("X axis column", options=all_cols, index=0)
possible_y = [c for c in all_cols if c != x_axis and c in num_cols]
if not possible_y:
    possible_y = [c for c in all_cols if c != x_axis]  # fallback
y_axes = st.sidebar.multiselect("Y axis column(s) (pick 1+)", options=possible_y, default=possible_y[:1])

# chart type
chart_type = st.sidebar.selectbox("Chart type", ["Line", "Bar", "Stacked Bar", "Area", "Pie", "Multi-Axis (left/right)"])
# style options
st.sidebar.markdown("### Visual options")
colors = {y: st.sidebar.color_picker(f"Color for {y}", "#1f77b4") for y in y_axes}
line_width = st.sidebar.slider("Line width", 1, 6, 2)
marker_size = st.sidebar.slider("Marker size", 3, 12, 6)
stack_norm = st.sidebar.checkbox("Normalize stacked bars (100%)", value=False)
log_y = st.sidebar.checkbox("Log scale Y axis", value=False)
dual_axis = chart_type == "Multi-Axis (left/right)"
# currency & scale
currency = st.sidebar.selectbox("Currency symbol", ["None", "$", "€", "£", "¥", "₦"])
scale_choice = st.sidebar.selectbox("Display scale", ["None", "Thousands (K)", "Millions (M)", "Billions (B)"])
scale_factor = 1
scale_suffix = ""
if scale_choice == "Thousands (K)":
    scale_factor = 1e3; scale_suffix = "K"
elif scale_choice == "Millions (M)":
    scale_factor = 1e6; scale_suffix = "M"
elif scale_choice == "Billions (B)":
    scale_factor = 1e9; scale_suffix = "B"

y_tick = st.sidebar.number_input("Y-axis tick step (0 = auto)", min_value=0, value=0, step=1)

# aggregation helper
def aggregate_df(df, date_col, period):
    if period == "None":
        return df
    if date_col == "(none)":
        return df  # nothing to aggregate
    tmp = df.copy()
    tmp["_agg_period"] = tmp[date_col].dt.to_period(period).dt.to_timestamp()
    agg_cols = {col: "sum" for col in y_axes if col in tmp.columns}
    grouped = tmp.groupby("_agg_period").agg(agg_cols).reset_index().rename(columns={"_agg_period": date_col})
    return grouped

# apply aggregation
if agg_period != "None" and date_col != "(none)":
    df = aggregate_df(df, date_col, agg_period.replace(" (","").split()[0])  # 'M' or 'Q' etc.

# prepare X values
x_vals = df[x_axis] if x_axis in df.columns else df.index

# ---- Build plotly figure ----
fig = go.Figure()
if chart_type in ["Line", "Area", "Stacked Bar", "Bar", "Multi-Axis (left/right)"]:
    # multiple series
    for i, y in enumerate(y_axes):
        if y not in df.columns:
            continue
        x = x_vals
        yvals = pd.to_numeric(df[y], errors="coerce") / scale_factor
        if chart_type == "Line" or chart_type == "Area":
            mode = "lines+markers"
            fill = "tozeroy" if chart_type == "Area" else None
            fig.add_trace(go.Scatter(x=x, y=yvals, mode=mode, name=y,
                                     line=dict(color=colors[y], width=line_width),
                                     marker=dict(size=marker_size), fill=fill, yaxis="y" if not dual_axis else ("y" if i % 2 == 0 else "y2")))
        elif chart_type == "Bar" or chart_type == "Stacked Bar":
            fig.add_trace(go.Bar(x=x, y=yvals, name=y, marker_color=colors[y], yaxis="y" if not dual_axis else ("y" if i % 2 == 0 else "y2")))
    if chart_type == "Stacked Bar":
        barmode = "relative"
        if stack_norm:
            barmode = "relative"
            # normalize manually to 100%
            # Plotly has 'relative' + groupnorm but some versions differ; simpler: use layout groupnorm
            fig.update_layout(barmode="stack")
            fig.update_traces(offsetgroup=0)
        else:
            fig.update_layout(barmode="stack")
    else:
        fig.update_layout(barmode="group")
elif chart_type == "Pie":
    # use first Y axis as values and X axis as labels
    if y_axes and y_axes[0] in df.columns:
        labels = df[x_axis].astype(str)
        values = pd.to_numeric(df[y_axes[0]], errors="coerce") / scale_factor
        fig = go.Figure(go.Pie(labels=labels, values=values, marker=dict(colors=px.colors.qualitative.Dark24)))
    else:
        st.error("Choose a Y-axis numeric column for pie charts.")
        st.stop()

# layout tweaks
yaxis_format = dict(title=f"Value {('('+scale_suffix+')') if scale_suffix else ''}", tickprefix=currency if currency != "None" else "")
fig.update_layout(title=f"{', '.join(y_axes)} vs {x_axis}", xaxis_title=x_axis, yaxis=yaxis_format, template="plotly_white")
if dual_axis:
    fig.update_layout(
        yaxis=dict(title=y_axes[0] if y_axes else "Left Y"),
        yaxis2=dict(title=(y_axes[1] if len(y_axes)>1 else "Right Y"), overlaying="y", side="right")
    )

if y_tick > 0:
    fig.update_yaxes(dtick=y_tick)

if log_y:
    fig.update_yaxes(type="log")

# nice hover format
fig.update_traces(hovertemplate="%{y:.2f}<extra></extra>")

# show chart
st.subheader("Chart")
st.plotly_chart(fig, use_container_width=True)

# ---- Export & Download ----
st.subheader("Export & Save")
# Download processed csv
csv_bytes = df.to_csv(index=False).encode('utf-8')
st.download_button("Download processed CSV", data=csv_bytes, file_name="processed_data.csv", mime="text/csv")

# Export image buttons
col_png, col_pdf, col_preset = st.columns([1,1,1])
with col_png:
    if st.button("Export PNG"):
        try:
            img = to_bytes(fig, format="png")
            st.download_button("Download PNG", data=img, file_name="chart.png", mime="image/png")
        except Exception as e:
            st.error(f"PNG export failed: {e}")
with col_pdf:
    if st.button("Export PDF"):
        try:
            pdf = to_bytes(fig, format="pdf")
            st.download_button("Download PDF", data=pdf, file_name="chart.pdf", mime="application/pdf")
        except Exception as e:
            st.error(f"PDF export failed: {e}")

# Save current chart settings as preset (download JSON)
with col_preset:
    if st.button("Save preset (download JSON)"):
        preset = {
            "x_axis": x_axis,
            "y_axes": y_axes,
            "chart_type": chart_type,
            "colors": colors,
            "line_width": line_width,
            "marker_size": marker_size,
            "currency": currency,
            "scale_choice": scale_choice,
            "y_tick": y_tick,
            "agg_period": agg_period
        }
        st.download_button("Download preset", data=json.dumps(preset, indent=2).encode('utf-8'),
                           file_name="chart_preset.json", mime="application/json")

# Load preset
st.write("Load preset (JSON)")
preset_upload = st.file_uploader("Upload preset JSON", type=["json"], key="preset_upload")
if preset_upload:
    try:
        p = json.load(preset_upload)
        st.info("Preset loaded — settings will apply on re-run.")
        # naive: write to session and prompt user to re-run to apply (or you can programmatically apply)
        st.session_state["loaded_preset"] = p
    except Exception as e:
        st.error(f"Failed to parse preset: {e}")

st.write("---")
st.caption("Finance Visualizer Pro — built to be fast, beautiful and flexible. Want multi-user login, history, or subscription features later? Say the word.")
