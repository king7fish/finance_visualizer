# finance_ai_pro.py
"""
Finance AI Pro (offline 'AI' heuristics + startup design)
- File types: xlsx, csv, json, docx, pptx
- Smart cleaning for messy rows
- Multi-sheet Excel support (auto-detect main sheet)
- Multi-Y plotting, dual axis, resampling
- Offline AI insights: growth, volatility, correlation, anomalies
- Exports: CSV, XLSX, PNG, PPTX
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches
from docx import Document
from datetime import datetime
from scipy import stats

# Layout & theme
st.set_page_config(page_title="Finance AI Pro", layout="wide", initial_sidebar_state="expanded")
PRIMARY = "#6C5CE7"    # vibrant violet
ACCENT = "#00BFA6"     # teal
BG = "#0f1724"         # near-black background for header
CARD = "#0b1220"
TEXT = "#E6EEF8"

st.markdown(f"""
<style>
/* page background */
[data-testid="stAppViewContainer"] {{ background-color: #071026; }}
[data-testid="stSidebar"] {{ background-color: #041427; color: {TEXT}; }}
h1, .css-18e3th9 {{ color: {TEXT}; }}
</style>
""", unsafe_allow_html=True)

# ---------- Helper Utilities ----------
@st.cache_data
def read_excel_sheets(bytes_io):
    xls = pd.ExcelFile(bytes_io)
    sheets = {}
    for s in xls.sheet_names:
        try:
            sheets[s] = pd.read_excel(xls, sheet_name=s, header=None)
        except Exception:
            sheets[s] = None
    return sheets

@st.cache_data
def read_csv_bytes(bytes_io):
    bytes_io.seek(0)
    return pd.read_csv(bytes_io)

@st.cache_data
def read_json_bytes(bytes_io):
    bytes_io.seek(0)
    return pd.read_json(bytes_io)

def extract_docx_tables(bytes_io):
    bytes_io.seek(0)
    doc = Document(bytes_io)
    tables = []
    for table in doc.tables:
        rows = []
        for r in table.rows:
            rows.append([c.text for c in r.cells])
        if len(rows) >= 2:
            df = pd.DataFrame(rows[1:], columns=rows[0])
        else:
            df = pd.DataFrame(rows)
        tables.append(df)
    return tables

def extract_pptx_tables(bytes_io):
    bytes_io.seek(0)
    prs = Presentation(bytes_io)
    tables = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "table"):
                tbl = shape.table
                rows = []
                for r in range(len(tbl.rows)):
                    rows.append([tbl.cell(r,c).text for c in range(len(tbl.columns))])
                if len(rows) >= 2:
                    df = pd.DataFrame(rows[1:], columns=rows[0])
                else:
                    df = pd.DataFrame(rows)
                tables.append(df)
    return tables

def smart_clean_header(df_raw):
    # If first row looks like header, promote it; otherwise sanitize column names
    df = df_raw.copy().reset_index(drop=True)
    # find a candidate header row within first 6 rows that has many string-like values
    header_idx = None
    for i in range(min(6, len(df))):
        row = df.iloc[i].astype(str).str.strip()
        nonempty = (row != "") & (row.str.lower() != "nan")
        if nonempty.sum() >= 0.6 * len(row):
            header_idx = i
            break
    if header_idx is not None:
        header = df.iloc[header_idx].astype(str).tolist()
        body = df.iloc[header_idx+1:].copy().reset_index(drop=True)
        # sanitize header
        header = [str(h).strip() if str(h).strip() != "" else f"col_{j}" for j,h in enumerate(header)]
        body.columns = header
        return body
    # else rename unnamed columns
    cols = []
    for j,c in enumerate(df.columns):
        name = str(c).strip()
        if name == "" or name.lower().startswith("unnamed"):
            cols.append(f"col_{j}")
        else:
            cols.append(name)
    df.columns = cols
    return df

def drop_mostly_text_rows(df, strictness=0.5):
    # convert commas and spaces in numeric-looking strings
    df_work = df.copy()
    for c in df_work.columns:
        df_work[c] = df_work[c].astype(str).str.replace(",", "").str.strip()
    # numeric check per row
    def row_numeric_fraction(row):
        nums = pd.to_numeric(row.replace("", np.nan), errors="coerce")
        return nums.notna().sum() / max(1, len(row))
    frac = df_work.apply(row_numeric_fraction, axis=1)
    keep = frac >= strictness
    removed = df.loc[~keep]
    kept = df.loc[keep].reset_index(drop=True)
    return kept, removed

def try_parse_dates(df):
    date_cols = []
    for c in df.columns:
        s = pd.to_datetime(df[c].astype(str), errors="coerce")
        if s.notna().sum() >= 0.4 * len(s):  # 40% parseable -> treat as date
            date_cols.append(c)
    return date_cols

def resample_safe(df, time_col, freq="M", agg="mean"):
    try:
        df2 = df.copy()
        df2[time_col] = pd.to_datetime(df2[time_col], errors="coerce")
        df2 = df2.dropna(subset=[time_col])
        df2 = df2.set_index(time_col)
        numeric = df2.select_dtypes(include=["number"])
        if numeric.shape[1] == 0:
            return df2.reset_index()
        res = numeric.resample(freq).agg(agg)
        res = res.reset_index()
        return res
    except Exception as e:
        st.warning(f"Resample failed: {e}")
        return df

def build_plot(df, x, ys, chart, colors=None, log_y=False):
    if chart == "Line":
        fig = px.line(df, x=x, y=ys, markers=True)
    elif chart == "Bar":
        fig = px.bar(df, x=x, y=ys)
    elif chart == "Area":
        fig = px.area(df, x=x, y=ys)
    elif chart == "Scatter":
        fig = px.scatter(df, x=x, y=ys if len(ys)==1 else ys[0], trendline="ols")
    elif chart == "Pie":
        # only first y supported for pie
        fig = px.pie(df, names=x, values=ys[0])
    else:
        fig = px.line(df, x=x, y=ys, markers=True)
    if colors:
        # set first color map only
        for i, trace in enumerate(fig.data):
            trace.marker.color = colors.get(ys[i], None) if i < len(ys) else None
    if log_y:
        fig.update_yaxes(type="log")
    fig.update_layout(template="plotly_dark", title_font=dict(size=18, color=PRIMARY), legend_title_text="Series")
    return fig

def detect_anomalies(series, z_thresh=3.5):
    # robust z-score
    arr = pd.to_numeric(series, errors="coerce")
    arr = arr.dropna()
    if len(arr) == 0:
        return []
    z = np.abs(stats.zscore(arr))
    if np.isnan(z).all():
        return []
    return list(np.where(z > z_thresh)[0])

def make_pptx_from_image_bytes(img_bytes, title="Chart"):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    left = Inches(0.5); top = Inches(0.8); width = Inches(9); height = Inches(5)
    pic = slide.shapes.add_picture(BytesIO(img_bytes), left, top, width=width, height=height)
    # title
    txBox = slide.shapes.add_textbox(Inches(0.4), Inches(0.2), Inches(9), Inches(0.6))
    tf = txBox.text_frame
    tf.text = title
    return prs

# ---------- UI ----------
st.markdown(f"<div style='background:{BG};padding:14px;border-radius:8px'><h1 style='color:{TEXT};'>Finance AI Pro — Offline Insights · Design: Modern</h1></div>", unsafe_allow_html=True)
st.write("")  # spacing

# Sidebar settings
st.sidebar.header("AI & Visualization Settings")
strictness = st.sidebar.slider("Cleaning strictness", 0.2, 0.9, 0.5)
resample_map = {"None": None, "Daily": "D", "Weekly": "W", "Monthly": "M", "Quarterly": "Q", "Yearly": "A"}

# Upload
uploaded = st.file_uploader("Upload Excel/CSV/JSON/DOCX/PPTX", type=["xlsx","csv","json","docx","pptx"])
if not uploaded:
    st.info("Upload a file to start. Try your Corporate finance file.")
    st.stop()

file_ext = uploaded.name.split(".")[-1].lower()
tables = {}

try:
    if file_ext in ("xls","xlsx"):
        sheets = read_excel_sheets(uploaded)
        # clean / convert each sheet to a table (smart header)
        for name, raw in sheets.items():
            if raw is None: 
                continue
            df_clean = smart_clean_header(raw)
            # drop empty-all columns
            df_clean = df_clean.loc[:, df_clean.notna().sum() > 0]
            tables[f"Sheet: {name}"] = df_clean

    elif file_ext == "csv":
        df0 = read_csv_bytes(uploaded)
        df0 = smart_clean_header(df0)
        tables["CSV"] = df0
    elif file_ext == "json":
        df0 = read_json_bytes(uploaded)
        df0 = smart_clean_header(df0)
        tables["JSON"] = df0
    elif file_ext
