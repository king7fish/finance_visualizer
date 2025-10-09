# finance_visualizer_advanced.py
"""
Finance Visualizer Advanced
Supports: CSV, XLS/XLSX, JSON, DOCX (tables), PPTX (tables)
Features:
 - multi-sheet Excel support
 - multi-table DOCX/PPTX support
 - auto detect date/numeric columns
 - choose X and Y axes (multi Y, dual axis)
 - date filtering and resampling (D/W/M/Q/Y) with aggregation
 - currency conversion via exchangerate.host (live) or manual rate
 - unit scaling & tick increments
 - export PNG, PDF, PPTX slide, CSV, JSON
 - caching for better performance on large files
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from io import BytesIO
import json
import requests
from zipfile import ZipFile
from pptx import Presentation
from pptx.util import Inches
from docx import Document
from datetime import datetime

st.set_page_config(page_title="Finance Visualizer Advanced", layout="wide")

# ---------------------- Helpers ----------------------
@st.cache_data
def fetch_exchange_rate(from_curr: str, to_curr: str):
    """Fetch conversion rate (1 from_curr -> ? to_curr) using exchangerate.host"""
    try:
        url = f"https://api.exchangerate.host/convert?from={from_curr}&to={to_curr}&amount=1"
        r = requests.get(url, timeout=8)
        data = r.json()
        if data.get("info") and "rate" in data["info"]:
            return float(data["info"]["rate"])
        return None
    except Exception:
        return None

@st.cache_data
def read_excel_all_sheets(file_bytes):
    xls = pd.ExcelFile(file_bytes)
    sheets = {}
    for s in xls.sheet_names:
        try:
            df = xls.parse(s, header=None)  # read raw, we'll try to clean
            sheets[s] = df
        except Exception:
            sheets[s] = None
    return sheets

@st.cache_data
def read_csv_bytes(file_bytes):
    file_bytes.seek(0)
    return pd.read_csv(file_bytes)

@st.cache_data
def read_json_bytes(file_bytes):
    file_bytes.seek(0)
    return pd.read_json(file_bytes)

def extract_tables_from_docx(file_bytes):
    file_bytes.seek(0)
    doc = Document(file_bytes)
    tables = []
    for t in doc.tables:
        rows = []
        for r in t.rows:
            rows.append([c.text for c in r.cells])
        # build df
        if len(rows) >= 2:
            df = pd.DataFrame(rows[1:], columns=rows[0])
        else:
            df = pd.DataFrame(rows)
        tables.append(df)
    return tables

def extract_tables_from_pptx(file_bytes):
    file_bytes.seek(0)
    prs = Presentation(file_bytes)
    tables = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "table"):
                tbl = shape.table
                rows = []
                for r in range(len(tbl.rows)):
                    row_cells = []
                    for c in range(len(tbl.columns)):
                        row_cells.append(tbl.cell(r, c).text)
                    rows.append(row_cells)
                if len(rows) >= 2:
                    df = pd.DataFrame(rows[1:], columns=rows[0])
                else:
                    df = pd.DataFrame(rows)
                tables.append(df)
    return tables

def try_make_datetime(series):
    s = pd.to_datetime(series, errors="coerce")
    return s

def detect_date_columns(df):
    candidates = []
    for c in df.columns:
        try:
            parsed = pd.to_datetime(df[c].astype(str), errors="coerce")
            if parsed.notna().sum() / max(1, len(parsed)) > 0.5:
                candidates.append(c)
        except Exception:
            continue
    return candidates

def detect_numeric_columns(df):
    nums = []
    for c in df.columns:
        try:
            coerced = pd.to_numeric(df[c].astype(str).str.replace(",",""), errors="coerce")
            if coerced.notna().sum() / max(1, len(coerced)) > 0.5:
                nums.append(c)
        except Exception:
            continue
    return nums

def smart_clean_header(df):
    # If first row contains headers (strings) and many columns named Unnamed -> try to promote a header
    # If first non-null row has many strings, use it as header
    sample = df.iloc[:5].astype(str).fillna("")
    # heuristic: if any column has a descriptive value in row 1, promote it
    first_non_empty_row_idx = None
    for i in range(min(5, len(df))):
        row = df.iloc[i].astype(str).str.strip()
        nonempty = (row != "") & (row.str.lower() != "nan")
        if nonempty.sum() >= 1:
            first_non_empty_row_idx = i
            break
    if first_non_empty_row_idx is not None:
        header_candidate = df.iloc[first_non_empty_row_idx].astype(str).tolist()
        # apply if candidate has unique-ish entries
        if len(set(header_candidate)) >= max(1, 0.6*len(header_candidate)):
            new_df = df.iloc[first_non_empty_row_idx+1:].copy()
            new_df.columns = [str(h).strip() if str(h).strip() != "" else f"col_{i}" for i,h in enumerate(header_candidate)]
            new_df = new_df.reset_index(drop=True)
            return new_df
    # fallback: if columns are Unnamed..., keep as-is but give safe names
    new_cols = []
    for i,c in enumerate(df.columns):
        name = str(c).strip()
        if name == "" or name.lower().startswith("unnamed"):
            new_cols.append(f"col_{i}")
        else:
            new_cols.append(name)
    df.columns = new_cols
    return df

@st.cache_data
def dataframe_from_sheet(df_raw):
    # Accepts raw df (no header) and tries to clean
    cleaned = smart_clean_header(df_raw)
    # strip columns
    cleaned.columns = [str(c).strip() for c in cleaned.columns]
    # drop empty columns
    cleaned = cleaned.loc[:, (cleaned.notna().sum() > 0)]
    return cleaned

def resample_time_series(df, date_col, freq, agg="sum"):
    tmp = df.copy()
    tmp[date_col] = pd.to_datetime(tmp[date_col], errors="coerce")
    tmp = tmp.dropna(subset=[date_col])
    tmp = tmp.set_index(date_col)
    if agg == "sum":
        res = tmp.resample(freq).sum().reset_index()
    else:
        res = tmp.resample(freq).mean().reset_index()
    return res

def build_plot(df, x_col, y_cols, chart_type, colors_map, line_width=2, marker_size=6, log_y=False, y_tick=0, dual_axis=False):
    fig = go.Figure()
    for i, y in enumerate(y_cols):
        if y not in df.columns:
            continue
        yvals = pd.to_numeric(df[y].astype(str).str.replace(",",""), errors="coerce")
        if chart_type in ("Line","Area"):
            mode = "lines+markers"
            fill = "tozeroy" if chart_type == "Area" else None
            fig.add_trace(go.Scatter(x=df[x_col], y=yvals, mode=mode, name=y,
                                     line=dict(color=colors_map.get(y,"#1f77b4"), width=line_width),
                                     marker=dict(size=marker_size), fill=fill,
                                     yaxis="y" if not dual_axis else ("y" if i % 2 == 0 else "y2")))
        elif chart_type in ("Bar","Stacked Bar"):
            fig.add_trace(go.Bar(x=df[x_col], y=yvals, name=y, marker_color=colors_map.get(y,"#1f77b4"),
                                 yaxis="y" if not dual_axis else ("y" if i % 2 == 0 else "y2")))
        elif chart_type == "Pie":
            # Only first y used
            values = yvals
            labels = df[x_col].astype(str)
            fig = go.Figure(go.Pie(labels=labels, values=values, marker=dict(colors=[colors_map.get(y_cols[0],"#1f77b4")])))
            break

    layout = dict(title=f"{', '.join(y_cols)} vs {x_col}", template="plotly_white")
    if dual_axis and len(y_cols) > 1:
        layout.update(yaxis=dict(title=y_cols[0]), yaxis2=dict(title=y_cols[1] if len(y_cols)>1 else "", overlaying="y", side="right"))
    if y_tick and y_tick>0:
        fig.update_yaxes(dtick=y_tick)
    if log_y:
        fig.update_yaxes(type="log")
    fig.update_layout(**layout)
    fig.update_traces(hovertemplate="%{y:.4f}<extra></extra>")
    return fig

def fig_to_png_bytes(fig):
    return fig.to_image(format="png", engine="kaleido")

def create_pptx_with_image(img_bytes, title="Chart"):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    # add image to slide
    img_stream = BytesIO(img_bytes)
    pic = slide.shapes.add_picture(img_stream, Inches(1), Inches(1.2), width=Inches(8))
    # add title as textbox
    left = Inches(0.5); top = Inches(0.1); width = Inches(9); height = Inches(0.7)
    textbox = slide.shapes.add_textbox(left, top, width, height)
    textbox.text_frame.text = title
    out = BytesIO()
    prs.save(out)
    out.seek(0)
    return out.read()

# ---------------------- UI ----------------------
st.title("Finance Visualizer — Advanced file & unit support")
st.write("Upload Excel/CSV/JSON/DOCX/PPTX. Pick sheet/table, choose X and Y, resample, convert units/currency, and export charts/PowerPoint.")

uploaded = st.file_uploader("Upload a file", type=["csv","xlsx","xls","json","docx","pptx"])

if not uploaded:
    st.info("Upload a file with data (CSV, Excel, JSON, Word (.docx) tables, or PowerPoint .pptx tables).")
    st.stop()

# --- read file by type ---
filetype = uploaded.name.lower().split(".")[-1]
tables = {}   # name -> dataframe
if filetype in ("xls","xlsx"):
    sheets_raw = read_excel_all_sheets(uploaded)
    for sname, raw in sheets_raw.items():
        if raw is None:
            continue
        try:
            df_clean = dataframe_from_sheet(raw)
            tables[f"Sheet: {sname}"] = df_clean
        except Exception:
            pass
elif filetype == "csv":
    try:
        uploaded.seek(0)
        df = pd.read_csv(uploaded)
        tables["CSV"] = df
    except Exception as e:
        st.error(f"Failed to read CSV: {e}")
        st.stop()
elif filetype == "json":
    try:
        uploaded.seek(0)
        df = pd.read_json(uploaded)
        tables["JSON"] = df
    except Exception as e:
        st.error(f"Failed to read JSON: {e}")
        st.stop()
elif filetype == "docx":
    try:
        tables_list = extract_tables_from_docx(uploaded)
        for i, t in enumerate(tables_list):
            tables[f"DOCX Table {i+1}"] = t
    except Exception as e:
        st.error(f"Failed to read DOCX: {e}")
        st.stop()
elif filetype == "pptx":
    try:
        tables_list = extract_tables_from_pptx(uploaded)
        for i, t in enumerate(tables_list):
            tables[f"PPTX Table {i+1}"] = t
    except Exception as e:
        st.error(f"Failed to read PPTX: {e}")
        st.stop()
else:
    st.error("Unsupported file type.")
    st.stop()

if not tables:
    st.error("No tables found in file.")
    st.stop()

# let user pick which table/sheet to analyze
table_name = st.selectbox("Choose sheet/table", list(tables.keys()))
df = tables[table_name].copy()
st.subheader("Preview (first 10 rows)")
st.dataframe(df.head(10))

# auto-detect columns
date_candidates = detect_date_columns(df)
numeric_candidates = detect_numeric_columns(df)
all_cols = list(df.columns)

col1, col2 = st.columns([2,1])
with col1:
    x_col = st.selectbox("X axis (column)", options=all_cols, index=0)
    y_cols = st.multiselect("Y axis (numeric columns) — pick one or more", options=numeric_candidates if numeric_candidates else [c for c in all_cols if c!=x_col], default=(numeric_candidates[:1] if numeric_candidates else [all_cols[1]]))
with col2:
    st.write("Detected:")
    st.write(f"Date-like: {date_candidates}")
    st.write(f"Numeric-like: {numeric_candidates}")

# if x_col is date_like, parse
is_x_date = x_col in date_candidates
if is_x_date:
    df[x_col] = pd.to_datetime(df[x_col], errors="coerce")

# date filter
if is_x_date:
    min_dt = df[x_col].min()
    max_dt = df[x_col].max()
    start_dt, end_dt = st.date_input("Filter date range", [min_dt.date(), max_dt.date()]) if not df[x_col].isna().all() else (None, None)
    if start_dt and end_dt:
        df = df[(df[x_col] >= pd.to_datetime(start_dt)) & (df[x_col] <= pd.to_datetime(end_dt))]

# resample options if x is date
agg_map = {"None": None, "Daily":"D","Weekly":"W","Monthly":"M","Quarterly":"Q","Yearly":"A"}
resample_choice = st.selectbox("Resample / aggregate (if X is date)", options=list(agg_map.keys()), index=2)
agg_method = st.selectbox("Aggregation method", options=["sum","mean"], index=1)

if agg_map[resample_choice] and is_x_date:
    df = resample_time_series(df, x_col, agg_map[resample_choice][0] if False else agg_map[resample_choice], agg=agg_method)

# currency conversion
st.markdown("### Currency & Unit conversions")
convert_currency = st.checkbox("Convert currency (fetch rate live)", value=False)
from_curr = None; to_curr = None; rate = None
if convert_currency:
    col_curr = st.text_input("Column containing currency codes (optional)", value="")
    from_curr = st.text_input("From currency (3-letter code, e.g., USD)", value="USD")
    to_curr = st.text_input("To currency (3-letter code, e.g., EUR)", value="EUR")
    rate = fetch_exchange_rate(from_curr.upper(), to_curr.upper())
    if rate:
        st.success(f"Live rate: 1 {from_curr.upper()} = {rate:.6f} {to_curr.upper()}")
    else:
        st.warning("Could not fetch live rate — enter manual conversion rate below.")
manual_rate = st.number_input("Manual conversion rate (multiply numeric values by this) — leave 0 to ignore", value=0.0, format="%.6f")
# apply currency/unit conversion to chosen y columns if requested
if convert_currency or manual_rate > 0:
    conv_rate = manual_rate if manual_rate>0 else (rate if rate else 1.0)
    if conv_rate == 0:
        conv_rate = 1.0
    for y in y_cols:
        try:
            df[y] = pd.to_numeric(df[y].astype(str).str.replace(",",""), errors="coerce") * conv_rate
        except Exception:
            pass

# unit scaling
scale_choice = st.selectbox("Scale display", ["None","Thousands (K)","Millions (M)","Billions (B)"], index=0)
scale_factor = 1.0
if scale_choice == "Thousands (K)":
    scale_factor = 1e3
elif scale_choice == "Millions (M)":
    scale_factor = 1e6
elif scale_choice == "Billions (B)":
    scale_factor = 1e9

st.write("Chart options")
chart_type = st.selectbox("Chart type", ["Line","Area","Bar","Stacked Bar","Pie","Multi-Axis (left/right)"])
colors = {y: st.color_picker(f"Color for {y}", "#1f77b4") for y in y_cols}
line_width = st.slider("Line width", 1, 6, 2)
marker_size = st.slider("Marker size", 3, 12, 6)
y_tick = st.number_input("Y-axis tick step (0 = auto)", min_value=0, value=0)
log_y = st.checkbox("Log scale Y axis", value=False)
dual_axis = (chart_type == "Multi-Axis (left/right)")

# ensure numeric Y prepared
for y in y_cols:
    df[y] = pd.to_numeric(df[y].astype(str).str.replace(",",""), errors="coerce") / scale_factor

# build plot
fig = build_plot(df, x_col, y_cols, chart_type, colors, line_width=line_width, marker_size=marker_size, log_y=log_y, y_tick=y_tick, dual_axis=dual_axis)

st.subheader("Chart")
st.plotly_chart(fig, use_container_width=True)

# Export buttons
st.subheader("Export")
csv_bytes = df.to_csv(index=False).encode("utf-8")
col_png, col_pdf, col_pptx, col_json = st.columns(4)
with col_png:
    if st.button("Export PNG"):
        try:
            img = fig_to_png_bytes(fig)
            st.download_button("Download PNG", data=img, file_name="chart.png", mime="image/png")
        except Exception as e:
            st.error(f"PNG export failed: {e}")
with col_pdf:
    if st.button("Export PDF"):
        try:
            pdf = fig.to_image(format="pdf", engine="kaleido")
            st.download_button("Download PDF", data=pdf, file_name="chart.pdf", mime="application/pdf")
        except Exception as e:
            st.error(f"PDF export failed: {e}")
with col_pptx:
    if st.button("Export PPTX Slide"):
        try:
            img = fig_to_png_bytes(fig)
            pptx_bytes = create_pptx_with_image(img, title=f"{', '.join(y_cols)} vs {x_col}")
            st.download_button("Download PPTX", data=pptx_bytes, file_name="chart_slide.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
        except Exception as e:
            st.error(f"PPTX export failed: {e}")
with col_json:
    if st.button("Export JSON (processed data)"):
        try:
            st.download_button("Download JSON", data=df.to_json(orient="records").encode("utf-8"), file_name="processed_data.json", mime="application/json")
        except Exception as e:
            st.error(f"JSON export failed: {e}")

st.success("Done — you can now share this app and it will handle many file types and complex sheets.")
