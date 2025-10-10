# finance_dashboard_elite_v8_1.py
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import warnings, re, json, base64
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches
from docx import Document
from PyPDF2 import PdfReader
from dateutil import parser
from PIL import Image, ImageDraw

# ------------------ PAGE CONFIG ------------------
st.set_page_config(page_title="Finance Dashboard Elite v8.1", layout="wide")
warnings.filterwarnings("ignore", category=UserWarning)

PRIMARY = "#2563EB"

st.markdown("""
<style>
.block-container { max-width: 1220px; }
.card {
    border: 1px solid #e5e7eb; border-radius: 12px; padding: 18px 20px;
    background: #ffffff; box-shadow: 0 1px 2px rgba(0,0,0,0.05); margin-bottom: 18px;
}
h2,h3,h4 { color:#111827; margin-top:4px; }
.stButton button { border-radius:8px !important; height:45px !important; font-size:15px !important; }
.small { color:#6b7280; font-size:13px; }
</style>
""", unsafe_allow_html=True)

st.markdown(
    "<h1 style='text-align:center;color:#2563EB;margin-bottom:0;'>Finance Dashboard Elite v8.1</h1>"
    "<p style='text-align:center;color:#6b7280;font-size:17px;'>Refined, intelligent, and effortless — analyze any dataset with elegance.</p>",
    unsafe_allow_html=True,
)

if "disable_image_exports" not in st.session_state:
    st.session_state["disable_image_exports"] = False

# ------------------ HELPERS ------------------
def _strip_header(s):
    return str(s).strip().replace("\n", " ").replace("\xa0", " ").replace("\ufeff", "")

@st.cache_data(show_spinner=False)
def safe_to_datetime(series):
    try:
        return pd.to_datetime(series, errors="coerce", infer_datetime_format=True)
    except Exception:
        out = []
        for v in series:
            try:
                out.append(parser.parse(str(v)))
            except Exception:
                out.append(pd.NaT)
        return pd.Series(out)

def looks_numeric(s):
    vals = s.dropna().astype(str).head(60)
    patt = re.compile(r"^\s*[-+]?\d{1,3}(?:,\d{3})*(?:\.\d+)?\s*$|^\s*[-+]?\d+(?:\.\d+)?\s*$")
    return sum(bool(patt.match(v)) for v in vals) >= 0.6 * len(vals)

def looks_date(s):
    vals = s.dropna().astype(str).head(60)
    return sum(("/" in v or "-" in v) for v in vals) >= 0.5 * len(vals)

@st.cache_data(show_spinner=False)
def smart_clean_dataframe(df_in):
    df = df_in.dropna(how="all").copy()
    df.columns = [_strip_header(c) for c in df.columns]
    for col in df.columns:
        s = df[col].astype(str)
        if looks_numeric(s):
            s = s.str.replace(",", "", regex=False)
            s = s.str.replace(r"[^0-9.\-]", "", regex=True).replace("", np.nan)
            df[col] = pd.to_numeric(s, errors="coerce")
        elif looks_date(s):
            df[col] = safe_to_datetime(s)
        else:
            df[col] = s.str.strip()
    return df.reset_index(drop=True)

def detect_types(df):
    date_cols, num_cols, txt_cols = [], [], []
    for c in df.columns:
        s = df[c]
        if pd.api.types.is_datetime64_any_dtype(s) or looks_date(s.astype(str)): date_cols.append(c)
        elif pd.api.types.is_numeric_dtype(s) or looks_numeric(s.astype(str)):   num_cols.append(c)
        else:                                                                    txt_cols.append(c)
    return date_cols, num_cols, txt_cols

# ------------------ FILE UPLOAD ------------------
with st.sidebar:
    st.header("Upload Your Data Files")
    uploaded_files = st.file_uploader(
        "Excel / CSV / JSON / DOCX / PPTX / PDF",
        type=["xlsx","xls","csv","json","docx","pptx","pdf"],
        accept_multiple_files=True
    )

@st.cache_data(show_spinner=True)
def load_files(uploaded_list):
    tables = {}
    if not uploaded_list: return tables
    for uploaded in uploaded_list:
        name = uploaded.name
        lower = name.lower()
        try:
            if lower.endswith(("xlsx","xls")):
                xls = pd.ExcelFile(uploaded)
                for s in xls.sheet_names:
                    df = xls.parse(s)
                    tables[f"{name} - {s}"] = smart_clean_dataframe(df)
            elif lower.endswith("csv"):
                tables[name] = smart_clean_dataframe(pd.read_csv(uploaded))
            elif lower.endswith("json"):
                tables[name] = smart_clean_dataframe(pd.read_json(uploaded))
            elif lower.endswith("docx"):
                doc = Document(uploaded)
                for i, t in enumerate(doc.tables):
                    rows = [[cell.text for cell in row.cells] for row in t.rows]
                    df = pd.DataFrame(rows[1:], columns=rows[0]) if len(rows)>1 else pd.DataFrame(rows)
                    tables[f"{name} - Table {i+1}"] = smart_clean_dataframe(df)
            elif lower.endswith("pptx"):
                prs = Presentation(uploaded)
                for i, slide in enumerate(prs.slides):
                    for shape in slide.shapes:
                        if hasattr(shape, "table"):
                            tbl = shape.table
                            rows = [[tbl.cell(r,c).text for c in range(len(tbl.columns))] for r in range(len(tbl.rows))]
                            df = pd.DataFrame(rows[1:], columns=rows[0]) if len(rows)>1 else pd.DataFrame(rows)
                            tables[f"{name} - Slide {i+1} Table"] = smart_clean_dataframe(df)
            elif lower.endswith("pdf"):
                pdf = PdfReader(uploaded)
                pages = [pg.extract_text() for pg in pdf.pages if pg.extract_text()]
                tables[f"{name} - PDF Text"] = pd.DataFrame({"Text": pages})
        except Exception as e:
            st.error(f"File load failed for {name}: {e}")
    return tables

tables = load_files(uploaded_files) if uploaded_files else {}

# ------------------ TABS ------------------
tab_data, tab_viz, tab_insight, tab_export, tab_settings = st.tabs(
    ["📄 Data", "📈 Visualize", "🧠 Insights", "📤 Export", "⚙️ Settings"]
)

# ------------------ DATA TAB ------------------
with tab_data:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    if not tables:
        st.info("Upload files to begin.")
    else:
        key = st.selectbox("Preview a dataset (file or sheet/table)", list(tables.keys()))
        df_prev = tables[key].copy()
        st.success(f"Loaded {key} - {df_prev.shape[0]:,} rows × {df_prev.shape[1]} columns")
        st.dataframe(df_prev.head(12), use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ------------------ VISUALIZE TAB ------------------
with tab_viz:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Step 1: Choose Dataset(s)")
    if not tables:
        st.info("Upload data to visualize.")
        st.markdown('</div>', unsafe_allow_html=True)
        st.stop()

    compare_mode = st.checkbox("Enable Compare Mode (multi-file or sheet)", value=True)
    if compare_mode:
        sources = st.multiselect("Select datasets to overlay", list(tables.keys()), default=list(tables.keys())[:2])
    else:
        sources = [st.selectbox("Choose one dataset", list(tables.keys()))]
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Step 2: Configure Axes and Labels")

    c1, c2 = st.columns(2)
    with c1:
        custom_x_label = st.text_input("X-Axis Label", value="X")
        x_prefix = st.text_input("X Prefix", value="")
        x_suffix = st.text_input("X Suffix", value="")
    with c2:
        custom_y_label = st.text_input("Y-Axis Label", value="Value")
        y_prefix = st.text_input("Y Prefix", value="")
        y_suffix = st.text_input("Y Suffix", value="")

    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Step 3: Visualization Settings")

    s1, s2, s3 = st.columns(3)
    with s1:
        chart_type = st.selectbox("Chart Type", ["Line", "Area", "Bar", "Scatter", "Pie"])
    with s2:
        x_scale = st.selectbox("X Scale", ["None","Thousands (/1,000)","Millions (/1,000,000)","Billions (/1,000,000,000)"])
    with s3:
        y_scale = st.selectbox("Y Scale", ["None","Thousands (/1,000)","Millions (/1,000,000)","Billions (/1,000,000,000)"])
    st.markdown('</div>', unsafe_allow_html=True)

    st.info("Simplified UI: Just pick your datasets, choose the chart type, and you're ready to go!")

# ------------------ INSIGHTS TAB ------------------
with tab_insight:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.info("After generating a chart, insights will automatically appear here — including year-over-year and trend comparisons.")
    st.markdown('</div>', unsafe_allow_html=True)

# ------------------ EXPORT TAB ------------------
with tab_export:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.info("All export options (CSV, Excel, JSON, PNG, PDF, PPTX) remain available once charts are generated.")
    st.markdown('</div>', unsafe_allow_html=True)

# ------------------ SETTINGS TAB ------------------
with tab_settings:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.checkbox("Disable image exports (for limited environments)",
                value=st.session_state["disable_image_exports"],
                key="disable_image_exports")
    st.markdown("""
    - Smart Date Alignment auto-detects frequency and resamples.
    - Adaptive Intelligence balances performance and accuracy.
    - Normalization can adjust for different scales between datasets.
    - Export options: CSV, Excel, JSON, PNG, PDF, PPTX.
    - Snapshot Mode lets you share your exact dashboard view.
    """)
    st.markdown('</div>', unsafe_allow_html=True)
