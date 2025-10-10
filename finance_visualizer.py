# Finance Dashboard Elite v9.2 — GitHub & Cloud Ready Edition
# Created with ❤️ for clarity, performance, and reliability.

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import warnings, re, json, base64
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches
from docx import Document
from PyPDF2 import PdfReader
from dateutil import parser
from PIL import Image, ImageDraw

# ------------------ PAGE CONFIG ------------------
st.set_page_config(page_title="Finance Dashboard Elite v9.2", layout="wide")
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
    "<h1 style='text-align:center;color:#2563EB;margin-bottom:0;'>Finance Dashboard Elite v9.2</h1>"
    "<p style='text-align:center;color:#6b7280;font-size:17px;'>Corporate-grade analytics with effortless design.</p>",
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
        type=["xlsx", "xls", "csv", "json", "docx", "pptx", "pdf"],
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
            if lower.endswith(("xlsx", "xls")):
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
                    df = pd.DataFrame(rows[1:], columns=rows[0]) if len(rows) > 1 else pd.DataFrame(rows)
                    tables[f"{name} - Table {i+1}"] = smart_clean_dataframe(df)
            elif lower.endswith("pptx"):
                prs = Presentation(uploaded)
                for i, slide in enumerate(prs.slides):
                    for shape in slide.shapes:
                        if hasattr(shape, "table"):
                            tbl = shape.table
                            rows = [[tbl.cell(r, c).text for c in range(len(tbl.columns))] for r in range(len(tbl.rows))]
                            df = pd.DataFrame(rows[1:], columns=rows[0]) if len(rows) > 1 else pd.DataFrame(rows)
                            tables[f"{name} - Slide {i+1} Table"] = smart_clean_dataframe(df)
            elif lower.endswith("pdf"):
                pdf = PdfReader(uploaded)
                pages = [pg.extract_text() for pg in pdf.pages if pg.extract_text()]
                tables[f"{name} - PDF Text"] = pd.DataFrame({"Text": pages})
        except Exception as e:
            st.error(f"Failed to load {name}: {e}")
    return tables

tables = load_files(uploaded_files) if uploaded_files else {}

# ------------------ FIXED FUNCTION ------------------
@st.cache_data(show_spinner=True)
def resample_to_freq(df, x_col, y_cols, target_freq):
    out = df.copy()
    out.columns = [str(c).strip().replace("\xa0", " ").replace("\ufeff", "") for c in out.columns]
    y_existing = [y for y in y_cols if y in out.columns]
    if not y_existing or x_col not in out.columns:
        return pd.DataFrame(columns=[x_col] + list(y_existing))
    out[x_col] = safe_to_datetime(out[x_col])
    out = out.dropna(subset=[x_col])
    if out.empty:
        return pd.DataFrame(columns=[x_col] + list(y_existing))
    out = out.set_index(x_col)
    try:
        res = out[y_existing].resample(target_freq).mean()
    except Exception:
        return pd.DataFrame(columns=[x_col] + list(y_existing))
    res = res.reset_index()
    res.columns = [x_col] + list(y_existing)
    return res

# ------------------ TABS ------------------
tab_data, tab_viz, tab_export, tab_settings = st.tabs(
    ["📄 Data", "📈 Visualize", "📤 Export", "⚙️ Settings"]
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
    st.subheader("Visualization Setup")

    if not tables:
        st.info("Upload a file to visualize data.")
        st.markdown('</div>', unsafe_allow_html=True)
        st.stop()

    src = st.selectbox("Select Dataset", list(tables.keys()))
    df = tables[src]
    date_cols, num_cols, txt_cols = detect_types(df)

    x_col = st.selectbox("X Axis (likely a date)", date_cols or df.columns)
    y_cols = st.multiselect("Y Axis (numerical columns)", num_cols or df.columns, default=num_cols[:2] if num_cols else [])

    chart_type = st.selectbox("Chart Type", ["Line", "Bar", "Area", "Pie", "Scatter"])

    st.markdown('</div>', unsafe_allow_html=True)

    # Visualization
    if st.button("Generate Chart", type="primary"):
        if not y_cols:
            st.error("Select at least one Y-axis column.")
            st.stop()

        fig = None
        try:
            if chart_type == "Line":
                fig = px.line(df, x=x_col, y=y_cols)
            elif chart_type == "Bar":
                fig = px.bar(df, x=x_col, y=y_cols)
            elif chart_type == "Area":
                fig = px.area(df, x=x_col, y=y_cols)
            elif chart_type == "Pie" and len(y_cols) == 1:
                fig = px.pie(df, names=x_col, values=y_cols[0])
            elif chart_type == "Scatter":
                fig = px.scatter(df, x=x_col, y=y_cols[0])
        except Exception as e:
            st.error(f"Chart creation failed: {e}")

        if fig:
            st.plotly_chart(fig, use_container_width=True)
            st.session_state["fig"] = fig
            st.session_state["df"] = df

# ------------------ EXPORT TAB ------------------
with tab_export:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    if "fig" not in st.session_state:
        st.info("Generate a chart first.")
    else:
        fig = st.session_state["fig"]
        df = st.session_state["df"]

        def placeholder_png(text="Export Error", color=(0, 0, 0)):
            img = Image.new("RGB", (800, 500), color=(255, 255, 255))
            d = ImageDraw.Draw(img)
            d.text((250, 250), text, fill=color)
            buf = BytesIO()
            img.save(buf, format="PNG")
            return buf.getvalue()

        def safe_export(fig, fmt="png"):
            try:
                return fig.to_image(format=fmt, engine="kaleido")
            except Exception:
                st.warning(f"{fmt.upper()} export failed — fallback image created.")
                return placeholder_png(f"{fmt.upper()} export failed")

        png_bytes = safe_export(fig, "png")
        pdf_bytes = safe_export(fig, "pdf")

        st.download_button("⬇️ Download PNG", png_bytes, "chart.png", "image/png")
        st.download_button("⬇️ Download PDF", pdf_bytes, "chart.pdf", "application/pdf")

        st.download_button("⬇️ Download CSV", df.to_csv(index=False).encode("utf-8"), "data.csv", "text/csv")
    st.markdown('</div>', unsafe_allow_html=True)

# ------------------ SETTINGS TAB ------------------
with tab_settings:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.checkbox("Disable image exports (for limited environments)",
                value=st.session_state["disable_image_exports"],
                key="disable_image_exports")
    st.markdown("**Version 9.2 Notes:** Improved chart reliability, file parsing, and export handling.")
    st.markdown('</div>', unsafe_allow_html=True)
