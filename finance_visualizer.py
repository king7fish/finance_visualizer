# finance_dashboard_elite_plus.py
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import warnings, re, requests
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches
from docx import Document
from PyPDF2 import PdfReader
from dateutil import parser

# ---------------- CONFIG ----------------
st.set_page_config(page_title="Finance Dashboard Elite+", layout="wide")
PRIMARY = "#2563EB"
warnings.filterwarnings("ignore", category=UserWarning)

st.markdown(
    f"<h1 style='text-align:center;color:{PRIMARY};margin-bottom:4px'>🏆 Finance Dashboard Elite+</h1>"
    "<p style='text-align:center;color:#6b7280;font-size:17px;'>Dual-axis control. Analyst-grade precision.</p>",
    unsafe_allow_html=True,
)

if "disable_image_exports" not in st.session_state:
    st.session_state["disable_image_exports"] = False

# ---------------- HELPERS ----------------
@st.cache_data(show_spinner=False)
def safe_to_datetime(series):
    try:
        return pd.to_datetime(series, errors="coerce", infer_datetime_format=True)
    except Exception:
        parsed = []
        for val in series:
            try:
                parsed.append(parser.parse(str(val)))
            except Exception:
                parsed.append(pd.NaT)
        return pd.Series(parsed)

def looks_numeric(s):
    vals = s.dropna().astype(str).head(60)
    patt = re.compile(r"^\s*[-+]?\d{1,3}(?:,\d{3})*(?:\.\d+)?\s*$|^\s*[-+]?\d+(?:\.\d+)?\s*$")
    return sum(bool(patt.match(v)) for v in vals) >= 0.6 * len(vals)

def looks_date(s):
    vals = s.dropna().astype(str).head(60)
    return sum(("/" in v or "-" in v) for v in vals) >= 0.5 * len(vals)

def smart_clean_dataframe(df_in):
    df = df_in.dropna(how="all").copy()
    df.columns = [str(c).strip().replace("\n", " ") for c in df.columns]
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
        if pd.api.types.is_datetime64_any_dtype(s) or looks_date(s): date_cols.append(c)
        elif pd.api.types.is_numeric_dtype(s) or looks_numeric(s):   num_cols.append(c)
        else:                                                        txt_cols.append(c)
    return date_cols, num_cols, txt_cols

# ---------------- EXPORT HELPERS ----------------
def fig_to_png_safe(fig):
    if st.session_state.get("disable_image_exports", False): return None
    try: return fig.to_image(format="png", engine="kaleido")
    except Exception: return None

def pptx_with_chart_failsafe(fig, title="Chart"):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(1)).text = title
    img = fig_to_png_safe(fig)
    if img:
        slide.shapes.add_picture(BytesIO(img), Inches(0.8), Inches(1.4), width=Inches(8.8))
    else:
        slide.shapes.add_textbox(Inches(0.8), Inches(1.4), Inches(9), Inches(3)).text_frame.text = "⚠️ Image export unavailable."
    out = BytesIO(); prs.save(out); out.seek(0)
    return out.getvalue()

# ---------------- FILE UPLOAD ----------------
st.sidebar.header("📁 Upload Data File")
u = st.sidebar.file_uploader(
    "Upload Excel / CSV / JSON / DOCX / PPTX / PDF",
    type=["xlsx","xls","csv","json","docx","pptx","pdf"]
)

@st.cache_data(show_spinner=True)
def load_file(uploaded):
    tables = {}
    name = uploaded.name.lower()
    try:
        if name.endswith(("xlsx","xls")):
            xls = pd.ExcelFile(uploaded)
            for s in xls.sheet_names:
                df = xls.parse(s)
                tables[f"Excel — {s}"] = smart_clean_dataframe(df)
        elif name.endswith("csv"):
            tables["CSV"] = smart_clean_dataframe(pd.read_csv(uploaded))
        elif name.endswith("json"):
            tables["JSON"] = smart_clean_dataframe(pd.read_json(uploaded))
        elif name.endswith("docx"):
            doc = Document(uploaded)
            for i, t in enumerate(doc.tables):
                rows = [[cell.text for cell in row.cells] for row in t.rows]
                df = pd.DataFrame(rows[1:], columns=rows[0]) if len(rows)>1 else pd.DataFrame(rows)
                tables[f"Word Table {i+1}"] = smart_clean_dataframe(df)
        elif name.endswith("pptx"):
            prs = Presentation(uploaded)
            for i, slide in enumerate(prs.slides):
                for shape in slide.shapes:
                    if hasattr(shape, "table"):
                        tbl = shape.table
                        rows = [[tbl.cell(r,c).text for c in range(len(tbl.columns))] for r in range(len(tbl.rows))]
                        df = pd.DataFrame(rows[1:], columns=rows[0]) if len(rows)>1 else pd.DataFrame(rows)
                        tables[f"PPT Table {i+1}"] = smart_clean_dataframe(df)
        elif name.endswith("pdf"):
            pdf = PdfReader(uploaded)
            pages = [pg.extract_text() for pg in pdf.pages if pg.extract_text()]
            tables["PDF Text"] = pd.DataFrame({"Text": pages})
    except Exception as e:
        st.error(f"⚠️ File load failed: {e}")
    return tables

tables = load_file(u) if u else {}

# ---------------- TABS ----------------
tab_data, tab_viz, tab_insight, tab_export, tab_settings = st.tabs(
    ["📄 Data", "📈 Visualize", "🧠 Insights", "📤 Export", "⚙️ Settings"]
)

# ---------------- DATA TAB ----------------
with tab_data:
    if not tables:
        st.info("Upload a file to begin.")
    else:
        key = st.selectbox("Select dataset", list(tables.keys()))
        df = tables[key].copy()
        st.success(f"Loaded {key} → {df.shape[0]} rows × {df.shape[1]} columns")
        st.dataframe(df.head(10))

# ---------------- VISUALIZE TAB ----------------
with tab_viz:
    if not tables:
        st.info("Upload data to visualize.")
    else:
        key = st.selectbox("Choose dataset for chart", list(tables.keys()))
        df = tables[key].copy()
        date_cols, num_cols, txt_cols = detect_types(df)

        # === Axis Configuration ===
        st.subheader("🧭 Axis Configuration")
        ignore_ai = st.toggle("Ignore AI auto-detection (manual control)", value=False)

        if ignore_ai:
            x_col = st.selectbox("X-Axis Column", df.columns)
            y_cols = st.multiselect("Y-Axis Columns", df.columns)
        else:
            x_col = st.selectbox("X-Axis (detected)", df.columns)
            y_cols = st.multiselect("Y-Axis (detected numeric)", num_cols, default=num_cols[:1] if num_cols else [])

        # === Label & Unit Controls ===
        st.subheader("🧩 Axis Labels, Units, & Scaling")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("#### X-Axis Settings")
            custom_x_label = st.text_input("X Label", value=x_col)
            x_prefix = st.text_input("X Prefix", value="")
            x_suffix = st.text_input("X Suffix", value="")
            x_scale = st.selectbox("X Scale", ["None","Thousands (÷1,000)","Millions (÷1,000,000)","Billions (÷1,000,000,000)"])
        with col2:
            st.markdown("#### Y-Axis Settings")
            custom_y_label = st.text_input("Y Label", value=", ".join(y_cols) if y_cols else "")
            y_prefix = st.text_input("Y Prefix", value="")
            y_suffix = st.text_input("Y Suffix", value="")
            y_scale = st.selectbox("Y Scale", ["None","Thousands (÷1,000)","Millions (÷1,000,000)","Billions (÷1,000,000,000)"])

        scale_map = {"None":1, "Thousands (÷1,000)":1_000, "Millions (÷1,000,000)":1_000_000, "Billions (÷1,000,000,000)":1_000_000_000}

        chart_type = st.selectbox("Chart Type", ["Line","Area","Bar","Scatter","Pie"])

        # === Date Range Filter ===
        if x_col in df.columns and pd.api.types.is_datetime64_any_dtype(df[x_col]):
            df[x_col] = safe_to_datetime(df[x_col])
            df = df.dropna(subset=[x_col])
            if not df.empty:
                min_d, max_d = df[x_col].min(), df[x_col].max()
                if min_d == max_d: max_d = min_d + pd.Timedelta(days=1)
                rng = st.date_input("Date Range", [min_d.date(), max_d.date()])
                df = df[(df[x_col] >= pd.to_datetime(rng[0])) & (df[x_col] <= pd.to_datetime(rng[1]))]

        # === Chart Rendering ===
        if y_cols:
            df_plot = df.copy()

            # Apply scaling
            if pd.api.types.is_numeric_dtype(df_plot[x_col]):
                df_plot[x_col] = df_plot[x_col] / scale_map[x_scale]
            for y in y_cols:
                if pd.api.types.is_numeric_dtype(df_plot[y]):
                    df_plot[y] = df_plot[y] / scale_map[y_scale]

            # Build chart
            if chart_type == "Line": fig = px.line(df_plot, x=x_col, y=y_cols, markers=True)
            elif chart_type == "Area": fig = px.area(df_plot, x=x_col, y=y_cols)
            elif chart_type == "Bar": fig = px.bar(df_plot, x=x_col, y=y_cols)
            elif chart_type == "Scatter": fig = px.scatter(df_plot, x=x_col, y=y_cols)
            elif chart_type == "Pie" and len(y_cols)==1: fig = px.pie(df_plot, names=x_col, values=y_cols[0])

            fig.update_layout(
                template="plotly_white", height=600,
                xaxis_title=f"{x_prefix}{custom_x_label}{x_suffix}",
                yaxis_title=f"{y_prefix}{custom_y_label}{y_suffix}"
            )
            st.plotly_chart(fig, use_container_width=True)
            st.session_state["last_fig"], st.session_state["last_df"] = fig, df_plot

# ---------------- SETTINGS TAB ----------------
with tab_settings:
    st.checkbox("Disable image exports (Option 4)",
                value=st.session_state["disable_image_exports"],
                key="disable_image_exports",
                help="Turn this on if your environment lacks Kaleido/Chrome. Data exports always work.")
    st.markdown("""
**What's New**
- X-axis scaling and prefixes/suffixes added 🎯  
- Cleaner grid layout for full dual-axis symmetry  
- Cached processing for massive datasets 🚀
""")
