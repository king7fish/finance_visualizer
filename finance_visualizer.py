# finance_dashboard_pro_plus.py
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

# ---------- CONFIG ----------
st.set_page_config(page_title="Finance Dashboard Pro+", layout="wide")
PRIMARY = "#3A86FF"
warnings.filterwarnings("ignore", category=UserWarning)

st.markdown(
    f"<h1 style='text-align:center;color:{PRIMARY}'>🏆 Finance Dashboard Pro+</h1>"
    "<p style='text-align:center;color:#6b7280'>AI-assisted insights with human control 🎛️</p>",
    unsafe_allow_html=True,
)

if "disable_image_exports" not in st.session_state:
    st.session_state["disable_image_exports"] = False

# ---------- HELPERS ----------
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

@st.cache_data(show_spinner=False)
def fetch_fx_rate(base, quote):
    try:
        r = requests.get(f"https://api.exchangerate.host/convert?from={base}&to={quote}&amount=1", timeout=5)
        return float(r.json().get("info", {}).get("rate", 1))
    except Exception:
        return 1.0

# ---------- EXPORT HELPERS ----------
def fig_to_png_safe(fig):
    if st.session_state.get("disable_image_exports", False): return None
    try: return fig.to_image(format="png", engine="kaleido")
    except Exception: return None

def fig_to_pdf_safe(fig):
    if st.session_state.get("disable_image_exports", False): return None
    try: return fig.to_image(format="pdf", engine="kaleido")
    except Exception: return None

def pptx_with_chart_failsafe(fig, title="Chart"):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(1)).text = title
    img = fig_to_png_safe(fig)
    if img:
        slide.shapes.add_picture(BytesIO(img), Inches(0.8), Inches(1.4), width=Inches(8.8))
    else:
        slide.shapes.add_textbox(Inches(0.8), Inches(1.4), Inches(9), Inches(3)).text_frame.text = (
            "⚠️ Image export disabled or unsupported."
        )
    out = BytesIO(); prs.save(out); out.seek(0)
    return out.getvalue()

# ---------- FILE UPLOAD ----------
st.sidebar.header("📁 Upload File")
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

# ---------- TABS ----------
tab_data, tab_viz, tab_insight, tab_export, tab_settings = st.tabs(
    ["📄 Data", "📈 Visualize", "🧠 Insights", "📤 Export", "⚙️ Settings"]
)

# ---------- DATA TAB ----------
with tab_data:
    if not tables:
        st.info("Upload a file to begin.")
    else:
        key = st.selectbox("Select dataset", list(tables.keys()))
        df = tables[key].copy()
        st.success(f"Loaded {key} → {df.shape[0]} rows, {df.shape[1]} columns")
        st.dataframe(df.head(10))

# ---------- VISUALIZE TAB ----------
with tab_viz:
    if not tables:
        st.info("Upload data to visualize.")
    else:
        key = st.selectbox("Choose dataset for chart", list(tables.keys()))
        df = tables[key].copy()

        # --- Axis settings ---
        st.markdown("### Axis Settings")
        ignore_ai = st.checkbox("🧠 Ignore AI auto-detection", value=False)
        if ignore_ai:
            x_col = st.selectbox("X Axis Column", df.columns)
            y_cols = st.multiselect("Y Axis Columns", df.columns)
        else:
            date_cols, num_cols, txt_cols = detect_types(df)
            x_col = st.selectbox("X Axis (AI)", df.columns)
            y_cols = st.multiselect("Y Axis (AI)", num_cols, default=num_cols[:1] if num_cols else [])

        # --- Custom labels ---
        st.markdown("### Custom Axis Labels / Units")
        custom_x_label = st.text_input("X-Axis Label", value=x_col)
        custom_y_label = st.text_input("Y-Axis Label", value=", ".join(y_cols) if y_cols else "")
        c1, c2 = st.columns(2)
        with c1: prefix = st.text_input("Y-Axis Prefix", value="")
        with c2: suffix = st.text_input("Y-Axis Suffix", value="")
        scale_choice = st.selectbox(
            "Y-Axis Scale", ["None","Thousands (÷1,000)","Millions (÷1,000,000)","Billions (÷1,000,000,000)"]
        )
        scale_map = {"None":1, "Thousands (÷1,000)":1_000, "Millions (÷1,000,000)":1_000_000, "Billions (÷1,000,000,000)":1_000_000_000}

        # --- Chart Type ---
        chart_type = st.selectbox("Chart Type", ["Line","Area","Bar","Scatter","Pie"])

        # --- Date Range ---
        if x_col in df.columns and pd.api.types.is_datetime64_any_dtype(df[x_col]):
            df[x_col] = safe_to_datetime(df[x_col])
            df = df.dropna(subset=[x_col])
            if not df.empty:
                min_d, max_d = df[x_col].min(), df[x_col].max()
                if min_d == max_d: max_d = min_d + pd.Timedelta(days=1)
                try:
                    rng = st.date_input("Date range", [min_d.date(), max_d.date()])
                except Exception:
                    rng = [min_d.date(), max_d.date()]
                df = df[(df[x_col] >= pd.to_datetime(rng[0])) & (df[x_col] <= pd.to_datetime(rng[1]))]

        # --- Chart Creation ---
        if y_cols:
            df_plot = df.copy()
            for y in y_cols:
                if pd.api.types.is_numeric_dtype(df_plot[y]):
                    df_plot[y] = df_plot[y] / scale_map[scale_choice]

            if chart_type == "Line": fig = px.line(df_plot, x=x_col, y=y_cols, markers=True)
            elif chart_type == "Area": fig = px.area(df_plot, x=x_col, y=y_cols)
            elif chart_type == "Bar": fig = px.bar(df_plot, x=x_col, y=y_cols)
            elif chart_type == "Scatter": fig = px.scatter(df_plot, x=x_col, y=y_cols)
            elif chart_type == "Pie" and len(y_cols)==1: fig = px.pie(df_plot, names=x_col, values=y_cols[0])

            fig.update_layout(template="plotly_white", height=600,
                              xaxis_title=custom_x_label,
                              yaxis_title=f"{prefix}{custom_y_label}{suffix}")
            st.plotly_chart(fig, use_container_width=True)
            st.session_state["last_fig"], st.session_state["last_df"] = fig, df_plot

# ---------- INSIGHTS TAB ----------
with tab_insight:
    if not tables:
        st.info("Upload data to analyze.")
    else:
        key = st.selectbox("Choose dataset for insights", list(tables.keys()))
        df = tables[key].copy()
        _, num_cols, _ = detect_types(df)
        if num_cols:
            st.subheader("Summary Statistics")
            st.dataframe(df[num_cols].describe().T.round(3))
            st.subheader("Insights")
            insights = []
            for c in num_cols[:6]:
                s = pd.to_numeric(df[c], errors="coerce").dropna()
                if len(s)>1:
                    trend = "rising 📈" if s.iloc[-1]>s.iloc[0] else "falling 📉"
                    insights.append(f"**{c}** avg {s.mean():,.2f}, range {s.min():,.2f}–{s.max():,.2f}, {trend}")
            st.markdown("<br>".join(insights) or "No numeric insights found.", unsafe_allow_html=True)

# ---------- EXPORT TAB ----------
with tab_export:
    df_to_export = st.session_state.get("last_df")
    fig_to_export = st.session_state.get("last_fig")
    export_df = df_to_export if isinstance(df_to_export, pd.DataFrame) else pd.DataFrame()
    disabled = export_df.empty

    st.subheader("📦 Data Exports")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.download_button("⬇️ CSV", export_df.to_csv(index=False).encode("utf-8"),
                           "cleaned_data.csv", "text/csv", disabled=disabled)
    with c2:
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            export_df.to_excel(w, index=False, sheet_name="Data")
        st.download_button("⬇️ Excel", buf.getvalue(), "cleaned_data.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", disabled=disabled)
    with c3:
        st.download_button("⬇️ JSON", export_df.to_json(orient="records").encode("utf-8"),
                           "cleaned_data.json", "application/json", disabled=disabled)

    st.markdown("---")
    st.subheader("📊 Chart Exports")
    if fig_to_export:
        png, pdf = fig_to_png_safe(fig_to_export), fig_to_pdf_safe(fig_to_export)
        d1, d2, d3 = st.columns(3)
        with d1:
            st.download_button("⬇️ PNG", png or b"", "chart.png", "image/png", disabled=(png is None))
        with d2:
            st.download_button("⬇️ PDF", pdf or b"", "chart.pdf", "application/pdf", disabled=(pdf is None))
        with d3:
            pptx_bytes = pptx_with_chart_failsafe(fig_to_export)
            st.download_button("⬇️ PPTX", pptx_bytes, "chart_slide.pptx",
                               "application/vnd.openxmlformats-officedocument.presentationml.presentation")
    else:
        st.info("Create a chart first in the Visualize tab.")

# ---------- SETTINGS TAB ----------
with tab_settings:
    st.checkbox("Disable image exports (Option 4)",
                value=st.session_state["disable_image_exports"],
                key="disable_image_exports",
                help="Turn this on if your environment lacks Kaleido/Chrome. Data exports always work.")
    st.markdown("""
**Notes**
- Axis customization now supports prefixes, suffixes, and scaling.  
- 'Ignore AI' gives full manual control.  
- Heavy operations cached for maximum speed. 🚀
""")
