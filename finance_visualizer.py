# finance_dashboard_elite_plus_v2.py
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
    "<p style='text-align:center;color:#6b7280;font-size:17px;'>Dual-axis control. Analyst-grade insights & exports.</p>",
    unsafe_allow_html=True,
)

if "disable_image_exports" not in st.session_state:
    st.session_state["disable_image_exports"] = False

# ---------------- HELPERS ----------------
@st.cache_data(show_spinner=False)
def safe_to_datetime(series):
    """Convert any series to datetime without noisy warnings."""
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

# --- Insights engine ---
def build_insights(df, x_col, y_cols):
    """Return narrative bullets + (optional) correlation for df, based on current selections."""
    bullets = []
    corr = None

    # Basic stats & trend
    if y_cols:
        for c in y_cols[:8]:
            s = pd.to_numeric(df[c], errors="coerce").dropna()
            if len(s) >= 2:
                trend = "rising 📈" if s.iloc[-1] > s.iloc[0] else "falling 📉"
                bullets.append(f"**{c}** — avg {s.mean():,.2f}, range {s.min():,.2f}–{s.max():,.2f}, {trend}.")
            elif len(s) == 1:
                bullets.append(f"**{c}** — single value {s.iloc[0]:,.2f}.")
            else:
                bullets.append(f"**{c}** — no numeric values after cleaning.")

    # Time-series YoY (if x is datetime and enough data)
    if x_col in df.columns and pd.api.types.is_datetime64_any_dtype(df[x_col]):
        tmp = df[[x_col] + [c for c in y_cols if c in df.columns]].copy()
        tmp[x_col] = safe_to_datetime(tmp[x_col])
        tmp = tmp.dropna(subset=[x_col]).sort_values(x_col)
        if not tmp.empty and len(y_cols) >= 1:
            m = tmp.set_index(x_col)[y_cols[0]].resample("M").mean()
            if len(m) >= 13 and m.iloc[-13] != 0:
                yoy = (m.iloc[-1] - m.iloc[-13]) / (abs(m.iloc[-13]) + 1e-9) * 100
                bullets.append(f"**{y_cols[0]}** YoY change ≈ {yoy:+.1f}%.")

    # Correlation (numeric only)
    num_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    if len(num_cols) >= 2:
        corr = df[num_cols].corr().round(3)

    if not bullets:
        bullets = ["Not enough usable data for narrative insights."]
    return bullets, corr

# ---------------- EXPORT HELPERS ----------------
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
        st.dataframe(df.head(12), use_container_width=True)

# ---------------- VISUALIZE TAB ----------------
with tab_viz:
    if not tables:
        st.info("Upload data to visualize.")
    else:
        key = st.selectbox("Choose dataset for chart", list(tables.keys()))
        raw_df = tables[key].copy()   # keep a raw-cleaned snapshot for exports
        date_cols, num_cols, txt_cols = detect_types(raw_df)

        # === Axis Configuration ===
        st.subheader("🧭 Axis Configuration")
        ignore_ai = st.toggle("Ignore AI auto-detection (manual control)", value=False)

        if ignore_ai:
            x_col = st.selectbox("X-Axis Column", raw_df.columns)
            y_cols = st.multiselect("Y-Axis Columns", raw_df.columns)
        else:
            x_col = st.selectbox("X-Axis (detected)", raw_df.columns)
            y_cols = st.multiselect("Y-Axis (detected numeric)", num_cols, default=num_cols[:1] if num_cols else [])

        # === Labels & Scaling ===
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

        # === Filter & Prepare Data ===
        df = raw_df.copy()

        # Time filter
        if x_col in df.columns and pd.api.types.is_datetime64_any_dtype(df[x_col]):
            df[x_col] = safe_to_datetime(df[x_col])
            df = df.dropna(subset=[x_col])
            if not df.empty:
                min_d, max_d = df[x_col].min(), df[x_col].max()
                if min_d == max_d: max_d = min_d + pd.Timedelta(days=1)
                try:
                    rng = st.date_input("Date Range", [min_d.date(), max_d.date()])
                except Exception:
                    rng = [min_d.date(), max_d.date()]
                df = df[(df[x_col] >= pd.to_datetime(rng[0])) & (df[x_col] <= pd.to_datetime(rng[1]))]

        # Create plotted dataframe (apply scaling, not prefixes/suffixes)
        df_plot = df.copy()
        if x_col in df_plot.columns and pd.api.types.is_numeric_dtype(df_plot[x_col]):
            df_plot[x_col] = df_plot[x_col] / scale_map[x_scale]
        for y in y_cols:
            if y in df_plot.columns and pd.api.types.is_numeric_dtype(df_plot[y]):
                df_plot[y] = df_plot[y] / scale_map[y_scale]

        # === Chart Rendering ===
        fig = None
        if y_cols:
            try:
                if chart_type == "Line":    fig = px.line(df_plot, x=x_col, y=y_cols, markers=True)
                elif chart_type == "Area":  fig = px.area(df_plot, x=x_col, y=y_cols)
                elif chart_type == "Bar":   fig = px.bar(df_plot, x=x_col, y=y_cols)
                elif chart_type == "Scatter": fig = px.scatter(df_plot, x=x_col, y=y_cols)
                elif chart_type == "Pie" and len(y_cols)==1: fig = px.pie(df_plot, names=x_col, values=y_cols[0])
            except Exception as e:
                st.error(f"Chart failed: {e}")

        if fig is not None:
            fig.update_layout(
                template="plotly_white", height=600,
                xaxis_title=f"{x_prefix}{custom_x_label}{x_suffix}",
                yaxis_title=f"{y_prefix}{custom_y_label}{y_suffix}"
            )
            st.plotly_chart(fig, use_container_width=True)

            # --- Save context for Insights/Export ---
            st.session_state["last_fig"] = fig
            st.session_state["last_df_plot"] = df_plot
            st.session_state["last_df_raw"] = raw_df
            st.session_state["last_meta"] = {
                "x_col": x_col, "y_cols": y_cols,
                "x_label": f"{x_prefix}{custom_x_label}{x_suffix}",
                "y_label": f"{y_prefix}{custom_y_label}{y_suffix}",
                "x_scale": x_scale, "y_scale": y_scale,
                "chart_type": chart_type, "dataset_key": key
            }
        else:
            st.info("Select at least one Y column to render a chart.")

# ---------------- INSIGHTS TAB ----------------
with tab_insight:
    if not tables:
        st.info("Upload data to analyze.")
    else:
        mode = st.radio("Run insights on:", ["Plotted (filtered + scaled)", "Raw (cleaned)"], horizontal=True)
        meta = st.session_state.get("last_meta", {})
        x_col = meta.get("x_col")
        y_cols = meta.get("y_cols", [])
        dataset_key = meta.get("dataset_key")

        if mode.startswith("Plotted"):
            df_src = st.session_state.get("last_df_plot")
        else:
            # Use the same dataset the user used in Visualize, if available
            if dataset_key and dataset_key in tables:
                df_src = tables[dataset_key].copy()
            else:
                # fallback: let the user pick a dataset
                st.warning("No prior chart context found. Choose a dataset:")
                pick = st.selectbox("Dataset", list(tables.keys()))
                df_src = tables[pick].copy()
                # allow manual selection if no meta
                cols = list(df_src.columns)
                x_col = st.selectbox("X column", cols)
                y_cols = st.multiselect("Y columns", cols)

        if isinstance(df_src, pd.DataFrame) and x_col and y_cols:
            bullets, corr = build_insights(df_src, x_col, y_cols)
            st.subheader("💡 Narrative Insights")
            st.markdown("<br>".join(f"• {b}" for b in bullets), unsafe_allow_html=True)
            if corr is not None:
                st.subheader("🔗 Correlation (numeric columns)")
                st.dataframe(corr)
        else:
            st.info("Create a chart first in the Visualize tab, or select columns above.")

# ---------------- EXPORT TAB ----------------
with tab_export:
    if not tables:
        st.info("Upload data to export.")
    else:
        fig_to_export = st.session_state.get("last_fig")
        df_plot = st.session_state.get("last_df_plot")
        df_raw = st.session_state.get("last_df_raw")

        st.subheader("📦 Export Data")
        which = st.radio("Choose data to export:", ["Plotted (filtered + scaled)", "Raw (cleaned)"], horizontal=True)
        export_df = df_plot if which.startswith("Plotted") else df_raw
        if not isinstance(export_df, pd.DataFrame):
            export_df = pd.DataFrame()

        disabled = export_df.empty
        c1, c2, c3 = st.columns(3)
        with c1:
            st.download_button("⬇️ CSV", export_df.to_csv(index=False).encode("utf-8"),
                               "data.csv", "text/csv", disabled=disabled)
        with c2:
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                export_df.to_excel(w, index=False, sheet_name="Data")
            st.download_button("⬇️ Excel", buf.getvalue(),
                               "data.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               disabled=disabled)
        with c3:
            st.download_button("⬇️ JSON", export_df.to_json(orient="records").encode("utf-8"),
                               "data.json", "application/json", disabled=disabled)

        st.markdown("---")
        st.subheader("📊 Export Chart")
        if fig_to_export is None:
            st.info("Create a chart in the Visualize tab first.")
        else:
            png, pdf = None, None
            if not st.session_state.get("disable_image_exports", False):
                png, pdf = fig_to_png_safe(fig_to_export), None
                # PDF often fails on minimal cloud images; keep PNG+PPTX reliable
                try:
                    pdf = fig_to_pdf_safe(fig_to_export)
                except Exception:
                    pdf = None
            d1, d2, d3 = st.columns(3)
            with d1:
                st.download_button("⬇️ PNG", png or b"", "chart.png", "image/png", disabled=(png is None))
            with d2:
                st.download_button("⬇️ PDF", pdf or b"", "chart.pdf", "application/pdf", disabled=(pdf is None))
            with d3:
                pptx_bytes = pptx_with_chart_failsafe(fig_to_export, title="Finance Dashboard — Chart")
                st.download_button("⬇️ PPTX", pptx_bytes,
                                   "chart_slide.pptx",
                                   "application/vnd.openxmlformats-officedocument.presentationml.presentation")

# ---------------- SETTINGS TAB ----------------
with tab_settings:
    st.checkbox("Disable image exports (Option 4)",
                value=st.session_state["disable_image_exports"],
                key="disable_image_exports",
                help="Turn on if your environment lacks Kaleido/Chrome. Data exports always work.")
    st.markdown("""
**Tips**
- Insights can now run on the **plotted** dataset (your exact chart context) or the **raw cleaned** dataset.
- Exports let you choose Plotted vs Raw to match your workflow.
- Axis scaling, prefixes, and suffixes do not affect the raw export—only the plotted export.
""")
