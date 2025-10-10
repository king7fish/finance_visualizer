# finance_dashboard_elite_v5.py
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import warnings, re
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches
from docx import Document
from PyPDF2 import PdfReader
from dateutil import parser
from PIL import Image, ImageDraw

# ---------------- CONFIG ----------------
st.set_page_config(page_title="Finance Dashboard Elite v5", layout="wide")
PRIMARY = "#2563EB"
warnings.filterwarnings("ignore", category=UserWarning)

# Minimal CSS polish (cards, headings, spacing)
st.markdown("""
<style>
/* Page width */
.block-container { max-width: 1200px; }

/* Section cards */
.card {
  border: 1px solid #e5e7eb;
  border-radius: 10px;
  padding: 16px 18px;
  background: #ffffff;
  box-shadow: 0 1px 2px rgba(0,0,0,0.03);
  margin-bottom: 14px;
}

/* Subheaders */
h3, h4 { margin-top: 6px; }

/* Help text tone */
.small { color:#6b7280; font-size: 13px; }

/* Tighten widgets a bit */
.css-1kyxreq, .stTextInput, .stSelectbox, .stMultiSelect, .stColorPicker { margin-bottom: 8px; }
</style>
""", unsafe_allow_html=True)

st.markdown(
    "<h1 style='text-align:center;color:#2563EB;margin-bottom:6px'>Finance Dashboard Elite v5</h1>"
    "<p style='text-align:center;color:#6b7280;font-size:16px;'>Compare files and sheets. Custom colours. Reliable exports. Clean UI.</p>",
    unsafe_allow_html=True,
)

# Session flags
if "disable_image_exports" not in st.session_state:
    st.session_state["disable_image_exports"] = False

# ---------------- HELPERS ----------------
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

def smart_clean_dataframe(df_in):
    df = df_in.dropna(how="all").copy()
    # strip weird chars from headers
    df.columns = [str(c).strip().replace("\n", " ").replace("\xa0", " ").replace("\ufeff", "") for c in df.columns]
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

def build_insights(df_long, x_col, series_col, value_col):
    bullets = []
    corr = None
    if df_long is None or df_long.empty:
        return ["Not enough data for insights."], None

    # Per-series stats
    for sname, sdf in df_long.groupby(series_col):
        s = pd.to_numeric(sdf[value_col], errors="coerce").dropna()
        if len(s) >= 2:
            trend = "rising" if s.iloc[-1] > s.iloc[0] else "falling"
            bullets.append(f"{sname}: avg {s.mean():,.2f}, range {s.min():,.2f} to {s.max():,.2f}, {trend}.")
        elif len(s) == 1:
            bullets.append(f"{sname}: single value {s.iloc[0]:,.2f}.")
        else:
            bullets.append(f"{sname}: no numeric values after cleaning.")

    # YoY for first series if datetime
    if x_col in df_long.columns and pd.api.types.is_datetime64_any_dtype(df_long[x_col]):
        uniq = list(df_long[series_col].unique())
        if uniq:
            first_series = uniq[0]
            tmp = df_long[df_long[series_col] == first_series][[x_col, value_col]].copy()
            tmp[x_col] = safe_to_datetime(tmp[x_col])
            tmp = tmp.dropna(subset=[x_col]).sort_values(x_col)
            if not tmp.empty:
                m = tmp.set_index(x_col)[value_col].resample("M").mean()
                if len(m) >= 13 and not pd.isna(m.iloc[-13]) and m.iloc[-13] != 0:
                    yoy = (m.iloc[-1] - m.iloc[-13]) / (abs(m.iloc[-13]) + 1e-9) * 100
                    bullets.append(f"{first_series}: approx YoY change {yoy:+.1f}%.")

    # Correlation across series (pivot to wide)
    try:
        wide = df_long.pivot_table(index=x_col, columns=series_col, values=value_col, aggfunc="mean")
        num_cols = wide.select_dtypes(include=["number"])
        if num_cols.shape[1] >= 2:
            corr = num_cols.corr().round(3)
    except Exception:
        corr = None

    if not bullets:
        bullets = ["Not enough usable data for narrative insights."]
    return bullets, corr

# ----------- Export helpers -----------
def fig_to_png_safe(fig):
    if st.session_state.get("disable_image_exports", False): return None
    try:
        return fig.to_image(format="png", engine="kaleido")
    except Exception:
        return None

def fig_to_pdf_safe(fig):
    if st.session_state.get("disable_image_exports", False): return None
    try:
        return fig.to_image(format="pdf", engine="kaleido")
    except Exception:
        return None

def pptx_with_chart_failsafe(fig, title="Chart"):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(1)).text = title
    png = fig_to_png_safe(fig) if fig is not None else None
    if png:
        slide.shapes.add_picture(BytesIO(png), Inches(0.8), Inches(1.4), width=Inches(8.8))
    else:
        slide.shapes.add_textbox(Inches(0.8), Inches(1.4), Inches(9), Inches(3)).text_frame.text = (
            "Chart image export unavailable (using placeholder)."
        )
    out = BytesIO(); prs.save(out); out.seek(0)
    return out.getvalue()

def placeholder_png(text="No Chart Available", color=(0,0,0)):
    img = Image.new("RGB", (900, 560), color=(255, 255, 255))
    d = ImageDraw.Draw(img)
    d.text((280, 260), text, fill=color)
    buf = BytesIO(); img.save(buf, format="PNG"); return buf.getvalue()

# ---------------- FILE UPLOAD ----------------
with st.sidebar:
    st.header("Upload Data Files")
    uploaded_files = st.file_uploader(
        "Excel / CSV / JSON / DOCX / PPTX / PDF",
        type=["xlsx","xls","csv","json","docx","pptx","pdf"],
        accept_multiple_files=True
    )

@st.cache_data(show_spinner=True)
def load_files(uploaded_list):
    tables = {}
    if not uploaded_list:
        return tables
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

# ---------------- TABS ----------------
tab_data, tab_viz, tab_insight, tab_export, tab_settings = st.tabs(
    ["Data", "Visualize", "Insights", "Export", "Settings"]
)

# ---------------- DATA TAB ----------------
with tab_data:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    if not tables:
        st.info("Upload files to begin.")
    else:
        key = st.selectbox("Preview a dataset (file or sheet/table)", list(tables.keys()))
        df_prev = tables[key].copy()
        st.success(f"Loaded {key} - {df_prev.shape[0]:,} rows x {df_prev.shape[1]} columns")
        st.dataframe(df_prev.head(12), use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ---------------- VISUALIZE TAB ----------------
with tab_viz:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Mode and Sources")
    compare_mode = st.checkbox("Enable Compare Mode (multi-file and multi-sheet)", value=True)
    max_sources = 5

    if not tables:
        st.info("Upload data to visualize.")
        st.markdown('</div>', unsafe_allow_html=True)
        st.stop()

    if compare_mode:
        sources = st.multiselect(
            "Pick up to 5 datasets to overlay",
            list(tables.keys()),
            default=list(tables.keys())[:min(2, len(tables))]
        )
        if len(sources) > max_sources:
            st.warning("Only the first 5 selections will be used.")
            sources = sources[:max_sources]
    else:
        sources = [st.selectbox("Choose one dataset", list(tables.keys()))]

    if not sources:
        st.info("Select at least one dataset.")
        st.markdown('</div>', unsafe_allow_html=True)
        st.stop()

    st.markdown('</div>', unsafe_allow_html=True)

    # Axis panel (clean UI)
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Axes and Labels")
    c1, c2 = st.columns(2)
    with c1:
        custom_x_label = st.text_input("X-Axis Label", value="X")
        x_prefix = st.text_input("X Prefix", value="")
        x_suffix = st.text_input("X Suffix", value="")
    with c2:
        custom_y_label = st.text_input("Y-Axis Label", value="Value")
        y_prefix = st.text_input("Y Prefix", value="")
        y_suffix = st.text_input("Y Suffix", value="")

    with st.expander("Units and Scaling"):
        s1, s2, s3 = st.columns(3)
        with s1:
            x_scale = st.selectbox("X Scale", ["None","Thousands (/1,000)","Millions (/1,000,000)","Billions (/1,000,000,000)"])
        with s2:
            y_scale = st.selectbox("Y Scale", ["None","Thousands (/1,000)","Millions (/1,000,000)","Billions (/1,000,000,000)"])
        with s3:
            chart_type = st.selectbox("Chart Type", ["Line","Area","Bar","Scatter","Pie"])
    st.markdown('</div>', unsafe_allow_html=True)

    scale_map = {"None":1, "Thousands (/1,000)":1_000, "Millions (/1,000,000)":1_000_000, "Billions (/1,000,000,000)":1_000_000_000}

    # Per-source mapping panel
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("X Mapping and Series Selection")
    st.markdown('<p class="small">For each source, choose an X column and one or more Y columns. Pick colours for each series.</p>', unsafe_allow_html=True)

    master_x_name = "Master_X"
    combined_long = []
    raw_combined_long = []
    series_colors = {}

    for src in sources:
        df_src = tables[src].copy()
        # ensure cleaned headers
        df_src.columns = [str(c).strip().replace("\xa0", " ").replace("\ufeff", "") for c in df_src.columns]
        date_cols, num_cols, txt_cols = detect_types(df_src)

        b1, b2 = st.columns([1, 1])
        with b1:
            x_col = st.selectbox(f"X column for [{src}]", df_src.columns, key=f"xcol_{src}")
        with b2:
            label_suffix = st.text_input(f"Legend suffix for [{src}] (optional)", value="", key=f"suffix_{src}")

        y_cols = st.multiselect(
            f"Y columns for [{src}]",
            df_src.columns,
            default=[c for c in df_src.columns if c in num_cols][:1],
            key=f"ycols_{src}"
        )

        # Color pickers per Y
        if y_cols:
            cols = st.columns(min(3, len(y_cols)) or 1)
            palette = ["#1f77b4","#d62728","#2ca02c","#9467bd","#ff7f0e"]
            for idx, y in enumerate(y_cols):
                with cols[idx % len(cols)]:
                    chosen = st.color_picker(f"Colour for {y} [{src}]", value=palette[idx % len(palette)], key=f"color_{src}_{y}")
                    series_name = f"{y}{(' ' + label_suffix) if label_suffix else ''}".strip()
                    series_colors[series_name] = chosen

        # --- Robust melt with validation (fixes your error) ---
        if x_col and y_cols:
            tmp = df_src.copy()
            tmp.columns = [str(c).strip().replace("\xa0", " ").replace("\ufeff", "") for c in tmp.columns]

            # validate columns
            valid_y = [y for y in y_cols if y in tmp.columns]
            if x_col not in tmp.columns:
                st.warning(f"Skipping [{src}] because X column '{x_col}' not found after cleaning.")
                continue
            if not valid_y:
                st.warning(f"Skipping [{src}] because none of the chosen Y columns exist after cleaning.")
                continue

            # normalize X if looks like date
            if pd.api.types.is_datetime64_any_dtype(tmp[x_col]) or looks_date(tmp[x_col].astype(str)):
                tmp[x_col] = safe_to_datetime(tmp[x_col])
                tmp = tmp.dropna(subset=[x_col])

            try:
                melt_raw = tmp[[x_col] + valid_y].melt(
                    id_vars=[x_col],
                    var_name="Series",
                    value_name="Value"
                )
            except Exception as e:
                st.warning(f"Skipping [{src}] due to melt error: {e}")
                continue

            if label_suffix:
                melt_raw["Series"] = melt_raw["Series"].astype(str) + " " + label_suffix

            melt_raw.rename(columns={x_col: master_x_name}, inplace=True)
            raw_combined_long.append(melt_raw)

    # Combine all long data
    if raw_combined_long:
        raw_long = pd.concat(raw_combined_long, axis=0, ignore_index=True)
    else:
        st.warning("Select at least one valid Y column across your sources.")
        st.markdown('</div>', unsafe_allow_html=True)
        st.stop()

    # Apply scaling for plotted version
    plotted = raw_long.copy()
    if pd.api.types.is_numeric_dtype(plotted[master_x_name]):
        plotted[master_x_name] = plotted[master_x_name] / scale_map[x_scale]
    plotted["Value"] = pd.to_numeric(plotted["Value"], errors="coerce") / scale_map[y_scale]

    # Date filter if X is datetime
    if pd.api.types.is_datetime64_any_dtype(raw_long[master_x_name]):
        try:
            min_d, max_d = raw_long[master_x_name].min(), raw_long[master_x_name].max()
            if pd.isna(min_d) or pd.isna(max_d) or min_d == max_d:
                pass
            else:
                d1, d2 = st.columns(2)
                with d1:
                    date_from = st.date_input("From", min_d.date())
                with d2:
                    date_to = st.date_input("To", max_d.date())
                plotted = plotted[(plotted[master_x_name] >= pd.to_datetime(date_from)) &
                                  (plotted[master_x_name] <= pd.to_datetime(date_to))]
        except Exception:
            pass

    # Draw chart
    fig = None
    try:
        if chart_type == "Pie":
            st.info("Pie shows aggregated totals by series.")
            pie_df = plotted.groupby("Series", as_index=False)["Value"].sum()
            fig = px.pie(pie_df, names="Series", values="Value", color="Series",
                         color_discrete_map=series_colors)
        else:
            args = dict(x=master_x_name, y="Value", color="Series", color_discrete_map=series_colors)
            if chart_type == "Line":
                fig = px.line(plotted, markers=True, **args)
            elif chart_type == "Area":
                fig = px.area(plotted, **args)
            elif chart_type == "Bar":
                fig = px.bar(plotted, **args)
            elif chart_type == "Scatter":
                fig = px.scatter(plotted, **args)
    except Exception as e:
        st.error(f"Chart failed: {e}")

    if fig is not None:
        fig.update_layout(
            template="plotly_white", height=650,
            xaxis_title=f"{x_prefix}{custom_x_label}{x_suffix}",
            yaxis_title=f"{y_prefix}{custom_y_label}{y_suffix}"
        )
        st.plotly_chart(fig, use_container_width=True)

        # Store context for other tabs
        st.session_state["last_fig"] = fig
        st.session_state["last_df_plot_long"] = plotted.copy()
        st.session_state["last_df_raw_long"] = raw_long.copy()
        st.session_state["last_meta"] = {
            "x_col": master_x_name,
            "series_col": "Series",
            "value_col": "Value",
            "chart_type": chart_type,
            "sources": sources
        }
    else:
        st.info("Select at least one Y column to render a chart.")
    st.markdown('</div>', unsafe_allow_html=True)

# ---------------- INSIGHTS TAB ----------------
with tab_insight:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    if not tables:
        st.info("Upload data to analyze.")
    else:
        mode = st.radio("Run insights on:", ["Plotted (filtered + scaled)", "Raw (cleaned)"], horizontal=True)
        meta = st.session_state.get("last_meta", {})
        x_col = meta.get("x_col", "Master_X")
        series_col = meta.get("series_col", "Series")
        value_col = meta.get("value_col", "Value")

        src_df = st.session_state.get("last_df_plot_long") if mode.startswith("Plotted") else st.session_state.get("last_df_raw_long")
        if isinstance(src_df, pd.DataFrame) and not src_df.empty:
            bullets, corr = build_insights(src_df, x_col, series_col, value_col)
            st.subheader("Narrative Insights")
            st.markdown("<br>".join(f"- {b}" for b in bullets), unsafe_allow_html=True)
            if corr is not None:
                st.subheader("Correlation (numeric series)")
                st.dataframe(corr)
        else:
            st.info("Make a chart in the Visualize tab first.")
    st.markdown('</div>', unsafe_allow_html=True)

# ---------------- EXPORT TAB ----------------
with tab_export:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    if not tables:
        st.info("Upload data to export.")
    else:
        fig_to_export = st.session_state.get("last_fig")
        df_plot_long = st.session_state.get("last_df_plot_long")
        df_raw_long = st.session_state.get("last_df_raw_long")

        st.subheader("Export Data")
        which = st.radio("Choose data to export:", ["Plotted (filtered + scaled)", "Raw (cleaned)"], horizontal=True)
        export_df = df_plot_long if which.startswith("Plotted") else df_raw_long
        export_df = export_df if isinstance(export_df, pd.DataFrame) else pd.DataFrame()

        c1, c2, c3 = st.columns(3)
        with c1:
            st.download_button("Download CSV", export_df.to_csv(index=False).encode("utf-8"),
                               "data.csv", "text/csv")
        with c2:
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                export_df.to_excel(w, index=False, sheet_name="Data")
            st.download_button("Download Excel", buf.getvalue(),
                               "data.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with c3:
            st.download_button("Download JSON", export_df.to_json(orient="records").encode("utf-8"),
                               "data.json", "application/json")

        st.markdown("---")
        st.subheader("Export Chart (always available)")

        if fig_to_export is None:
            st.warning("No chart found. Generating placeholder exports.")
            png_bytes = placeholder_png("No Chart Available", color=(0,0,0))
            pdf_bytes = png_bytes
            pptx_bytes = pptx_with_chart_failsafe(None, title="Finance Dashboard - No Chart")
        else:
            png_bytes = fig_to_png_safe(fig_to_export)
            if not png_bytes:
                png_bytes = placeholder_png("Chart Export Failed", color=(255,0,0))
            pdf_try = fig_to_pdf_safe(fig_to_export)
            pdf_bytes = pdf_try if pdf_try else png_bytes
            pptx_bytes = pptx_with_chart_failsafe(fig_to_export, title="Finance Dashboard - Chart")

        d1, d2, d3 = st.columns(3)
        with d1:
            st.download_button("Download PNG", png_bytes, "chart.png", "image/png")
        with d2:
            st.download_button("Download PDF", pdf_bytes, "chart.pdf", "application/pdf")
        with d3:
            st.download_button("Download PPTX", pptx_bytes, "chart_slide.pptx",
                               "application/vnd.openxmlformats-officedocument.presentationml.presentation")

        st.info("All chart export types are always available. Fallback placeholders are auto-generated if needed.")
    st.markdown('</div>', unsafe_allow_html=True)

# ---------------- SETTINGS TAB ----------------
with tab_settings:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.checkbox("Disable image exports (Option 4)",
                value=st.session_state["disable_image_exports"],
                key="disable_image_exports",
                help="Turn on if your environment lacks Kaleido/Chrome. Data exports always work.")
    st.markdown("""
Notes
- Compare Mode overlays multiple files or sheets with custom colours.
- Insights can run on Plotted (filtered + scaled) or Raw (cleaned) long-form data.
- Exports are guaranteed: CSV, Excel, JSON, PNG, PDF, PPTX.
""")
    st.markdown('</div>', unsafe_allow_html=True)
