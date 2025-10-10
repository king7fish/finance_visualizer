# finance_dashboard_elite_v4.py
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
st.set_page_config(page_title="Finance Dashboard Elite v4", layout="wide")
PRIMARY = "#2563EB"
warnings.filterwarnings("ignore", category=UserWarning)

st.markdown(
    f"<h1 style='text-align:center;color:{PRIMARY};margin-bottom:4px'>🏆 Finance Dashboard Elite v4</h1>"
    "<p style='text-align:center;color:#6b7280;font-size:17px;'>Compare files & sheets • Custom colours • Bulletproof exports</p>",
    unsafe_allow_html=True,
)

# Session defaults
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
    """Fast, safe cleaning for messy real-world tables."""
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

def build_insights(df_long, x_col, series_col, value_col):
    """Narrative bullets + correlation matrix for plotted or raw (long or can be pivoted)."""
    bullets = []
    corr = None
    if df_long is None or df_long.empty:
        return ["Not enough data for insights."], None

    # Per-series stats & trend
    for sname, sdf in df_long.groupby(series_col):
        s = pd.to_numeric(sdf[value_col], errors="coerce").dropna()
        if len(s) >= 2:
            trend = "rising 📈" if s.iloc[-1] > s.iloc[0] else "falling 📉"
            bullets.append(f"**{sname}** — avg {s.mean():,.2f}, range {s.min():,.2f}–{s.max():,.2f}, {trend}.")
        elif len(s) == 1:
            bullets.append(f"**{sname}** — single value {s.iloc[0]:,.2f}.")
        else:
            bullets.append(f"**{sname}** — no numeric values after cleaning.")

    # If x is datetime, add YoY for first series with enough points
    if x_col in df_long.columns and pd.api.types.is_datetime64_any_dtype(df_long[x_col]):
        first_series = next(iter(df_long[series_col].unique()), None)
        if first_series is not None:
            tmp = df_long[df_long[series_col] == first_series][[x_col, value_col]].copy()
            tmp[x_col] = safe_to_datetime(tmp[x_col])
            tmp = tmp.dropna(subset=[x_col]).sort_values(x_col)
            if not tmp.empty:
                m = tmp.set_index(x_col)[value_col].resample("M").mean()
                if len(m) >= 13 and not pd.isna(m.iloc[-13]) and m.iloc[-13] != 0:
                    yoy = (m.iloc[-1] - m.iloc[-13]) / (abs(m.iloc[-13]) + 1e-9) * 100
                    bullets.append(f"**{first_series}** YoY change ≈ {yoy:+.1f}%.")

    # Correlations across series (pivot to wide)
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

# ----------- Export helpers (always-available) -----------
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
            ⚠️ Chart image export unavailable (using placeholder)."
        )
    out = BytesIO(); prs.save(out); out.seek(0)
    return out.getvalue()

def placeholder_png(text="No Chart Available", color=(0,0,0)):
    img = Image.new("RGB", (900, 560), color=(255, 255, 255))
    d = ImageDraw.Draw(img)
    d.text((280, 260), text, fill=color)
    buf = BytesIO(); img.save(buf, format="PNG"); return buf.getvalue()

# ---------------- FILE UPLOAD ----------------
st.sidebar.header("📁 Upload Data File")
u = st.sidebar.file_uploader(
    "Upload Excel / CSV / JSON / DOCX / PPTX / PDF",
    type=["xlsx","xls","csv","json","docx","pptx","pdf"],
    accept_multiple_files=True  # allow multiple files
)

@st.cache_data(show_spinner=True)
def load_files(uploaded_list):
    """Return dict: {label: DataFrame} including sheets/tables for each file."""
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
                    tables[f"{name} — {s}"] = smart_clean_dataframe(df)
            elif lower.endswith("csv"):
                tables[f"{name}"] = smart_clean_dataframe(pd.read_csv(uploaded))
            elif lower.endswith("json"):
                tables[f"{name}"] = smart_clean_dataframe(pd.read_json(uploaded))
            elif lower.endswith("docx"):
                doc = Document(uploaded)
                for i, t in enumerate(doc.tables):
                    rows = [[cell.text for cell in row.cells] for row in t.rows]
                    df = pd.DataFrame(rows[1:], columns=rows[0]) if len(rows)>1 else pd.DataFrame(rows)
                    tables[f"{name} — Table {i+1}"] = smart_clean_dataframe(df)
            elif lower.endswith("pptx"):
                prs = Presentation(uploaded)
                for i, slide in enumerate(prs.slides):
                    for shape in slide.shapes:
                        if hasattr(shape, "table"):
                            tbl = shape.table
                            rows = [[tbl.cell(r,c).text for c in range(len(tbl.columns))] for r in range(len(tbl.rows))]
                            df = pd.DataFrame(rows[1:], columns=rows[0]) if len(rows)>1 else pd.DataFrame(rows)
                            tables[f"{name} — Slide {i+1} Table"] = smart_clean_dataframe(df)
            elif lower.endswith("pdf"):
                pdf = PdfReader(uploaded)
                pages = [pg.extract_text() for pg in pdf.pages if pg.extract_text()]
                tables[f"{name} — PDF Text"] = pd.DataFrame({"Text": pages})
        except Exception as e:
            st.error(f"⚠️ File load failed for {name}: {e}")

    return tables

tables = load_files(u) if u else {}

# ---------------- TABS ----------------
tab_data, tab_viz, tab_insight, tab_export, tab_settings = st.tabs(
    ["📄 Data", "📈 Visualize", "🧠 Insights", "📤 Export", "⚙️ Settings"]
)

# ---------------- DATA TAB ----------------
with tab_data:
    if not tables:
        st.info("Upload files to begin.")
    else:
        key = st.selectbox("Preview a dataset (file or sheet/table)", list(tables.keys()))
        df = tables[key].copy()
        st.success(f"Loaded {key} → {df.shape[0]:,} rows × {df.shape[1]} columns")
        st.dataframe(df.head(12), use_container_width=True)

# ---------------- VISUALIZE TAB ----------------
with tab_viz:
    if not tables:
        st.info("Upload data to visualize.")
    else:
        st.subheader("🧭 Mode & Source Selection")
        compare_mode = st.toggle("Enable Compare Mode (multiple sources & sheets)", value=True)
        max_sources = 5  # reasonable default

        if compare_mode:
            sources = st.multiselect(
                "Pick up to 5 datasets to overlay",
                list(tables.keys()),
                default=list(tables.keys())[:min(2, len(tables))],
                max_selections=max_sources
            )
        else:
            sources = [st.selectbox("Choose one dataset", list(tables.keys()))]

        if not sources:
            st.info("Select at least one dataset.")
            st.stop()

        # ---------------- Axis and Labels ----------------
        st.subheader("🧩 Axis, Labels, Scaling")
        colA, colB = st.columns(2)
        with colA:
            custom_x_label = st.text_input("X-Axis Label", value="X")
            x_prefix = st.text_input("X Prefix", value="")
            x_suffix = st.text_input("X Suffix", value="")
            x_scale = st.selectbox("X Scale", ["None","Thousands (÷1,000)","Millions (÷1,000,000)","Billions (÷1,000,000,000)"])
        with colB:
            custom_y_label = st.text_input("Y-Axis Label", value="Value")
            y_prefix = st.text_input("Y Prefix", value="")
            y_suffix = st.text_input("Y Suffix", value="")
            y_scale = st.selectbox("Y Scale", ["None","Thousands (÷1,000)","Millions (÷1,000,000)","Billions (÷1,000,000,000)"])

        scale_map = {"None":1, "Thousands (÷1,000)":1_000, "Millions (÷1,000,000)":1_000_000, "Billions (÷1,000,000,000)":1_000_000_000}
        chart_type = st.selectbox("Chart Type", ["Line","Area","Bar","Scatter","Pie"])

        st.markdown("—")
        st.subheader("🔗 X-Axis Mapping (for each source)")
        st.caption("Pick which column in each dataset should map to the shared X-axis (e.g., Date or Category).")

        # For each source, choose X column (mapped to a common name) and choose Y columns + colours
        per_source_configs = []
        series_colors = {}
        master_x_name = "Master_X"  # internal combined x column

        # We’ll build a long format combined DataFrame for plotting
        combined_long = []

        for i, src in enumerate(sources, start=1):
            df_src = tables[src].copy()
            date_cols, num_cols, txt_cols = detect_types(df_src)

            c1, c2 = st.columns([1, 1])
            with c1:
                x_col = st.selectbox(f"X column for [{src}]", df_src.columns, key=f"xcol_{src}")
            with c2:
                label_suffix = st.text_input(f"Legend label suffix for [{src}] (optional)", value="", key=f"suffix_{src}")

            # Y columns choice for this source
            y_cols = st.multiselect(
                f"Y columns for [{src}]",
                df_src.columns,
                default=[c for c in df_src.columns if c in num_cols][:1],
                key=f"ycols_{src}"
            )

            # Color pickers per chosen Y column
            if y_cols:
                st.caption("Colours for this dataset's series:")
                col_blocks = st.columns(min(3, len(y_cols)) or 1)
                for idx, y in enumerate(y_cols):
                    col_block = col_blocks[idx % len(col_blocks)]
                    with col_block:
                        default_hex = ["#1f77b4","#d62728","#2ca02c","#9467bd","#ff7f0e"][idx % 5]
                        chosen = st.color_picker(f"{y}", value=default_hex, key=f"color_{src}_{y}")
                        # Build series name
                        series_name = f"{y}{(' ' + label_suffix) if label_suffix else ''}".strip()
                        series_colors[series_name] = chosen

            # Prepare standardized long data for this source
            if x_col and y_cols:
                tmp = df_src.copy()

                # Standardize X as datetime if it looks like date
                if pd.api.types.is_datetime64_any_dtype(tmp[x_col]) or looks_date(tmp[x_col].astype(str)):
                    tmp[x_col] = safe_to_datetime(tmp[x_col])
                    tmp = tmp.dropna(subset=[x_col])

                # Melt to long for chosen y_cols
                melt = tmp[[x_col] + y_cols].melt(id_vars=[x_col], var_name="Series", value_name="Value")
                # Add suffix to legend names
                if label_suffix:
                    melt["Series"] = melt["Series"].astype(str) + " " + label_suffix
                melt.rename(columns={x_col: master_x_name}, inplace=True)
                combined_long.append(melt)

        # Concatenate all sources (outer)
        if combined_long:
            combined = pd.concat(combined_long, axis=0, ignore_index=True)
        else:
            st.warning("Select at least one Y column across your sources.")
            st.stop()

        # Apply scaling to plotted values
        # X scaling only for numeric x
        if pd.api.types.is_numeric_dtype(combined[master_x_name]):
            combined[master_x_name] = combined[master_x_name] / scale_map[x_scale]

        # Y scaling
        combined["Value"] = pd.to_numeric(combined["Value"], errors="coerce")
        combined.loc[:, "Value"] = combined["Value"] / scale_map[y_scale]

        # Chart rendering
        fig = None
        try:
            if chart_type == "Pie":
                # Pie only makes sense for a single series or snapshot; constrain UX
                st.info("Pie charts are best with a single dataset & one metric. Showing aggregated totals by Series.")
                pie_df = combined.groupby("Series", as_index=False)["Value"].sum()
                fig = px.pie(pie_df, names="Series", values="Value", color="Series",
                             color_discrete_map=series_colors)
            else:
                if chart_type == "Line":
                    fig = px.line(combined, x=master_x_name, y="Value", color="Series",
                                  color_discrete_map=series_colors, markers=True)
                elif chart_type == "Area":
                    fig = px.area(combined, x=master_x_name, y="Value", color="Series",
                                  color_discrete_map=series_colors)
                elif chart_type == "Bar":
                    fig = px.bar(combined, x=master_x_name, y="Value", color="Series",
                                 color_discrete_map=series_colors)
                elif chart_type == "Scatter":
                    fig = px.scatter(combined, x=master_x_name, y="Value", color="Series",
                                     color_discrete_map=series_colors)
        except Exception as e:
            st.error(f"Chart failed: {e}")

        if fig is not None:
            fig.update_layout(
                template="plotly_white", height=650,
                xaxis_title=f"{x_prefix}{custom_x_label}{x_suffix}",
                yaxis_title=f"{y_prefix}{custom_y_label}{y_suffix}"
            )
            st.plotly_chart(fig, use_container_width=True)

            # Save context for Insights/Export
            st.session_state["last_fig"] = fig
            st.session_state["last_df_plot_long"] = combined.copy()
            # For raw combined (unscaled), rebuild quickly:
            raw_combined = pd.concat(combined_long, axis=0, ignore_index=True) if combined_long else pd.DataFrame()
            st.session_state["last_df_raw_long"] = raw_combined
            st.session_state["last_meta"] = {
                "x_col": master_x_name,
                "series_col": "Series",
                "value_col": "Value",
                "x_label": f"{x_prefix}{custom_x_label}{x_suffix}",
                "y_label": f"{y_prefix}{custom_y_label}{y_suffix}",
                "chart_type": chart_type,
                "sources": sources
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
        x_col = meta.get("x_col", "Master_X")
        series_col = meta.get("series_col", "Series")
        value_col = meta.get("value_col", "Value")

        if mode.startswith("Plotted"):
            df_src = st.session_state.get("last_df_plot_long")
        else:
            df_src = st.session_state.get("last_df_raw_long")

        if isinstance(df_src, pd.DataFrame) and not df_src.empty:
            bullets, corr = build_insights(df_src, x_col, series_col, value_col)
            st.subheader("💡 Narrative Insights")
            st.markdown("<br>".join(f"• {b}" for b in bullets), unsafe_allow_html=True)
            if corr is not None:
                st.subheader("🔗 Correlation (numeric series)")
                st.dataframe(corr)
        else:
            st.info("Make a chart in the Visualize tab first (or select data).")

# ---------------- EXPORT TAB ----------------
with tab_export:
    if not tables:
        st.info("Upload data to export.")
    else:
        fig_to_export = st.session_state.get("last_fig")
        df_plot_long = st.session_state.get("last_df_plot_long")
        df_raw_long = st.session_state.get("last_df_raw_long")

        st.subheader("📦 Export Data")
        which = st.radio("Choose data to export:", ["Plotted (filtered + scaled)", "Raw (cleaned)"], horizontal=True)
        export_df = df_plot_long if which.startswith("Plotted") else df_raw_long
        export_df = export_df if isinstance(export_df, pd.DataFrame) else pd.DataFrame()

        # Data exports (always enabled)
        c1, c2, c3 = st.columns(3)
        with c1:
            st.download_button("⬇️ CSV", export_df.to_csv(index=False).encode("utf-8"),
                               "data.csv", "text/csv")
        with c2:
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                export_df.to_excel(w, index=False, sheet_name="Data")
            st.download_button("⬇️ Excel", buf.getvalue(),
                               "data.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with c3:
            st.download_button("⬇️ JSON", export_df.to_json(orient="records").encode("utf-8"),
                               "data.json", "application/json")

        st.markdown("---")
        st.subheader("📊 Export Chart (always available)")

        # Chart exports (always available with placeholder fallbacks)
        if fig_to_export is None:
            st.warning("No chart found. Generating placeholder exports.")
            png_bytes = placeholder_png("No Chart Available", color=(0,0,0))
            pdf_bytes = png_bytes  # reuse for minimal environments
            pptx_bytes = pptx_with_chart_failsafe(None, title="Finance Dashboard — No Chart")
        else:
            png_bytes = fig_to_png_safe(fig_to_export)
            if not png_bytes:
                png_bytes = placeholder_png("Chart Export Failed", color=(255,0,0))
            pdf_try = fig_to_pdf_safe(fig_to_export)
            pdf_bytes = pdf_try if pdf_try else png_bytes
            pptx_bytes = pptx_with_chart_failsafe(fig_to_export, title="Finance Dashboard — Chart")

        d1, d2, d3 = st.columns(3)
        with d1:
            st.download_button("⬇️ PNG", png_bytes, "chart.png", "image/png")
        with d2:
            st.download_button("⬇️ PDF", pdf_bytes, "chart.pdf", "application/pdf")
        with d3:
            st.download_button("⬇️ PPTX", pptx_bytes, "chart_slide.pptx",
                               "application/vnd.openxmlformats-officedocument.presentationml.presentation")

        st.info("✅ All chart export types are always available. Fallback placeholders are auto-generated if needed.")

# ---------------- SETTINGS TAB ----------------
with tab_settings:
    st.checkbox("Disable image exports (Option 4)",
                value=st.session_state["disable_image_exports"],
                key="disable_image_exports",
                help="Turn on if your environment lacks Kaleido/Chrome. Data exports always work.")
    st.markdown("""
**Notes**
- Compare Mode overlays multiple files or sheets on one chart with custom colours.
- Insights can run on **Plotted** (filtered + scaled) or **Raw** (cleaned) long-form data.
- Exports are bulletproof: CSV, Excel, JSON, PNG, PDF, PPTX (with graceful fallbacks).
""")
