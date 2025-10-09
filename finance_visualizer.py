import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches
from docx import Document
from PyPDF2 import PdfReader
import re, requests, traceback

# ---------- Streamlit setup ----------
st.set_page_config(page_title="Finance Dashboard Pro", layout="wide")
PRIMARY = "#3A86FF"
st.markdown(
    f"<h1 style='text-align:center;color:{PRIMARY}'>🏆 Finance Dashboard Pro</h1>"
    "<p style='text-align:center;color:#6b7280'>Messy data in → Analyst-level insights out</p>",
    unsafe_allow_html=True,
)

if "disable_image_exports" not in st.session_state:
    st.session_state["disable_image_exports"] = False

# ---------- Helper utilities ----------
def looks_numeric(s):
    if s is None or not len(s): return False
    vals = s.dropna().astype(str).head(60)
    patt = re.compile(r"^\s*[-+]?\d{1,3}(?:,\d{3})*(?:\.\d+)?\s*$|^\s*[-+]?\d+(?:\.\d+)?\s*$")
    return sum(bool(patt.match(v)) for v in vals) >= 0.5 * len(vals)

def looks_date(s):
    vals = s.dropna().astype(str).head(60)
    return sum(("/" in v or "-" in v) for v in vals) >= 0.5 * len(vals)

def promote_header_if_found(df_raw):
    df = df_raw.copy()
    for i in range(min(10, len(df))):
        row = df.iloc[i].astype(str).str.strip()
        if (row != "").sum() >= max(2, int(0.5 * len(row))):
            headers = df.iloc[i].astype(str).str.strip().tolist()
            new = df.iloc[i + 1 :].copy()
            new.columns = [h if h and h.lower() != "unnamed" else f"col_{j}" for j, h in enumerate(headers)]
            return new.reset_index(drop=True)
    df.columns = [f"col_{i}" if not str(c).strip() else str(c) for i, c in enumerate(df.columns)]
    return df

def smart_clean_dataframe(df_in):
    """Safely normalise columns; no .str errors, no nested data issues."""
    df = df_in.copy().dropna(how="all")
    df.columns = [str(c).strip().replace("\n", " ") for c in df.columns]

    for col in df.columns:
        s = df[col]
        try:
            if not isinstance(s, pd.Series): s = pd.Series(s)
            s = s.apply(lambda v: str(v) if not isinstance(v, (dict, list, set)) else str(v))
        except Exception:
            s = pd.Series([str(v) for v in df[col].values])

        if looks_numeric(s):
            cleaned = s.astype(str).str.replace(",", "", regex=False)
            cleaned = cleaned.str.replace(r"[^0-9.\-]", "", regex=True).replace("", np.nan)
            df[col] = pd.to_numeric(cleaned, errors="coerce")
        elif looks_date(s):
            df[col] = pd.to_datetime(s, errors="coerce")
        else:
            df[col] = s.astype(str).str.strip()
    return df.reset_index(drop=True)

def detect_types(df):
    date_cols, num_cols, txt_cols = [], [], []
    for c in df.columns:
        s = df[c]
        if pd.api.types.is_datetime64_any_dtype(s) or looks_date(s): date_cols.append(c)
        elif pd.api.types.is_numeric_dtype(s) or looks_numeric(s):   num_cols.append(c)
        else:                                                        txt_cols.append(c)
    return date_cols, num_cols, txt_cols

def safe_resample(df, time_col, freq, how):
    df = df.copy()
    df[time_col] = pd.to_datetime(df[time_col], errors="coerce")
    df = df.dropna(subset=[time_col])
    if df.empty: return df
    num = df.select_dtypes(include=["number"])
    if num.empty: return df
    agg = how if how in ["sum","mean","median","max","min"] else "mean"
    return num.resample(freq, on=time_col).agg(agg).reset_index()

def fetch_fx_rate(base, quote):
    try:
        r = requests.get(f"https://api.exchangerate.host/convert?from={base}&to={quote}&amount=1", timeout=8)
        return float(r.json().get("info", {}).get("rate", None))
    except Exception:
        return None

# ---------- Export helpers ----------
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
    slide.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(9), Inches(1)).text = title
    img = fig_to_png_safe(fig)
    if img:
        slide.shapes.add_picture(BytesIO(img), Inches(0.8), Inches(1.4), width=Inches(8.8))
    else:
        slide.shapes.add_textbox(Inches(0.8), Inches(1.4), Inches(9), Inches(3)).text_frame.text = (
            "⚠️ Chart image export unavailable (Kaleido missing or disabled)."
        )
    buf = BytesIO(); prs.save(buf); buf.seek(0)
    return buf.getvalue()

# ---------- File upload ----------
st.sidebar.header("📁 Upload File")
u = st.sidebar.file_uploader("Upload Excel/CSV/JSON/DOCX/PPTX/PDF",
                             type=["xlsx","xls","csv","json","docx","pptx","pdf"])

tables = {}
if u:
    try:
        name = u.name.lower()
        if name.endswith(("xlsx","xls")):
            xls = pd.ExcelFile(u)
            for s in xls.sheet_names:
                raw = xls.parse(s, header=None)
                tables[f"Excel — {s}"] = smart_clean_dataframe(promote_header_if_found(raw))
        elif name.endswith("csv"):
            tables["CSV"] = smart_clean_dataframe(pd.read_csv(u))
        elif name.endswith("json"):
            tables["JSON"] = smart_clean_dataframe(pd.read_json(u))
        elif name.endswith("docx"):
            doc = Document(u)
            for i, t in enumerate(doc.tables):
                rows = [[cell.text for cell in row.cells] for row in t.rows]
                df = pd.DataFrame(rows[1:], columns=rows[0]) if len(rows)>1 else pd.DataFrame(rows)
                tables[f"Word Table {i+1}"] = smart_clean_dataframe(df)
        elif name.endswith("pptx"):
            prs = Presentation(u)
            for i, slide in enumerate(prs.slides):
                for shape in slide.shapes:
                    if hasattr(shape, "table"):
                        tbl = shape.table
                        rows = [[tbl.cell(r,c).text for c in range(len(tbl.columns))] for r in range(len(tbl.rows))]
                        df = pd.DataFrame(rows[1:], columns=rows[0]) if len(rows)>1 else pd.DataFrame(rows)
                        tables[f"PPT Table {i+1}"] = smart_clean_dataframe(df)
        elif name.endswith("pdf"):
            pdf = PdfReader(u)
            pages = [pg.extract_text() for pg in pdf.pages if pg.extract_text()]
            tables["PDF Text"] = pd.DataFrame({"Text": pages})
    except Exception as e:
        st.error(f"Upload failed: {e}")

# ---------- Tabs ----------
tab_data, tab_viz, tab_insight, tab_export, tab_settings = st.tabs(
    ["📄 Data", "📈 Visualize", "🧠 Insights", "📤 Export", "⚙️ Settings"]
)

# ---------- Data tab ----------
with tab_data:
    if not tables:
        st.info("Upload a file to begin.")
    else:
        key = st.selectbox("Choose table/sheet", list(tables.keys()))
        df = tables[key].copy()
        st.success(f"Loaded {key} → {df.shape[0]} rows, {df.shape[1]} columns")
        st.dataframe(df.head(10))

# ---------- Visualization ----------
with tab_viz:
    if not tables:
        st.info("Upload data to visualize.")
    else:
        key = st.selectbox("Table for visualization", list(tables.keys()))
        df = tables[key].copy()
        date_cols, num_cols, txt_cols = detect_types(df)
        x_col = st.selectbox("X Axis", df.columns)
        y_cols = st.multiselect("Y Axis", num_cols, default=num_cols[:1] if num_cols else [])

        # Date range safe handling
        if x_col in date_cols:
            df[x_col] = pd.to_datetime(df[x_col], errors="coerce")
            df = df.dropna(subset=[x_col])
            if not df.empty:
                min_d, max_d = df[x_col].min(), df[x_col].max()
                if pd.isna(min_d) or pd.isna(max_d):
                    today = pd.Timestamp.today(); min_d, max_d = today - pd.Timedelta(days=30), today
                if min_d == max_d: max_d = min_d + pd.Timedelta(days=1)
                try:
                    rng = st.date_input("Date range",
                        [min_d.date(), max_d.date()],
                        min_value=min_d.date(), max_value=max_d.date())
                except Exception:
                    rng = [min_d.date(), max_d.date()]
                df = df[(df[x_col] >= pd.to_datetime(rng[0])) & (df[x_col] <= pd.to_datetime(rng[1]))]

        chart_type = st.selectbox("Chart Type", ["Line","Area","Bar","Scatter","Pie"])
        if y_cols:
            try:
                if chart_type == "Line":    fig = px.line(df, x=x_col, y=y_cols, markers=True)
                elif chart_type == "Area":  fig = px.area(df, x=x_col, y=y_cols)
                elif chart_type == "Bar":   fig = px.bar(df, x=x_col, y=y_cols)
                elif chart_type == "Scatter": fig = px.scatter(df, x=x_col, y=y_cols)
                elif chart_type == "Pie" and len(y_cols)==1: fig = px.pie(df, names=x_col, values=y_cols[0])
                else: st.stop()
                fig.update_layout(template="plotly_white", height=600)
                st.plotly_chart(fig, use_container_width=True)
                st.session_state["last_fig"], st.session_state["last_df"] = fig, df
            except Exception as e:
                st.error(f"Chart failed: {e}")

# ---------- Insights ----------
with tab_insight:
    if not tables:
        st.info("Upload data to analyze.")
    else:
        key = st.selectbox("Choose table for insights", list(tables.keys()))
        df = tables[key].copy()
        _, num_cols, _ = detect_types(df)
        if num_cols:
            st.subheader("Summary Statistics")
            st.dataframe(df[num_cols].describe().T)
            st.subheader("Insights")
            insights = []
            for c in num_cols[:6]:
                s = pd.to_numeric(df[c], errors="coerce").dropna()
                if len(s)>1:
                    trend = "rising 📈" if s.iloc[-1]>s.iloc[0] else "falling 📉"
                    insights.append(f"**{c}** avg {s.mean():,.2f}, range {s.min():,.2f}–{s.max():,.2f}, {trend}")
            st.markdown("<br>".join(insights) or "No numeric insights available.", unsafe_allow_html=True)

# ---------- Export ----------
with tab_export:
    if not tables:
        st.info("Upload data to export.")
    else:
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
            st.download_button("⬇️ Excel", buf.getvalue(),
                               "cleaned_data.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               disabled=disabled)
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
                st.download_button("⬇️ PPTX", pptx_bytes,
                                   "chart_slide.pptx",
                                   "application/vnd.openxmlformats-officedocument.presentationml.presentation")
        else:
            st.info("Create a chart first in the Visualize tab.")

# ---------- Settings ----------
with tab_settings:
    st.checkbox("Disable image exports (Option 4)",
                value=st.session_state["disable_image_exports"],
                key="disable_image_exports",
                help="Turn on if your environment lacks Kaleido/Chrome — data exports still work.")
    st.markdown("""
**Notes**
- Option 3 (cloud-safe export) prevents Kaleido crashes.  
- Option 4 disables image exports explicitly.  
- Data exports (CSV / Excel / JSON) always work.
""")
