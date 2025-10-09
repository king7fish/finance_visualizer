# finance_dashboard_pro.py
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
import requests
import re

# ---------- APP SETUP ----------
st.set_page_config(page_title="Finance Dashboard Pro", layout="wide")
PRIMARY_COLOR = "#3A86FF"

st.markdown(f"""
<h1 style="text-align:center; color:{PRIMARY_COLOR}; margin-bottom:0">📊 Finance Dashboard Pro</h1>
<p style="text-align:center; color:#6c757d; margin-top:4px;">Smart, Fast, and Reliable — Built for Real-World Data</p>
""", unsafe_allow_html=True)

# ---------- UTILITIES ----------
def looks_numeric(series):
    vals = series.dropna().astype(str).head(60)
    patt = re.compile(r"^-?\d+(\.\d+)?$")
    return sum(bool(patt.match(v.replace(",", ""))) for v in vals) > 0.5 * len(vals)

def looks_date(series):
    vals = series.dropna().astype(str).head(60)
    return sum(("/" in v or "-" in v) for v in vals) > 0.5 * len(vals)

def smart_clean(df):
    df = df.dropna(how="all").copy()
    df.columns = [str(c).strip().replace("\n", " ") for c in df.columns]
    for col in df.columns:
        s = df[col]
        if looks_numeric(s):
            cleaned = s.astype(str).replace(",", "", regex=False)
            cleaned = cleaned.str.replace(r"[^0-9.\-]", "", regex=True).replace("", np.nan)
            df[col] = pd.to_numeric(cleaned, errors="coerce")
        elif looks_date(s):
            df[col] = pd.to_datetime(s, errors="coerce")
        else:
            df[col] = s.astype(str).str.strip()
    return df.reset_index(drop=True)

def pptx_with_chart(fig, title="Chart"):
    img = fig.to_image(format="png", engine="kaleido")
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(9), Inches(1)).text = title
    slide.shapes.add_picture(BytesIO(img), Inches(1), Inches(1.3), width=Inches(8))
    out = BytesIO(); prs.save(out); out.seek(0)
    return out.getvalue()

def fetch_fx(base, quote):
    try:
        r = requests.get(f"https://api.exchangerate.host/convert?from={base}&to={quote}&amount=1", timeout=6)
        return float(r.json()["info"]["rate"])
    except Exception:
        return None

def safe_resample(df, time_col, freq, how):
    df = df.copy()
    df[time_col] = pd.to_datetime(df[time_col], errors="coerce")
    df = df.dropna(subset=[time_col])
    if df.empty: return df
    numeric = df.select_dtypes(include=["number"])
    agg = how if how in ["sum", "mean", "max", "min", "median"] else "mean"
    res = numeric.resample(freq, on=time_col).agg(agg).reset_index()
    return res

# ---------- FILE UPLOAD ----------
st.sidebar.header("📁 Upload Your File")
file = st.sidebar.file_uploader("Upload Excel, CSV, JSON, DOCX, PPTX or PDF", 
                                type=["xlsx","xls","csv","json","docx","pptx","pdf"])

tables = {}
if file:
    name = file.name.lower()
    if name.endswith(("xlsx","xls")):
        xls = pd.ExcelFile(file)
        for sheet in xls.sheet_names:
            raw = xls.parse(sheet)
            tables[f"Excel — {sheet}"] = smart_clean(raw)
    elif name.endswith("csv"):
        tables["CSV"] = smart_clean(pd.read_csv(file))
    elif name.endswith("json"):
        tables["JSON"] = smart_clean(pd.read_json(file))
    elif name.endswith("docx"):
        doc = Document(file)
        for i, t in enumerate(doc.tables):
            rows = [[cell.text for cell in row.cells] for row in t.rows]
            df = pd.DataFrame(rows[1:], columns=rows[0]) if len(rows) > 1 else pd.DataFrame(rows)
            tables[f"Word Table {i+1}"] = smart_clean(df)
    elif name.endswith("pptx"):
        prs = Presentation(file)
        for i, slide in enumerate(prs.slides):
            for shape in slide.shapes:
                if hasattr(shape, "table"):
                    tbl = shape.table
                    rows = [[tbl.cell(r,c).text for c in range(len(tbl.columns))] for r in range(len(tbl.rows))]
                    df = pd.DataFrame(rows[1:], columns=rows[0]) if len(rows)>1 else pd.DataFrame(rows)
                    tables[f"PPT Table {i+1}"] = smart_clean(df)
    elif name.endswith("pdf"):
        pdf = PdfReader(file)
        pages = [pg.extract_text() for pg in pdf.pages if pg.extract_text()]
        tables["PDF Text"] = pd.DataFrame({"Text": pages})

# ---------- TABS ----------
tab_data, tab_viz, tab_insight, tab_export = st.tabs(["📄 Data", "📈 Visualize", "🧠 Insights", "📤 Export"])

# ---------- DATA TAB ----------
with tab_data:
    if not tables:
        st.info("Upload a file to begin.")
    else:
        key = st.selectbox("Choose sheet/table", list(tables.keys()))
        df = tables[key].copy()
        st.success(f"Loaded: {key} — {df.shape[0]} rows, {df.shape[1]} columns")
        with st.expander("🔍 Preview Data"):
            st.dataframe(df.head(12))
        with st.expander("🧠 Data Types"):
            st.write(df.dtypes)

# ---------- VISUALIZE TAB ----------
with tab_viz:
    if not tables:
        st.info("Upload data first.")
    else:
        key = st.selectbox("Active table", list(tables.keys()), key="viz_key")
        df = tables[key].copy()

        # Identify columns
        date_cols = [c for c in df.columns if pd.api.types.is_datetime64_any_dtype(df[c])]
        num_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]

        x_col = st.selectbox("X Axis", df.columns)
        y_cols = st.multiselect("Y Axis", num_cols, default=num_cols[:1] if num_cols else None)

        if x_col in date_cols:
            df[x_col] = pd.to_datetime(df[x_col], errors="coerce")
            df = df.dropna(subset=[x_col])
            if not df.empty:
                min_d, max_d = df[x_col].min(), df[x_col].max()
                if pd.isna(min_d) or pd.isna(max_d):
                    today = pd.Timestamp.today()
                    min_d, max_d = today - pd.Timedelta(days=30), today
                if min_d == max_d:
                    max_d = min_d + pd.Timedelta(days=1)
                try:
                    rng = st.date_input(
                        "Date range",
                        [min_d.date(), max_d.date()],
                        min_value=min_d.date(),
                        max_value=max_d.date(),
                    )
                except Exception:
                    st.warning("Invalid dates found — showing full range automatically.")
                    rng = [min_d.date(), max_d.date()]
                if isinstance(rng, (list, tuple)) and len(rng) == 2:
                    df = df[(df[x_col] >= pd.to_datetime(rng[0])) & (df[x_col] <= pd.to_datetime(rng[1]))]

                freq = st.selectbox("Resample Frequency", ["None","D","W","M","Q","Y"], index=2)
                agg = st.selectbox("Aggregation", ["mean","sum","max","min","median"], index=0)
                if freq != "None":
                    df = safe_resample(df, x_col, freq, agg)

        chart_type = st.selectbox("Chart Type", ["Line","Area","Bar","Scatter","Pie"])
        if not y_cols:
            st.warning("Select at least one numeric column to plot.")
            st.stop()

        if chart_type == "Line":
            fig = px.line(df, x=x_col, y=y_cols, markers=True)
        elif chart_type == "Area":
            fig = px.area(df, x=x_col, y=y_cols)
        elif chart_type == "Bar":
            fig = px.bar(df, x=x_col, y=y_cols)
        elif chart_type == "Scatter":
            fig = px.scatter(df, x=x_col, y=y_cols)
        elif chart_type == "Pie":
            if len(y_cols) != 1:
                st.error("Pie chart requires one Y-axis.")
                st.stop()
            fig = px.pie(df, names=x_col, values=y_cols[0])
        else:
            st.error("Unsupported chart type.")
            st.stop()

        fig.update_layout(height=600, template="plotly_white")
        st.plotly_chart(fig, use_container_width=True)
        st.session_state["fig"] = fig

# ---------- INSIGHTS TAB ----------
with tab_insight:
    if not tables:
        st.info("Upload data to analyze.")
    else:
        key = st.selectbox("Choose table", list(tables.keys()), key="ins_key")
        df = tables[key].copy()
        num_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
        if not num_cols:
            st.warning("No numeric data found.")
        else:
            st.subheader("Summary Statistics")
            st.dataframe(df[num_cols].describe().T)
            st.subheader("AI-style Insights")
            insights = []
            for c in num_cols:
                s = pd.to_numeric(df[c], errors="coerce").dropna()
                if not s.empty:
                    trend = "rising 📈" if s.iloc[-1] > s.iloc[0] else "declining 📉"
                    insights.append(f"**{c}** — avg: {s.mean():,.2f}, range: {s.min():,.2f}–{s.max():,.2f}, {trend}.")
            st.markdown("<br>".join(insights), unsafe_allow_html=True)
            if len(num_cols) > 1:
                st.subheader("Correlation Matrix")
                st.dataframe(df[num_cols].corr().round(3))

# ---------- EXPORT TAB ----------
with tab_export:
    if not tables:
        st.info("Upload data to export.")
    else:
        key = st.selectbox("Export table", list(tables.keys()), key="exp_key")
        df = tables[key].copy()

        col1, col2, col3 = st.columns(3)
        with col1:
            st.download_button("⬇️ Download CSV", df.to_csv(index=False).encode('utf-8'), "cleaned_data.csv", "text/csv")
        with col2:
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as w:
                df.to_excel(w, index=False, sheet_name="Data")
            st.download_button("⬇️ Download Excel", buffer.getvalue(), "cleaned_data.xlsx")
        with col3:
            st.download_button("⬇️ Download JSON", df.to_json(orient="records").encode('utf-8'), "cleaned_data.json", "application/json")

        if "fig" in st.session_state and st.session_state["fig"]:
            fig = st.session_state["fig"]
            st.download_button("⬇️ Download PPTX Chart", pptx_with_chart(fig), "chart_slide.pptx",
                               "application/vnd.openxmlformats-officedocument.presentationml.presentation")
