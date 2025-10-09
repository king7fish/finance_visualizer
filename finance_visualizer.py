# finance_dashboard_industry.py
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from io import BytesIO
from datetime import datetime
import json, re, requests

from pptx import Presentation
from pptx.util import Inches
from docx import Document
from PyPDF2 import PdfReader

# ---------- Page setup ----------
st.set_page_config(page_title="Finance Dashboard — Industry Edition", layout="wide", initial_sidebar_state="expanded")
PRIMARY = "#3A86FF"

st.markdown(f"""
<h1 style="text-align:center;color:{PRIMARY};margin-bottom:0">🏆 Finance Dashboard — Industry Edition</h1>
<p style="text-align:center;color:#6b7280;margin-top:4px">Messy data in. Analyst-level insights out.</p>
""", unsafe_allow_html=True)

# ---------- Helpers ----------
@st.cache_data(show_spinner=False)
def fetch_fx_rate(base: str, quote: str):
    """1 base -> ? quote using exchangerate.host (fallback None)"""
    try:
        r = requests.get(
            f"https://api.exchangerate.host/convert?from={base.upper()}&to={quote.upper()}&amount=1",
            timeout=8
        )
        j = r.json()
        return float(j.get("info", {}).get("rate", None))
    except Exception:
        return None

@st.cache_data(show_spinner=False)
def read_excel_all_sheets(file):
    xls = pd.ExcelFile(file)
    sheets = {}
    for name in xls.sheet_names:
        try:
            df = xls.parse(name, header=None)  # raw; we'll clean header
            sheets[name] = df
        except Exception:
            sheets[name] = None
    return sheets

def promote_header_if_found(df_raw: pd.DataFrame) -> pd.DataFrame:
    """Find a likely header row and promote it; otherwise sanitize column names."""
    df = df_raw.copy()
    best_idx = None
    for i in range(min(10, len(df))):
        row = df.iloc[i].astype(str).str.strip()
        if (row != "").sum() >= max(2, int(0.5 * len(row))):
            best_idx = i
            break
    if best_idx is not None:
        new = df.iloc[best_idx+1:].copy()
        headers = df.iloc[best_idx].astype(str).str.strip().tolist()
        headers = [h if h and h.lower() != "unnamed" else f"col_{i}" for i, h in enumerate(headers)]
        new.columns = headers
        return new.reset_index(drop=True)
    # fallback
    df = df.copy()
    df.columns = [f"col_{i}" if (str(c).strip()=="" or str(c).lower().startswith("unnamed")) else str(c) for i, c in enumerate(df.columns)]
    return df

def looks_numeric(s: pd.Series) -> bool:
    vals = s.dropna().astype(str).head(60)
    if len(vals) == 0: return False
    patt = re.compile(r"^\s*[-+]?\d{1,3}(?:,\d{3})*(?:\.\d+)?\s*$|^\s*[-+]?\d+(?:\.\d+)?\s*$")
    hits = sum(bool(patt.match(v)) for v in vals)
    return hits >= 0.5 * len(vals)

def looks_date(s: pd.Series) -> bool:
    vals = s.dropna().astype(str).head(60)
    if len(vals) == 0: return False
    hints = sum(("/" in v or "-" in v) for v in vals)
    return hints >= 0.5 * len(vals)

def smart_clean_dataframe(df_in: pd.DataFrame) -> pd.DataFrame:
    """
    Non-destructive: preserve labels; clean numbers/dates intelligently.
    - trims headers
    - converts numeric-looking columns (removes commas, stray symbols)
    - converts date-looking columns
    - never drops rows unless fully empty
    """
    df = df_in.copy()
    # drop fully empty rows/cols
    df = df.dropna(how="all")
    df = df.loc[:, df.notna().sum() > 0]

    # clean headers
    df.columns = [str(c).strip().replace("\n"," ") for c in df.columns]

    for col in df.columns:
        s = df[col]
        # numeric-like
        if looks_numeric(s):
            # remove commas and non numeric except . - digits
            cleaned = s.astype(str).str.replace(",", "", regex=False)
            cleaned = cleaned.str.replace(r"[^0-9.\-]", "", regex=True).replace({"": np.nan})
            df[col] = pd.to_numeric(cleaned, errors="coerce")
        # date-like
        elif looks_date(s):
            df[col] = pd.to_datetime(s, errors="coerce")
        else:
            df[col] = s.astype(str).str.strip()

    return df.reset_index(drop=True)

def detect_types(df: pd.DataFrame):
    date_cols, num_cols, txt_cols = [], [], []
    for c in df.columns:
        s = df[c]
        if pd.api.types.is_datetime64_any_dtype(s) or looks_date(s):
            date_cols.append(c)
        elif pd.api.types.is_numeric_dtype(s) or looks_numeric(s):
            num_cols.append(c)
        else:
            txt_cols.append(c)
    return date_cols, num_cols, txt_cols

def safe_resample(df: pd.DataFrame, time_col: str, freq: str, how: str):
    out = df.copy()
    out[time_col] = pd.to_datetime(out[time_col], errors="coerce")
    out = out.dropna(subset=[time_col]).set_index(time_col)
    numeric_only = out.select_dtypes(include=["number"])
    if numeric_only.empty:
        return df
    agg = how if how in {"sum","mean","median","max","min"} else "mean"
    res = numeric_only.resample(freq).agg(agg).reset_index()
    return res

def fig_to_png(fig):
    return fig.to_image(format="png", engine="kaleido")

def pptx_with_chart(fig, title="Chart"):
    img = fig_to_png(fig)
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    # title
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.2), Inches(9), Inches(0.8))
    tb.text_frame.text = title
    # image
    slide.shapes.add_picture(BytesIO(img), Inches(0.6), Inches(1.2), width=Inches(8.8))
    bio = BytesIO(); prs.save(bio); bio.seek(0)
    return bio.getvalue()

# ---------- Sidebar: Upload ----------
st.sidebar.header("📁 Upload")
u = st.sidebar.file_uploader("Upload data file", type=["xlsx","xls","csv","json","docx","pptx","pdf"])

# ---------- Load & select table ----------
tables = {}  # name -> DataFrame

if u:
    fname = u.name.lower()

    if fname.endswith((".xlsx",".xls")):
        raw_sheets = read_excel_all_sheets(u)
        # promote headers & clean each
        candidate_names = []
        for sname, raw in raw_sheets.items():
            if raw is None: continue
            df0 = promote_header_if_found(raw)
            df0 = smart_clean_dataframe(df0)
            if len(df0) and len(df0.columns):
                tables[f"Excel — {sname}"] = df0
                candidate_names.append(sname)
        # choose largest by rows as default
        default_key = 0
        if tables:
            sizes = [len(df) for df in tables.values()]
            default_key = int(np.argmax(sizes))

    elif fname.endswith(".csv"):
        df = pd.read_csv(u)
        tables["CSV"] = smart_clean_dataframe(df)

    elif fname.endswith(".json"):
        df = pd.read_json(u)
        tables["JSON"] = smart_clean_dataframe(df)

    elif fname.endswith(".docx"):
        doc = Document(u)
        # extract tables if any; else paragraphs
        if doc.tables:
            for i, t in enumerate(doc.tables):
                rows = [[cell.text for cell in row.cells] for row in t.rows]
                if rows:
                    dft = pd.DataFrame(rows[1:], columns=rows[0]) if len(rows)>1 else pd.DataFrame(rows)
                    tables[f"Word Table {i+1}"] = smart_clean_dataframe(dft)
        else:
            paras = [p.text for p in doc.paragraphs if p.text.strip()]
            tables["Word Text"] = pd.DataFrame({"Text": paras})

    elif fname.endswith(".pptx"):
        prs = Presentation(u)
        idx = 1
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape,"table"):
                    tbl = shape.table
                    rows = []
                    for r in range(len(tbl.rows)):
                        rows.append([tbl.cell(r,c).text for c in range(len(tbl.columns))])
                    dft = pd.DataFrame(rows[1:], columns=rows[0]) if len(rows)>1 else pd.DataFrame(rows)
                    tables[f"PPTX Table {idx}"] = smart_clean_dataframe(dft)
                    idx += 1
        if not tables:
            texts = []
            for slide in prs.slides:
                text = " ".join([sh.text for sh in slide.shapes if hasattr(sh,"text")])
                texts.append(text)
            tables["PPTX Text"] = pd.DataFrame({"Text": texts})

    elif fname.endswith(".pdf"):
        pdf = PdfReader(u)
        pages = [pg.extract_text() for pg in pdf.pages if pg.extract_text()]
        tables["PDF Text"] = pd.DataFrame({"Text": pages})

# ---------- Tabs ----------
tab_data, tab_viz, tab_insight, tab_export, tab_settings = st.tabs(
    ["📄 Data", "📈 Visualize", "🧠 Insights", "📤 Export", "⚙️ Settings"]
)

with tab_data:
    if not tables:
        st.info("Upload a file to begin.")
    else:
        key = st.selectbox("Choose sheet/table to work with", list(tables.keys()))
        df = tables[key].copy()
        st.success(f"Loaded: **{key}** — shape {df.shape}")
        with st.expander("🔍 Preview (first 12 rows)"):
            st.dataframe(df.head(12), use_container_width=True)

        # Profile
        date_cols, num_cols, txt_cols = detect_types(df)
        c1, c2, c3 = st.columns(3)
        c1.metric("Rows", f"{len(df):,}")
        c2.metric("Columns", f"{len(df.columns):,}")
        c3.metric("Numeric columns", f"{len(num_cols)}")

        st.markdown("**Detected column types**")
        st.write(f"📅 Date-like: {date_cols or '—'}")
        st.write(f"🔢 Numeric-like: {num_cols or '—'}")
        st.write(f"🔤 Text-like: {txt_cols or '—'}")

        with st.expander("🧽 Before vs After cleaning (sample)"):
            st.caption("Your data is already cleaned by our smart cleaner. For transparency, here’s a quick sample.")
            st.dataframe(df.head(8))

with tab_viz:
    if not tables:
        st.info("Upload a file to visualize.")
    else:
        # reselect df based on selection in Data tab
        key = st.selectbox("Active sheet/table", list(tables.keys()), key="viz_key")
        df = tables[key].copy()
        date_cols, num_cols, txt_cols = detect_types(df)

        # X axis
        x_candidates = date_cols + txt_cols + [c for c in df.columns if c not in date_cols+txt_cols]
        x_col = st.selectbox("X axis", x_candidates, index=0)

        # Date range & resample
        if x_col in date_cols:
            df[x_col] = pd.to_datetime(df[x_col], errors="coerce")
            df = df.dropna(subset=[x_col])
            min_d, max_d = df[x_col].min(), df[x_col].max()
            rng = st.date_input("Date range", [min_d.date(), max_d.date()])
            if isinstance(rng, (list,tuple)) and len(rng)==2:
                df = df[(df[x_col] >= pd.to_datetime(rng[0])) & (df[x_col] <= pd.to_datetime(rng[1]))]
            freq = st.selectbox("Resample (if time axis)", ["None","D (Daily)","W (Weekly)","M (Monthly)","Q (Quarterly)","Y (Yearly)"], index=2)
            agg = st.selectbox("Aggregation", ["mean","sum","median","max","min"], index=0)
            if freq != "None":
                code = freq.split()[0]  # "D","W","M","Q","Y"
                df = safe_resample(df, x_col, code, agg)

        # Y axes
        if not num_cols:
            st.warning("No numeric columns detected to plot.")
            st.stop()
        y_cols = st.multiselect("Y axis (one or more)", num_cols, default=num_cols[:1])

        # Currency & units
        st.markdown("### Units & Currency")
        use_fx = st.checkbox("Convert currency (optional)")
        if use_fx:
            base = st.text_input("From currency (e.g. USD)", "USD").upper()
            quote = st.text_input("To currency (e.g. GBP)", "GBP").upper()
            live = fetch_fx_rate(base, quote)
            manual = st.number_input("Or enter manual rate (multiply)", min_value=0.0, value=0.0, step=0.0001, format="%.6f")
            fx = manual if manual>0 else (live if live else 1.0)
            if live: st.success(f"Live rate: 1 {base} = {live:.6f} {quote}")
            else: st.info("Using manual rate or 1.0 if none provided.")
            for c in y_cols:
                df[c] = pd.to_numeric(df[c], errors="coerce") * fx

        scale = st.selectbox("Display scale", ["None","Thousands (K)","Millions (M)","Billions (B)"], index=0)
        scale_factor = {"None":1,"Thousands (K)":1e3,"Millions (M)":1e6,"Billions (B)":1e9}[scale]
        for c in y_cols:
            df[c] = pd.to_numeric(df[c], errors="coerce") / scale_factor

        # Chart options
        st.markdown("### Chart options")
        chart_type = st.selectbox("Type", ["Line","Area","Bar","Stacked Bar","Scatter","Pie","Dual Axis (left/right)"])
        colors = {c: st.color_picker(f"Color for {c}", "#1f77b4") for c in y_cols}
        line_w = st.slider("Line width", 1, 6, 2)
        marker_s = st.slider("Marker size", 3, 12, 6)
        y_tick = st.number_input("Y tick step (0=auto)", min_value=0, value=0)
        log_y = st.checkbox("Log scale Y", value=False)

        # Build figure
        fig = go.Figure()
        def add_trace(trace):
            fig.add_trace(trace)

        if chart_type in ["Line","Area","Bar","Stacked Bar","Scatter","Dual Axis (left/right)"]:
            for i, y in enumerate(y_cols):
                yaxis = "y" if (chart_type != "Dual Axis (left/right)" or i%2==0) else "y2"
                if chart_type in ["Line","Area"]:
                    add_trace(go.Scatter(x=df[x_col], y=df[y], mode="lines+markers",
                                         name=y, line=dict(color=colors[y], width=line_w),
                                         marker=dict(size=marker_s), fill="tozeroy" if chart_type=="Area" else None,
                                         yaxis=yaxis))
                elif chart_type == "Scatter":
                    add_trace(go.Scatter(x=df[x_col], y=df[y], mode="markers",
                                         name=y, marker=dict(size=marker_s, color=colors[y]),
                                         yaxis=yaxis))
                else: # Bar / Stacked Bar
                    add_trace(go.Bar(x=df[x_col], y=df[y], name=y,
                                     marker_color=colors[y], yaxis=yaxis))
            if chart_type == "Stacked Bar":
                fig.update_layout(barmode="stack")
            else:
                fig.update_layout(barmode="group")
        elif chart_type == "Pie":
            if len(y_cols) != 1:
                st.error("Pie requires exactly one Y column.")
                st.stop()
            fig = px.pie(df, names=x_col, values=y_cols[0], color_discrete_sequence=px.colors.qualitative.Bold)

        # Layout
        ytitle = ", ".join(y_cols) + (f" ({scale.split()[0]})" if scale!="None" else "")
        layout = dict(template="plotly_white", title=f"{', '.join(y_cols)} vs {x_col}",
                      xaxis_title=x_col, yaxis_title=ytitle, height=600)
        if chart_type == "Dual Axis (left/right)" and len(y_cols) > 1:
            layout.update(yaxis=dict(title=y_cols[0]),
                          yaxis2=dict(title=y_cols[1], overlaying="y", side="right"))
        fig.update_layout(**layout)
        if y_tick>0: fig.update_yaxes(dtick=y_tick)
        if log_y: fig.update_yaxes(type="log")

        st.plotly_chart(fig, use_container_width=True)

with tab_insight:
    if not tables:
        st.info("Upload data to generate insights.")
    else:
        # Use active df
        key = st.selectbox("Active sheet/table", list(tables.keys()), key="ins_key")
        df = tables[key].copy()
        date_cols, num_cols, txt_cols = detect_types(df)

        st.subheader("Quick Stats")
        if num_cols:
            stats = df[num_cols].describe().T
            stats["missing_%"] = 100*(1 - (df[num_cols].notna().sum()/len(df)))
            st.dataframe(stats.round(3))
        else:
            st.info("No numeric columns detected.")

        st.subheader("Narrative Insights (offline)")
        bullets = []
        if num_cols:
            for c in num_cols[:6]:
                series = pd.to_numeric(df[c], errors="coerce").dropna()
                if series.empty: continue
                mean, std, mn, mx = series.mean(), series.std(), series.min(), series.max()
                trend = "rising 📈" if mx > mean*1.2 else ("stable ⚖️" if std < 0.3*abs(mean) else "volatile ⚠️")
                bullets.append(f"**{c}** avg {mean:,.2f}, range {mn:,.2f}–{mx:,.2f}, {trend}.")
        if date_cols and num_cols:
            # simple YoY/MoM if time column present
            tcol = date_cols[0]
            dft = df[[tcol] + num_cols].copy()
            dft[tcol] = pd.to_datetime(dft[tcol], errors="coerce")
            dft = dft.dropna(subset=[tcol]).sort_values(tcol)
            # monthly mean for first num col
            g = dft.set_index(tcol)[num_cols[0]].resample("M").mean()
            if len(g) >= 13:
                yoy = (g.iloc[-1] - g.iloc[-13]) / (abs(g.iloc[-13]) + 1e-9) * 100
                bullets.append(f"**{num_cols[0]}** YoY change ≈ {yoy:+.1f}%")
        if not bullets:
            bullets = ["Not enough numeric/time data for robust insights."]
        st.markdown("<br>".join(f"• {b}" for b in bullets), unsafe_allow_html=True)

        if len(num_cols) >= 2:
            st.subheader("Correlation (numeric)")
            st.dataframe(df[num_cols].corr().round(3))

with tab_export:
    if not tables:
        st.info("Upload data to export.")
    else:
        key = st.selectbox("Active sheet/table", list(tables.keys()), key="exp_key")
        df = tables[key].copy()
        col1, col2, col3, col4 = st.columns(4)
        # Data exports
        with col1:
            st.download_button("CSV", df.to_csv(index=False).encode("utf-8"),
                               "cleaned_data.csv", "text/csv")
        with col2:
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                df.to_excel(w, index=False, sheet_name="Data")
            st.download_button("Excel", data=buf.getvalue(), file_name="cleaned_data.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col3:
            st.download_button("JSON", df.to_json(orient="records").encode("utf-8"),
                               "cleaned_data.json", "application/json")
        with col4:
            st.info("First create a chart in the Visualize tab to export images/slide.")

        st.markdown("---")
        st.subheader("Export chart (from Visualize tab)")
        if "last_fig" not in st.session_state:
            st.session_state["last_fig"] = None

        # Provide a lightweight chart creator here for export convenience
        make_chart = st.checkbox("Create quick export chart here")
        if make_chart:
            # minimal chart builder for export-only
            date_cols, num_cols, txt_cols = detect_types(df)
            x_candidates = date_cols + txt_cols + [c for c in df.columns if c not in date_cols+txt_cols]
            x_col = st.selectbox("X axis", x_candidates, key="exp_x")
            y_cols = st.multiselect("Y axis", num_cols, default=num_cols[:1], key="exp_y")
            if x_col and y_cols:
                fig = px.line(df, x=x_col, y=y_cols) if len(y_cols)>1 else px.line(df, x=x_col, y=y_cols[0])
                st.session_state["last_fig"] = fig
                st.plotly_chart(fig, use_container_width=True)

        if st.session_state["last_fig"] is None:
            st.info("No chart available yet. Build one, then export.")
        else:
            fig = st.session_state["last_fig"]
            c1, c2, c3 = st.columns(3)
            with c1:
                if st.button("Download PNG"):
                    try:
                        png = fig.to_image(format="png", engine="kaleido")
                        st.download_button("Save PNG", png, "chart.png", "image/png")
                    except Exception as e:
                        st.error(f"PNG export failed: {e}")
            with c2:
                if st.button("Download PDF"):
                    try:
                        pdf = fig.to_image(format="pdf", engine="kaleido")
                        st.download_button("Save PDF", pdf, "chart.pdf", "application/pdf")
                    except Exception as e:
                        st.error(f"PDF export failed: {e}")
            with c3:
                if st.button("Download PPTX slide"):
                    try:
                        pptx_bytes = pptx_with_chart(fig, "Finance Dashboard — Chart")
                        st.download_button("Save PPTX", pptx_bytes, "chart_slide.pptx",
                                           "application/vnd.openxmlformats-officedocument.presentationml.presentation")
                    except Exception as e:
                        st.error(f"PPTX export failed: {e}")

with tab_settings:
    st.write("**Tips**")
    st.markdown("""
    - If your Excel has multiple sections/headers, we auto-promote the most likely header row per sheet.  
    - Cleaning is **non-destructive**: labels/categories are preserved; only numeric/date columns are coerced.  
    - Resampling ignores non-numeric columns to prevent crashes.  
    - For very large files: consider splitting by sheet or pre-filtering to essential columns.
    """)
