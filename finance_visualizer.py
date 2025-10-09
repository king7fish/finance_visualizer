# finance_dashboard_pro.py
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import re, requests
from io import BytesIO
from datetime import datetime

from pptx import Presentation
from pptx.util import Inches
from docx import Document
from PyPDF2 import PdfReader

# -------------------- App setup --------------------
st.set_page_config(page_title="Finance Dashboard Pro", layout="wide", initial_sidebar_state="expanded")
PRIMARY = "#3A86FF"

st.markdown(
    f"<h1 style='text-align:center;color:{PRIMARY};margin-bottom:0'>🏆 Finance Dashboard Pro</h1>"
    "<p style='text-align:center;color:#6b7280;margin-top:4px'>Messy data in. Analyst-level insights out.</p>",
    unsafe_allow_html=True,
)

# -------------------- Settings (global state) --------------------
if "disable_image_exports" not in st.session_state:
    st.session_state["disable_image_exports"] = False

# -------------------- Helpers --------------------
def looks_numeric(series: pd.Series) -> bool:
    vals = series.dropna().astype(str).head(60)
    patt = re.compile(r"^\s*[-+]?\d{1,3}(?:,\d{3})*(?:\.\d+)?\s*$|^\s*[-+]?\d+(?:\.\d+)?\s*$")
    return sum(bool(patt.match(v)) for v in vals) >= 0.5 * len(vals)

def looks_date(series: pd.Series) -> bool:
    vals = series.dropna().astype(str).head(60)
    return sum(("/" in v or "-" in v) for v in vals) >= 0.5 * len(vals)

def promote_header_if_found(df_raw: pd.DataFrame) -> pd.DataFrame:
    df = df_raw.copy()
    candidate = None
    for i in range(min(10, len(df))):
        row = df.iloc[i].astype(str).str.strip()
        if (row != "").sum() >= max(2, int(0.5 * len(row))):
            candidate = i
            break
    if candidate is not None:
        new = df.iloc[candidate + 1 :].copy()
        headers = df.iloc[candidate].astype(str).str.strip().tolist()
        headers = [h if h and h.lower() != "unnamed" else f"col_{i}" for i, h in enumerate(headers)]
        new.columns = headers
        return new.reset_index(drop=True)
    df.columns = [f"col_{i}" if (str(c).strip() == "" or str(c).lower().startswith("unnamed")) else str(c) for i, c in enumerate(df.columns)]
    return df

def smart_clean_dataframe(df_in: pd.DataFrame) -> pd.DataFrame:
    """Preserve labels; only coerce numeric/date columns; never mass-drop real data."""
    df = df_in.copy()
    df = df.dropna(how="all")
    df = df.loc[:, df.notna().sum() > 0]
    df.columns = [str(c).strip().replace("\n", " ") for c in df.columns]
    for col in df.columns:
        s = df[col]
        if looks_numeric(s):
            cleaned = s.astype(str).str.replace(",", "", regex=False)
            cleaned = cleaned.str.replace(r"[^0-9.\-]", "", regex=True).replace("", np.nan)
            df[col] = pd.to_numeric(cleaned, errors="coerce")
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
    out = out.dropna(subset=[time_col])
    if out.empty:
        return df
    out = out.sort_values(time_col)
    numeric_only = out.select_dtypes(include=["number"])
    if numeric_only.empty:
        return out
    agg = how if how in {"sum","mean","median","max","min"} else "mean"
    res = numeric_only.resample(freq, on=time_col).agg(agg).reset_index()
    return res

def fetch_fx_rate(base: str, quote: str):
    try:
        r = requests.get(f"https://api.exchangerate.host/convert?from={base.upper()}&to={quote.upper()}&amount=1", timeout=8)
        j = r.json()
        return float(j.get("info", {}).get("rate", None))
    except Exception:
        return None

# ---- Export helpers (Option 3 + Option 4 supported) ----
def fig_to_png_safe(fig):
    """Try to convert figure to PNG with Kaleido. Return bytes or None if unsupported."""
    if st.session_state.get("disable_image_exports", False):
        return None
    try:
        return fig.to_image(format="png", engine="kaleido")
    except Exception:
        return None

def fig_to_pdf_safe(fig):
    if st.session_state.get("disable_image_exports", False):
        return None
    try:
        return fig.to_image(format="pdf", engine="kaleido")
    except Exception:
        return None

def pptx_with_chart_failsafe(fig, title="Chart"):
    """
    Option 3 (Cloud failsafe):
    - Try PNG export via Kaleido, embed in PPTX
    - If Kaleido unavailable, create a PPTX with the chart title and a note (no crash)
    """
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(9), Inches(1)).text = title

    img = fig_to_png_safe(fig)
    if img is not None:
        slide.shapes.add_picture(BytesIO(img), Inches(0.8), Inches(1.4), width=Inches(8.8))
    else:
        # Fallback: no image export available in this environment
        tb = slide.shapes.add_textbox(Inches(0.8), Inches(1.4), Inches(9), Inches(3))
        tf = tb.text_frame
        tf.text = "Chart image export is unavailable in this environment.\n\nData can be exported from the app."
    buf = BytesIO(); prs.save(buf); buf.seek(0)
    return buf.getvalue()

# -------------------- File upload --------------------
st.sidebar.header("📁 Upload")
u = st.sidebar.file_uploader("Upload Excel/CSV/JSON/DOCX/PPTX/PDF", type=["xlsx","xls","csv","json","docx","pptx","pdf"])

tables = {}
if u:
    fname = u.name.lower()
    if fname.endswith((".xlsx",".xls")):
        xls = pd.ExcelFile(u)
        for s in xls.sheet_names:
            try:
                raw = xls.parse(s, header=None)  # raw for header promotion
                df0 = promote_header_if_found(raw)
            except Exception:
                df0 = xls.parse(s)
            tables[f"Excel — {s}"] = smart_clean_dataframe(df0)
    elif fname.endswith(".csv"):
        tables["CSV"] = smart_clean_dataframe(pd.read_csv(u))
    elif fname.endswith(".json"):
        tables["JSON"] = smart_clean_dataframe(pd.read_json(u))
    elif fname.endswith(".docx"):
        doc = Document(u)
        if doc.tables:
            for i, t in enumerate(doc.tables):
                rows = [[cell.text for cell in row.cells] for row in t.rows]
                dft = pd.DataFrame(rows[1:], columns=rows[0]) if len(rows) > 1 else pd.DataFrame(rows)
                tables[f"Word Table {i+1}"] = smart_clean_dataframe(dft)
        else:
            paras = [p.text for p in doc.paragraphs if p.text.strip()]
            tables["Word Text"] = pd.DataFrame({"Text": paras})
    elif fname.endswith(".pptx"):
        prs = Presentation(u)
        idx = 1
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "table"):
                    tbl = shape.table
                    rows = [[tbl.cell(r,c).text for c in range(len(tbl.columns))] for r in range(len(tbl.rows))]
                    dft = pd.DataFrame(rows[1:], columns=rows[0]) if len(rows)>1 else pd.DataFrame(rows)
                    tables[f"PPTX Table {idx}"] = smart_clean_dataframe(dft); idx += 1
        if not tables:
            texts = [" ".join([sh.text for sh in slide.shapes if hasattr(sh,"text")]) for slide in prs.slides]
            tables["PPTX Text"] = pd.DataFrame({"Text": texts})
    elif fname.endswith(".pdf"):
        pdf = PdfReader(u)
        pages = [pg.extract_text() for pg in pdf.pages if pg.extract_text()]
        tables["PDF Text"] = pd.DataFrame({"Text": pages})

# -------------------- Tabs --------------------
tab_data, tab_viz, tab_insight, tab_export, tab_settings = st.tabs(
    ["📄 Data", "📈 Visualize", "🧠 Insights", "📤 Export", "⚙️ Settings"]
)

# -------------------- Data tab --------------------
with tab_data:
    if not tables:
        st.info("Upload a file to begin.")
    else:
        key = st.selectbox("Choose sheet/table", list(tables.keys()))
        df = tables[key].copy()
        st.success(f"Loaded: **{key}** — shape {df.shape}")
        with st.expander("🔍 Preview (first 12 rows)"):
            st.dataframe(df.head(12), use_container_width=True)
        date_cols, num_cols, txt_cols = detect_types(df)
        c1, c2, c3 = st.columns(3)
        c1.metric("Rows", f"{len(df):,}")
        c2.metric("Columns", f"{len(df.columns):,}")
        c3.metric("Numeric columns", f"{len(num_cols)}")
        st.markdown("**Detected types**")
        st.write(f"📅 Date-like: {date_cols or '—'}")
        st.write(f"🔢 Numeric-like: {num_cols or '—'}")
        st.write(f"🔤 Text-like: {txt_cols or '—'}")

# -------------------- Visualize tab --------------------
with tab_viz:
    if not tables:
        st.info("Upload data to visualize.")
    else:
        key = st.selectbox("Active sheet/table", list(tables.keys()), key="viz_key")
        df = tables[key].copy()
        date_cols, num_cols, txt_cols = detect_types(df)

        # X axis
        x_candidates = date_cols + txt_cols + [c for c in df.columns if c not in date_cols + txt_cols]
        x_col = st.selectbox("X axis", x_candidates, index=0)

        # Safe date handling
        if x_col in date_cols:
            df[x_col] = pd.to_datetime(df[x_col], errors="coerce")
            df = df.dropna(subset=[x_col])
            if df.empty:
                st.warning("No valid dates found after cleaning.")
            else:
                min_d, max_d = df[x_col].min(), df[x_col].max()
                # Fallbacks
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
                    st.warning("Date input failed; using full range automatically.")
                    rng = [min_d.date(), max_d.date()]

                if isinstance(rng, (list, tuple)) and len(rng) == 2:
                    df = df[(df[x_col] >= pd.to_datetime(rng[0])) & (df[x_col] <= pd.to_datetime(rng[1]))]

                # Resample
                freq = st.selectbox("Resample (if time axis)", ["None","D (Daily)","W (Weekly)","M (Monthly)","Q (Quarterly)","Y (Yearly)"], index=2)
                agg = st.selectbox("Aggregation", ["mean","sum","median","max","min"], index=0)
                if freq != "None":
                    code = freq.split()[0]  # D/W/M/Q/Y
                    df = safe_resample(df, x_col, code, agg)

        # Y axes
        if not num_cols:
            st.warning("No numeric columns detected to plot.")
            st.stop()
        y_cols = st.multiselect("Y axis (one or more)", num_cols, default=num_cols[:1])

        # Units & currency
        st.markdown("### Units / Currency (optional)")
        use_fx = st.checkbox("Convert currency")
        if use_fx:
            base = st.text_input("From (e.g., USD)", "USD").upper()
            quote = st.text_input("To (e.g., GBP)", "GBP").upper()
            live = fetch_fx_rate(base, quote)
            manual = st.number_input("Manual rate (multiply)", min_value=0.0, value=0.0, step=0.0001, format="%.6f")
            fx = manual if manual > 0 else (live if live else 1.0)
            if live:
                st.success(f"Live rate: 1 {base} = {live:.6f} {quote}")
            for c in y_cols:
                df[c] = pd.to_numeric(df[c], errors="coerce") * fx

        scale = st.selectbox("Display scale", ["None","Thousands (K)","Millions (M)","Billions (B)"], index=0)
        scale_factor = {"None":1,"Thousands (K)":1e3,"Millions (M)":1e6,"Billions (B)":1e9}[scale]
        for c in y_cols:
            df[c] = pd.to_numeric(df[c], errors="coerce") / scale_factor

        # Chart options
        st.markdown("### Chart")
        chart_type = st.selectbox("Type", ["Line","Area","Bar","Stacked Bar","Scatter","Pie","Dual Axis (left/right)"])
        colors = {c: st.color_picker(f"Color for {c}", "#1f77b4") for c in y_cols}
        line_w = st.slider("Line width", 1, 6, 2)
        marker_s = st.slider("Marker size", 3, 12, 6)
        y_tick = st.number_input("Y tick step (0=auto)", min_value=0, value=0)
        log_y = st.checkbox("Log scale Y", value=False)

        # Build figure
        fig = go.Figure()
        if chart_type in ["Line","Area","Bar","Stacked Bar","Scatter","Dual Axis (left/right)"]:
            for i, y in enumerate(y_cols):
                yaxis = "y" if (chart_type != "Dual Axis (left/right)" or i % 2 == 0) else "y2"
                if chart_type in ["Line","Area"]:
                    fig.add_trace(go.Scatter(
                        x=df[x_col], y=df[y], mode="lines+markers",
                        name=y, line=dict(color=colors[y], width=line_w),
                        marker=dict(size=marker_s), fill="tozeroy" if chart_type=="Area" else None, yaxis=yaxis
                    ))
                elif chart_type == "Scatter":
                    fig.add_trace(go.Scatter(
                        x=df[x_col], y=df[y], mode="markers",
                        name=y, marker=dict(size=marker_s, color=colors[y]), yaxis=yaxis
                    ))
                else:
                    fig.add_trace(go.Bar(
                        x=df[x_col], y=df[y], name=y, marker_color=colors[y], yaxis=yaxis
                    ))
            if chart_type == "Stacked Bar":
                fig.update_layout(barmode="stack")
            else:
                fig.update_layout(barmode="group")
        elif chart_type == "Pie":
            if len(y_cols) != 1:
                st.error("Pie requires exactly one Y column.")
                st.stop()
            fig = px.pie(df, names=x_col, values=y_cols[0], color_discrete_sequence=px.colors.qualitative.Bold)

        # Layout & axes
        layout = dict(template="plotly_white", title=f"{', '.join(y_cols)} vs {x_col}", xaxis_title=x_col, height=600)
        if chart_type == "Dual Axis (left/right)" and len(y_cols) > 1:
            layout.update(yaxis=dict(title=y_cols[0]), yaxis2=dict(title=y_cols[1], overlaying="y", side="right"))
        fig.update_layout(**layout)
        if y_tick > 0: fig.update_yaxes(dtick=y_tick)
        if log_y: fig.update_yaxes(type="log")

        st.plotly_chart(fig, use_container_width=True)
        st.session_state["last_fig"] = fig
        st.session_state["last_df"] = df

# -------------------- Insights tab --------------------
with tab_insight:
    if not tables:
        st.info("Upload data to analyze.")
    else:
        key = st.selectbox("Active sheet/table", list(tables.keys()), key="ins_key")
        df = tables[key].copy()
        date_cols, num_cols, _ = detect_types(df)

        st.subheader("Summary Statistics")
        if num_cols:
            st.dataframe(df[num_cols].describe().T.round(3))
        else:
            st.info("No numeric columns detected.")

        st.subheader("Narrative Insights (offline)")
        bullets = []
        if num_cols:
            for c in num_cols[:8]:
                s = pd.to_numeric(df[c], errors="coerce").dropna()
                if s.empty: continue
                trend = "rising 📈" if s.iloc[-1] > s.iloc[0] else "declining 📉"
                bullets.append(f"**{c}** — avg {s.mean():,.2f}, range {s.min():,.2f}–{s.max():,.2f}, {trend}.")
        if date_cols and num_cols:
            tcol = date_cols[0]
            tmp = df[[tcol] + num_cols].copy()
            tmp[tcol] = pd.to_datetime(tmp[tcol], errors="coerce")
            tmp = tmp.dropna(subset=[tcol]).sort_values(tcol)
            if not tmp.empty:
                m = tmp.set_index(tcol)[num_cols[0]].resample("M").mean()
                if len(m) >= 13:
                    yoy = (m.iloc[-1] - m.iloc[-13]) / (abs(m.iloc[-13]) + 1e-9) * 100
                    bullets.append(f"**{num_cols[0]}** YoY ≈ {yoy:+.1f}%")
        st.markdown("<br>".join(f"• {b}" for b in bullets) or "No insights available.", unsafe_allow_html=True)

        if len(num_cols) >= 2:
            st.subheader("Correlation Matrix")
            st.dataframe(df[num_cols].corr().round(3))

# -------------------- Export tab --------------------
with tab_export:
    if not tables:
        st.info("Upload data to export.")
    else:
        df_to_export = st.session_state.get("last_df")
        fig_to_export = st.session_state.get("last_fig")

        st.subheader("Data exports")
        c1, c2, c3 = st.columns(3)
        with c1:
            st.download_button("⬇️ CSV", (df_to_export or pd.DataFrame()).to_csv(index=False).encode("utf-8"),
                               "cleaned_data.csv", "text/csv")
        with c2:
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                (df_to_export or pd.DataFrame()).to_excel(w, index=False, sheet_name="Data")
            st.download_button("⬇️ Excel", buf.getvalue(), "cleaned_data.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with c3:
            st.download_button("⬇️ JSON",
                               (df_to_export or pd.DataFrame()).to_json(orient="records").encode("utf-8"),
                               "cleaned_data.json", "application/json")

        st.markdown("---")
        st.subheader("Chart exports")
        if fig_to_export is None:
            st.info("Create a chart in the Visualize tab first.")
        else:
            png = fig_to_png_safe(fig_to_export)
            pdf = fig_to_pdf_safe(fig_to_export)
            d1, d2, d3 = st.columns(3)
            with d1:
                if png:
                    st.download_button("⬇️ PNG", png, "chart.png", "image/png")
                else:
                    st.warning("PNG export unavailable (Kaleido missing or disabled).")
            with d2:
                if pdf:
                    st.download_button("⬇️ PDF", pdf, "chart.pdf", "application/pdf")
                else:
                    st.warning("PDF export unavailable (Kaleido missing or disabled).")
            with d3:
                pptx_bytes = pptx_with_chart_failsafe(fig_to_export, title="Finance Dashboard — Chart")
                st.download_button("⬇️ PPTX", pptx_bytes, "chart_slide.pptx",
                                   "application/vnd.openxmlformats-officedocument.presentationml.presentation")

# -------------------- Settings tab --------------------
with tab_settings:
    st.checkbox("Disable image exports (Option 4)", value=st.session_state["disable_image_exports"],
                key="disable_image_exports",
                help="Turn this on if your environment lacks Kaleido/Chrome. Data exports still work.")
    st.markdown("""
**Notes**
- Image export uses Plotly Kaleido. If your cloud environment lacks Chrome-like binaries, image export may be unavailable.  
- Option 3 is baked in: PPTX export falls back gracefully so your app never crashes.  
- Option 4 lets you disable all image exports explicitly.
""")
