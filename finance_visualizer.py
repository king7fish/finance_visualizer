# finance_visualizer_pro.py

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import json
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from datetime import datetime
from docx import Document

st.set_page_config(page_title="Finance Data Visualizer Pro", layout="wide")

st.title("📊 Finance Data Visualizer Pro")
st.markdown(
    """
    Upload your **financial, corporate, or quantitative** data and generate interactive visual insights.  
    Supports multiple file types: `.xlsx`, `.csv`, `.json`, `.docx`, `.pptx`.
    """
)

# ------------------- File Upload -------------------
uploaded_file = st.file_uploader(
    "Upload your data file",
    type=["xlsx", "csv", "json", "docx", "pptx"],
    help="Supported file types: Excel, CSV, JSON, Word, PowerPoint",
)

if uploaded_file:
    file_type = uploaded_file.name.split(".")[-1].lower()

    if file_type == "xlsx":
        xls = pd.ExcelFile(uploaded_file)
        sheet_name = st.selectbox("Select sheet to load", xls.sheet_names)
        df = xls.parse(sheet_name)
    elif file_type == "csv":
        df = pd.read_csv(uploaded_file)
    elif file_type == "json":
        df = pd.read_json(uploaded_file)
    elif file_type == "docx":
        document = Document(uploaded_file)
        text_data = [p.text for p in document.paragraphs if p.text.strip()]
        df = pd.DataFrame({"Paragraphs": text_data})
    elif file_type == "pptx":
        prs = Presentation(uploaded_file)
        slides_text = []
        for slide in prs.slides:
            slide_text = " ".join([shape.text for shape in slide.shapes if hasattr(shape, "text")])
            slides_text.append(slide_text)
        df = pd.DataFrame({"Slide Text": slides_text})
    else:
        st.error("Unsupported file type.")
        st.stop()

    st.success(f"✅ File loaded successfully! Shape: {df.shape}")

    # Display first few rows
    st.dataframe(df.head())

    # ------------------- Preprocessing -------------------
    st.subheader("🧹 Data Cleaning & Setup")

    df.columns = df.columns.astype(str)
    numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    date_cols = [col for col in df.columns if "date" in col.lower()]

    if not numeric_cols:
        st.warning("No numeric columns detected. Try a different file or clean the data first.")

    x_col = st.selectbox("Select X-axis (usually Date or Index)", df.columns)
    y_col = st.selectbox("Select Y-axis (numeric column)", numeric_cols)

    # ------------------- Resampling Function -------------------
    def resample_time_series(df, time_col, freq='M', agg='mean'):
        """
        Resample time-series data by frequency while ignoring non-numeric columns.
        freq: 'D' (daily), 'W' (weekly), 'M' (monthly), 'Q' (quarterly), 'Y' (yearly)
        agg: 'mean', 'sum', 'median', etc.
        """
        try:
            df[time_col] = pd.to_datetime(df[time_col], errors='coerce')
            df = df.dropna(subset=[time_col])
            df = df.set_index(time_col)

            # Select numeric columns only
            numeric_df = df.select_dtypes(include=['number'])

            # Perform resampling safely
            resampled = numeric_df.resample(freq).agg(agg)

            # Reattach time as a column
            resampled.reset_index(inplace=True)

            return resampled
        except Exception as e:
            st.error(f"Resampling failed: {e}")
            return df

    # ------------------- Visualization Controls -------------------
    st.subheader("📈 Visualization Options")

    chart_type = st.selectbox(
        "Choose Chart Type",
        ["Line", "Bar", "Scatter", "Pie"],
        index=0
    )

    freq = st.selectbox(
        "Time Resampling Frequency",
        ["None", "Daily (D)", "Weekly (W)", "Monthly (M)", "Quarterly (Q)", "Yearly (Y)"],
        index=2
    )
    agg_method = st.selectbox("Aggregation Method", ["mean", "sum", "median", "max", "min"])

    if freq != "None":
        df = resample_time_series(df, x_col, freq=freq.split(" ")[-1][1:-1], agg=agg_method)

    currency_options = ["USD ($)", "EUR (€)", "GBP (£)", "JPY (¥)", "KES (KSh)", "Custom"]
    selected_currency = st.selectbox("Select Currency / Unit", currency_options)

    if selected_currency == "Custom":
        selected_currency = st.text_input("Enter custom currency or unit (e.g., kg, m/s², hours)")

    # ------------------- Chart Generation -------------------
    st.subheader("📊 Chart Preview")

    fig = None
    if chart_type == "Line":
        fig = px.line(df, x=x_col, y=y_col, title=f"{y_col} over {x_col}", markers=True)
    elif chart_type == "Bar":
        fig = px.bar(df, x=x_col, y=y_col, title=f"{y_col} by {x_col}")
    elif chart_type == "Scatter":
        fig = px.scatter(df, x=x_col, y=y_col, title=f"{y_col} vs {x_col}", trendline="ols")
    elif chart_type == "Pie":
        fig = px.pie(df, names=x_col, values=y_col, title=f"{y_col} Distribution by {x_col}")

    if fig:
        fig.update_layout(
            title_font=dict(size=20),
            xaxis_title=x_col,
            yaxis_title=f"{y_col} ({selected_currency})",
            template="plotly_white",
        )
        st.plotly_chart(fig, use_container_width=True)

    # ------------------- Export Options -------------------
    st.subheader("📤 Export Options")

    export_format = st.selectbox("Choose Export Format", ["CSV", "Excel", "JSON", "PPTX"])

    if export_format == "CSV":
        csv = df.to_csv(index=False).encode('utf-8')
        st.download_button("Download CSV", csv, file_name="visualized_data.csv", mime="text/csv")

    elif export_format == "Excel":
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        st.download_button("Download Excel", output.getvalue(), file_name="visualized_data.xlsx")

    elif export_format == "JSON":
        json_data = df.to_json(orient="records", indent=2)
        st.download_button("Download JSON", json_data, file_name="visualized_data.json")

    elif export_format == "PPTX":
        pptx_output = BytesIO()
        prs = Presentation()
        slide_layout = prs.slide_layouts[5]
        slide = prs.slides.add_slide(slide_layout)

        title_shape = slide.shapes.title or slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(1))
        title_shape.text = f"{y_col} Visualization"

        # Save chart as image and embed
        img_bytes = fig.to_image(format="png", width=1000, height=600, scale=2)
        img_stream = BytesIO(img_bytes)
        slide.shapes.add_picture(img_stream, Inches(1), Inches(1.5), Inches(8), Inches(4.5))

        prs.save(pptx_output)
        st.download_button("Download PowerPoint Slide", pptx_output.getvalue(), file_name="chart_slide.pptx")

    st.markdown("---")
    st.info("✨ Tip: Upload any file type — Excel, CSV, JSON, Word, or PowerPoint — and generate interactive visuals in seconds!")
