import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import io
import json
from pptx import Presentation
from docx import Document
from PyPDF2 import PdfReader

st.set_page_config(page_title="Finance Dashboard Pro", layout="wide")

# ---------------- HEADER ----------------
st.markdown("""
    <h1 style='text-align:center; color:#3A86FF;'>📊 Finance Dashboard Pro</h1>
    <p style='text-align:center; color:#6c757d;'>Analyze, visualize, and understand your financial data instantly</p>
""", unsafe_allow_html=True)

# ---------------- FILE UPLOAD ----------------
st.sidebar.header("📁 Upload Your Data")
uploaded_file = st.sidebar.file_uploader("Upload a file", type=["xlsx", "csv", "json", "docx", "pptx", "pdf"])

if uploaded_file:
    file_name = uploaded_file.name.lower()

    # Detect file type
    if file_name.endswith(".xlsx") or file_name.endswith(".xls"):
        sheet_names = pd.ExcelFile(uploaded_file).sheet_names
        selected_sheet = st.sidebar.selectbox("Choose a sheet", sheet_names)
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)

    elif file_name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)

    elif file_name.endswith(".json"):
        df = pd.read_json(uploaded_file)

    elif file_name.endswith(".docx"):
        doc = Document(uploaded_file)
        text = "\n".join([p.text for p in doc.paragraphs])
        df = pd.DataFrame({"Extracted Text": text.split("\n")})

    elif file_name.endswith(".pptx"):
        prs = Presentation(uploaded_file)
        slides = []
        for slide in prs.slides:
            content = " ".join([shape.text for shape in slide.shapes if hasattr(shape, "text")])
            slides.append(content)
        df = pd.DataFrame({"Slide Content": slides})

    elif file_name.endswith(".pdf"):
        pdf = PdfReader(uploaded_file)
        text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
        df = pd.DataFrame({"Extracted Text": text.split("\n")})

    else:
        st.error("Unsupported file type.")
        st.stop()

    # ---------------- CLEANING ----------------
    st.subheader("🧽 Data Cleaning")
    df.columns = df.columns.map(lambda x: str(x).strip().replace('\n', ' '))

    for col in df.columns:
        df[col] = pd.to_numeric(df[col].replace({',': ''}, regex=True), errors='ignore')

    def is_mostly_text(row):
        numeric_count = sum(pd.to_numeric(row, errors='coerce').notna())
        return numeric_count < len(row) / 2

    text_rows = df[df.apply(is_mostly_text, axis=1)]
    df_cleaned = df[~df.apply(is_mostly_text, axis=1)]

    if len(text_rows) > 0:
        with st.expander("⚠️ View Removed Text Rows"):
            st.dataframe(text_rows)

    st.success(f"Cleaned data: {df_cleaned.shape[0]} rows remain (removed {len(text_rows)} messy rows).")
    df = df_cleaned.reset_index(drop=True)

    # ---------------- DISPLAY ----------------
    with st.expander("🔍 View Cleaned Data"):
        st.dataframe(df)

    # ---------------- AXIS SELECTION ----------------
    st.sidebar.header("⚙️ Visualization Settings")
    x_col = st.sidebar.selectbox("X-Axis", df.columns)
    y_col = st.sidebar.multiselect("Y-Axis", df.columns, default=[df.columns[1]] if len(df.columns) > 1 else None)
    chart_type = st.sidebar.selectbox("Chart Type", ["Line", "Bar", "Area", "Scatter", "Pie"])
    currency = st.sidebar.selectbox("Currency Symbol", ["$", "€", "£", "¥", "None"])
    unit = st.sidebar.text_input("Unit (e.g., hrs, kg, %)", "")

    # ---------------- VISUALIZATION ----------------
    if x_col and y_col:
        if chart_type == "Line":
            fig = px.line(df, x=x_col, y=y_col, markers=True, title="📈 Line Chart")
        elif chart_type == "Bar":
            fig = px.bar(df, x=x_col, y=y_col, title="📊 Bar Chart")
        elif chart_type == "Area":
            fig = px.area(df, x=x_col, y=y_col, title="🌊 Area Chart")
        elif chart_type == "Scatter":
            fig = px.scatter(df, x=x_col, y=y_col, title="🔹 Scatter Plot")
        elif chart_type == "Pie" and len(y_col) == 1:
            fig = px.pie(df, names=x_col, values=y_col[0], title="🥧 Pie Chart")
        else:
            st.error("Pie charts require exactly one Y column.")
            st.stop()

        # Label updates
        fig.update_layout(
            title_font=dict(size=22, color="#3A86FF"),
            xaxis_title=f"{x_col} ({unit})" if unit else x_col,
            yaxis_title=f"{', '.join(y_col)} ({currency}{unit})" if currency or unit else ", ".join(y_col),
            template="plotly_dark",
            height=600
        )

        st.plotly_chart(fig, use_container_width=True)

    # ---------------- AI INSIGHTS ----------------
    st.subheader("🧠 AI Insights (Offline Mode)")

    numeric_df = df.select_dtypes(include=[np.number])
    insights = []

    if not numeric_df.empty:
        desc = numeric_df.describe()
        for col in numeric_df.columns:
            mean_val = desc[col]['mean']
            std_val = desc[col]['std']
            min_val = desc[col]['min']
            max_val = desc[col]['max']
            trend = "rising 📈" if mean_val < max_val * 0.9 else "stable ⚖️" if std_val < mean_val * 0.3 else "volatile ⚠️"
            insights.append(f"**{col}** has an average of {mean_val:.2f} ({trend}). Range: {min_val:.2f}–{max_val:.2f}.")

        st.markdown("<br>".join(insights), unsafe_allow_html=True)

        corr = numeric_df.corr()
        if len(corr.columns) > 1:
            st.markdown("### 🔗 Correlation Matrix")
            st.dataframe(corr)
    else:
        st.warning("No numeric data available for insight generation.")

    # ---------------- DOWNLOAD ----------------
    st.download_button(
        label="💾 Download Cleaned Data (CSV)",
        data=df.to_csv(index=False).encode('utf-8'),
        file_name="cleaned_finance_data.csv",
        mime="text/csv"
    )

else:
    st.info("👆 Upload a file from the sidebar to get started!")
