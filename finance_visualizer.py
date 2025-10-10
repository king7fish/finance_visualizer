# ------------------ EXPORT TAB ------------------
def safe_export_fig(fig, fmt="png"):
    """Return bytes safely, fallback if Kaleido fails."""
    if fig is None:
        return placeholder_png("No Chart Available", color=(0, 0, 0))
    if st.session_state.get("disable_image_exports", False):
        return placeholder_png("Image export disabled", color=(255, 165, 0))
    try:
        return fig.to_image(format=fmt, engine="kaleido")
    except Exception as e:
        st.warning(f"⚠️ {fmt.upper()} export fallback used: {e}")
        return placeholder_png(f"{fmt.upper()} export unavailable", color=(255, 0, 0))

with tab_export:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    if "v9_plotted" not in st.session_state or st.session_state["v9_plotted"].empty:
        st.info("Generate a chart first.")
    else:
        fig_to_export = st.session_state.get("v9_fig")
        df_plot_long = st.session_state["v9_plotted"]
        df_raw_long = st.session_state["v9_raw_long"]

        st.subheader("Export Data")
        which = st.radio("Choose data to export:", ["Plotted (filtered + scaled)", "Raw (aligned/cleaned)"], horizontal=True)
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
        st.subheader("Export Chart")

        png_bytes = safe_export_fig(fig_to_export, "png")
        pdf_bytes = safe_export_fig(fig_to_export, "pdf")
        pptx_bytes = pptx_with_chart_failsafe(fig_to_export, title="Finance Dashboard - Chart")

        d1, d2, d3 = st.columns(3)
        with d1: st.download_button("Download PNG", png_bytes, "chart.png", "image/png")
        with d2: st.download_button("Download PDF", pdf_bytes, "chart.pdf", "application/pdf")
        with d3: st.download_button("Download PPTX", pptx_bytes, "chart_slide.pptx",
                                    "application/vnd.openxmlformats-officedocument.presentationml.presentation")
    st.markdown('</div>', unsafe_allow_html=True)
