from pathlib import Path
import tempfile

import streamlit as st

from agent import (
    load_source_a,
    load_source_b,
    calculate_dashboard,
)

from powerpoint import (
    build_powerpoint,
)


# =========================================================
# PAGE
# =========================================================

st.set_page_config(
    page_title="CSU Prime Dashboard",
    page_icon="📊",
    layout="centered",
)

st.title("CSU Prime Dashboard Generator")

st.write(
    "Upload the monthly reporting files, "
    "select the reporting period, "
    "and generate the PowerPoint dashboard."
)


# =========================================================
# FILE UPLOADS
# =========================================================

source_a_file = st.file_uploader(
    "Upload Source A — Legal Entity / Completion / ARVE",
    type=["xlsx"],
)

source_b_file = st.file_uploader(
    "Upload Source B — Timesheets per Project",
    type=["xlsx"],
)


# =========================================================
# REPORTING PERIOD
# =========================================================

reporting_year = st.number_input(
    "Reporting year",
    min_value=2025,
    max_value=2035,
    value=2026,
    step=1,
)


month_options = {
    "January": 1,
    "February": 2,
    "March": 3,
    "April": 4,
    "May": 5,
    "June": 6,
    "July": 7,
    "August": 8,
    "September": 9,
    "October": 10,
    "November": 11,
    "December": 12,
}


selected_month_names = st.multiselect(
    "Reporting months",
    options=list(month_options.keys()),
    default=["July", "August"],
)


selected_months = [
    month_options[name]
    for name in selected_month_names
]


# =========================================================
# GENERATE
# =========================================================

if st.button(
    "Generate CSU Dashboard",
    type="primary",
):

    if source_a_file is None:
        st.error(
            "Please upload Source A."
        )
        st.stop()

    if source_b_file is None:
        st.error(
            "Please upload Source B."
        )
        st.stop()

    if not selected_months:
        st.error(
            "Please select at least one reporting month."
        )
        st.stop()


    with st.spinner(
        "Processing files and generating dashboard..."
    ):

        try:

            # -------------------------------------------------
            # SAVE UPLOADED FILES TEMPORARILY
            # -------------------------------------------------

            temp_dir = Path(
                tempfile.mkdtemp(
                    prefix="csu_dashboard_"
                )
            )

            source_a_path = (
                temp_dir
                / "source_a.xlsx"
            )

            source_b_path = (
                temp_dir
                / "source_b.xlsx"
            )


            with open(
                source_a_path,
                "wb",
            ) as f:

                f.write(
                    source_a_file.getbuffer()
                )


            with open(
                source_b_path,
                "wb",
            ) as f:

                f.write(
                    source_b_file.getbuffer()
                )


            # -------------------------------------------------
            # LOAD
            # -------------------------------------------------

            source_a = load_source_a(
                source_a_path
            )

            source_b = load_source_b(
                source_b_path
            )


            # -------------------------------------------------
            # CALCULATE
            # -------------------------------------------------

            metrics = calculate_dashboard(
                source_a,
                source_b,
                int(reporting_year),
                selected_months,
            )


            # -------------------------------------------------
            # SHOW KPI PREVIEW
            # -------------------------------------------------

            k = metrics[
                "kpis"
            ]


            st.success(
                "Calculations completed successfully."
            )


            col1, col2, col3 = st.columns(3)

            with col1:

                st.metric(
                    "HC Tagged",
                    (
                        f"{k['hc_tagged']} "
                        f"/ {k['total_hc']}"
                    ),
                )

                st.metric(
                    "Completion",
                    f"{k['completion_rate']:.0%}",
                )


            with col2:

                st.metric(
                    "HC Compliance",
                    f"{k['hc_compliance']:.0%}",
                )

                st.metric(
                    "ARVE",
                    f"{k['arve']:.0%}",
                )


            with col3:

                st.metric(
                    "ARVI",
                    f"{k['arvi']:.0%}",
                )

                st.metric(
                    "Overbooking",
                    f"{k['overbooking']:.0%}",
                )


            # -------------------------------------------------
            # GENERATE POWERPOINT
            # -------------------------------------------------

            output_pptx = Path(
                build_powerpoint(
                    metrics
                )
            )


            # -------------------------------------------------
            # DOWNLOAD
            # -------------------------------------------------

            with open(
                output_pptx,
                "rb",
            ) as f:

                pptx_bytes = (
                    f.read()
                )


            st.success(
                "PowerPoint generated successfully."
            )


            st.download_button(
                label="Download PowerPoint",
                data=pptx_bytes,
                file_name=output_pptx.name,
                mime=(
                    "application/"
                    "vnd.openxmlformats-officedocument."
                    "presentationml.presentation"
                ),
            )


        except Exception as exc:

            st.error(
                "Generation failed."
            )

            st.exception(
                exc
            )