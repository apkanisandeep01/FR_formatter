import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime
import streamlit.components.v1 as components


# ==================================================
# PAGE CONFIG
# ==================================================
st.set_page_config(
    page_title="FR Excel Formatter",
    page_icon="🌾",
    layout="wide",
    initial_sidebar_state="collapsed"
)


# ==================================================
# MAIN APP CSS
# ==================================================
st.markdown(
    """
    <style>
        .stApp {
            background: linear-gradient(180deg, #f1f8f1 0%, #ffffff 42%);
            color: #16351a;
        }

        .block-container {
            max-width: 1180px;
            padding-top: 2rem;
            padding-bottom: 2rem;
        }

        h1, h2, h3, p, label {
            color: #1b5e20;
        }

        /* FILE UPLOADER */
        div[data-testid="stFileUploader"] {
            margin-top: .4rem;
            margin-bottom: 1.2rem;
        }

        div[data-testid="stFileUploader"] label,
        div[data-testid="stFileUploader"] label p {
            color: #0d4715 !important;
            font-size: 1.05rem !important;
            font-weight: 800 !important;
        }

        div[data-testid="stFileUploader"] section {
            border: 3px dashed #2e7d32 !important;
            border-radius: 16px !important;
            padding: 1rem !important;
            background: linear-gradient(135deg, #f4fbf4 0%, #eaf7eb 100%) !important;
            box-shadow: 0 4px 14px rgba(46, 125, 50, .10);
        }

        div[data-testid="stFileUploader"] section:hover {
            border-color: #1b5e20 !important;
            background: #e4f4e5 !important;
        }

        div[data-testid="stFileUploader"] small,
        div[data-testid="stFileUploader"] span,
        div[data-testid="stFileUploader"] p,
        div[data-testid="stFileUploader"] div {
            color: #234d28 !important;
            font-weight: 600 !important;
        }

        div[data-testid="stFileUploaderDropzone"] button {
            background: #2e7d32 !important;
            color: #ffffff !important;
            border: none !important;
            border-radius: 10px !important;
            font-weight: 800 !important;
        }

        div[data-testid="stFileUploaderDropzone"] button * {
            color: #ffffff !important;
        }

        /* UPLOADED FILE NAME */
        div[data-testid="stFileUploaderFile"] {
            background: #ffffff !important;
            border: 1px solid #9ccc9e !important;
            border-radius: 10px !important;
            box-shadow: 0 2px 6px rgba(27, 94, 32, .08);
        }

        div[data-testid="stFileUploaderFile"] * {
            color: #16351a !important;
            font-weight: 700 !important;
        }

        /* METRIC CARDS */
        [data-testid="metric-container"],
        div[data-testid="stMetric"] {
            background: linear-gradient(135deg, #ffffff 0%, #edf8ee 100%) !important;
            border: 1px solid #b7ddb9 !important;
            border-left: 6px solid #2e7d32 !important;
            border-radius: 14px !important;
            padding: 16px 18px !important;
            box-shadow: 0 3px 12px rgba(27, 94, 32, .10) !important;
        }

        [data-testid="metric-container"] *,
        div[data-testid="stMetric"] * {
            color: #1b5e20 !important;
            font-weight: 800 !important;
        }

        [data-testid="stMetricValue"] *,
        div[data-testid="stMetricValue"] * {
            color: #103b16 !important;
        }

        /* DATAFRAME */
        div[data-testid="stDataFrame"] {
            border: 1px solid #c8e6c9 !important;
            border-radius: 12px !important;
            overflow: hidden !important;
            background: #ffffff !important;
        }

        /* DOWNLOAD BUTTONS */
        div.stDownloadButton > button {
            background: #2e7d32 !important;
            color: #ffffff !important;
            border: none !important;
            border-radius: 12px !important;
            font-weight: 800 !important;
            padding: .75rem 1rem !important;
            width: 100%;
        }

        div.stDownloadButton > button * {
            color: #ffffff !important;
        }

        div.stDownloadButton > button:hover {
            background: #1b5e20 !important;
            box-shadow: 0 6px 16px rgba(46, 125, 50, .30);
            transform: translateY(-1px);
        }

        /* HIDE DEFAULT STREAMLIT FOOTER */
        footer {
            visibility: hidden;
        }

        /* MOBILE */
        @media (max-width: 640px) {
            .block-container {
                padding-top: 1.1rem;
                padding-left: 1rem;
                padding-right: 1rem;
                padding-bottom: 1rem;
            }

            [data-testid="metric-container"],
            div[data-testid="stMetric"] {
                padding: 12px 13px !important;
            }
        }
    </style>
    """,
    unsafe_allow_html=True
)


# ==================================================
# HEADER COMPONENT
# ==================================================
header_html = """
<style>
    body {
        margin: 0;
        padding: 0;
        font-family: Arial, sans-serif;
        background: transparent;
    }

    .fr-hero {
        background: linear-gradient(135deg, #1b5e20 0%, #2e7d32 55%, #558b2f 100%);
        border-radius: 18px;
        padding: 28px 30px;
        color: #ffffff;
        box-sizing: border-box;
        box-shadow: 0 10px 30px rgba(27, 94, 32, .24);
    }

    .badge {
        display: inline-block;
        margin-bottom: 11px;
        padding: 4px 11px;
        border-radius: 999px;
        background: rgba(255, 255, 255, .17);
        color: #ffffff;
        font-size: 12px;
        font-weight: 700;
        letter-spacing: .55px;
    }

    h1 {
        margin: 0 0 8px 0;
        color: #ffffff;
        font-size: 31px;
        line-height: 1.2;
    }

    p {
        margin: 0;
        max-width: 900px;
        color: #e8f5e9;
        font-size: 15px;
        line-height: 1.55;
    }

    @media (max-width: 640px) {
        .fr-hero {
            padding: 21px 18px;
            border-radius: 15px;
        }

        h1 {
            font-size: 25px;
        }

        p {
            font-size: 14px;
        }
    }
</style>

<div class="fr-hero">
    <div class="badge">FARMER RECORDS · EXCEL FORMATTER</div>
    <h1>🌾 FR Excel Formatter</h1>
    <p>
        Upload multiple farmer-record Excel files to create one cleaner village-level report.
        Download a combined Excel report or separate Excel files for each village.
    </p>
</div>
"""

components.html(
    header_html,
    height=185,
    scrolling=False
)


# ==================================================
# APP DESCRIPTION COMPONENT
# ==================================================
description_html = """
<style>
    body {
        margin: 0;
        padding: 0;
        font-family: Arial, sans-serif;
        background: transparent;
    }

    .how-it-works {
        box-sizing: border-box;
        margin: 0;
        padding: 16px 18px;
        border: 1px solid #c8e6c9;
        border-left: 6px solid #2e7d32;
        border-radius: 14px;
        background: linear-gradient(135deg, #ffffff 0%, #eff9f0 100%);
        box-shadow: 0 3px 12px rgba(27, 94, 32, .08);
    }

    h3 {
        margin: 0 0 7px 0;
        color: #1b5e20;
        font-size: 16px;
        font-weight: 800;
    }

    p {
        margin: 0;
        color: #234d28;
        font-size: 14px;
        line-height: 1.55;
    }

    b {
        color: #1b5e20;
    }

    @media (max-width: 640px) {
        .how-it-works {
            padding: 14px;
        }

        p {
            font-size: 13px;
        }
    }
</style>

<div class="how-it-works">
    <h3>🧹 How this formatter removes duplicates</h3>
    <p>
        The app combines all uploaded Excel files and groups records using the same
        <b>Bucket ID</b> and <b>Village LGD Code</b>. Records with the same pair are
        treated as one record. It keeps the latest farmer, father and mobile details,
        while joining unique village names, survey numbers and sub-survey numbers into
        one organised row.
    </p>
</div>
"""

components.html(
    description_html,
    height=135,
    scrolling=False
)


# ==================================================
# FILE UPLOAD SECTION
# ==================================================
uploaded_files = st.file_uploader(
    "📁 Upload Excel files",
    type="xlsx",
    accept_multiple_files=True
)


# ==================================================
# CORE DATA PROCESSING LOGIC
# ==================================================
if uploaded_files:

    # Read and combine files
    dfs = [pd.read_excel(file) for file in uploaded_files]
    raw_df = pd.concat(dfs, ignore_index=True)

    # Group records by Bucket ID and Village LGD Code
    processed_df = raw_df.groupby(
        ["Bucket ID", "Village LGD Code"]
    ).agg({
        "Village Name": lambda values: ", ".join(
            map(str, pd.unique(values))
        ),
        "Farmer Name": "last",
        "Identifier Name": "last",
        "Farmer Mobile Number": "last",
        "Survey Number": lambda values: ", ".join(
            map(str, pd.unique(values))
        ),
        "Sub Survey Number": lambda values: ", ".join(
            map(str, pd.unique(values))
        )
    }).reset_index()

    # Rename columns for cleaner UI
    processed_df.columns = [
        "Bucket ID",
        "Village LGD Code",
        "Village_Name",
        "Farmer_Name",
        "Father_Name",
        "Contact_No",
        "Sy_No",
        "Sub_Sy"
    ]

    # ==================================================
    # PROCESSING OVERVIEW
    # ==================================================
    st.subheader("📊 Processing Overview")

    metric_1, metric_2, metric_3 = st.columns(3)

    metric_1.metric(
        "Total Rows Processed",
        f"{len(raw_df):,}"
    )

    metric_2.metric(
        "Unique Records Remaining",
        f"{len(processed_df):,}"
    )

    metric_3.metric(
        "Duplicates Removed",
        f"{len(raw_df) - len(processed_df):,}"
    )

    st.divider()

    # ==================================================
    # DATA PREVIEW
    # ==================================================
    st.subheader("🔍 Data Preview")

    st.dataframe(
        processed_df.head(5),
        use_container_width=True,
        hide_index=True
    )

    # ==================================================
    # DOWNLOAD CENTER
    # ==================================================
    st.subheader("💾 Download Center")

    col_full, col_split = st.columns(2)

    # ----------------------------------------------
    # FULL COMBINED EXCEL DOWNLOAD
    # ----------------------------------------------
    with col_full:
        st.info("📦 **Combined File**")

        output_full = io.BytesIO()

        with pd.ExcelWriter(
            output_full,
            engine="xlsxwriter"
        ) as writer:
            processed_df.to_excel(
                writer,
                index=False,
                sheet_name="All_Villages"
            )

        st.download_button(
            label="⬇️ Download Full Excel",
            data=output_full.getvalue(),
            file_name="Full_Farmer_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    # ----------------------------------------------
    # VILLAGE-WISE ZIP DOWNLOAD
    # ----------------------------------------------
    with col_split:
        st.success("📁 **Split by Village**")

        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(
            zip_buffer,
            "a",
            zipfile.ZIP_DEFLATED,
            False
        ) as zip_file:

            unique_villages = processed_df["Village_Name"].unique()

            for village in unique_villages:
                village_df = processed_df[
                    processed_df["Village_Name"] == village
                ]

                village_output = io.BytesIO()

                with pd.ExcelWriter(
                    village_output,
                    engine="xlsxwriter"
                ) as writer:
                    village_df.to_excel(
                        writer,
                        index=False,
                        sheet_name="Village_Report"
                    )

                clean_name = "".join(
                    [
                        character
                        for character in str(village)
                        if character.isalnum() or character in (" ", "_")
                    ]
                ).rstrip()

                if not clean_name:
                    clean_name = "Unknown_Village"

                zip_file.writestr(
                    f"{clean_name}.xlsx",
                    village_output.getvalue()
                )

        st.download_button(
            label="🗂️ Download All Villages (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="Villages_Split_Reports.zip",
            mime="application/zip",
            use_container_width=True
        )

else:
    st.info(
        "👆 Upload one or more Excel files to begin formatting farmer records."
    )


# ==================================================
# FOOTER COMPONENT
# ==================================================
PORTFOLIO_URL = "https://apkanisandeep01.github.io/my-portfolio/"
GITHUB_URL = "https://github.com/apkanisandeep01"
LINKEDIN_URL = "https://www.linkedin.com/in/sandeep-data-analyst-uk"
CONTACT_EMAIL = "apkansiandeep00@gmail.com"
APP_VERSION = "v1.0"


footer_html = f"""
<style>
    body {{
        margin: 0;
        padding: 0;
        font-family: Arial, sans-serif;
        background: transparent;
    }}

    .fr-footer {{
        box-sizing: border-box;
        margin-top: 0;
        padding: 16px 18px 12px;
        border-radius: 14px;
        background: linear-gradient(135deg, #0f2e13 0%, #1b5e20 100%);
        color: #d7ead7;
        text-align: center;
        box-shadow: 0 8px 20px rgba(15, 46, 19, .20);
    }}

    .brand {{
        font-size: 17px;
        font-weight: 800;
        color: #ffffff;
        margin-bottom: 3px;
    }}

    .tagline {{
        font-size: 12px;
        color: #a9d0a9;
        margin-bottom: 8px;
    }}

    .dev {{
        font-size: 12px;
        color: #e6f2e6;
    }}

    .dev b {{
        color: #ffffff;
    }}

    .links {{
        margin-top: 9px;
    }}

    .links a {{
        display: inline-block;
        margin: 3px 4px;
        padding: 6px 12px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 600;
        color: #ffffff !important;
        text-decoration: none;
        background: rgba(255, 255, 255, .10);
        border: 1px solid rgba(255, 255, 255, .20);
    }}

    .links a:hover {{
        background: #43a047;
        border-color: #43a047;
    }}

    .features {{
        display: flex;
        justify-content: center;
        flex-wrap: wrap;
        gap: 6px;
        margin-top: 10px;
        font-size: 11px;
    }}

    .features span {{
        background: rgba(255, 255, 255, .08);
        padding: 3px 8px;
        border-radius: 7px;
        color: #e6f2e6;
    }}

    .bottom {{
        margin-top: 10px;
        padding-top: 8px;
        border-top: 1px solid rgba(255, 255, 255, .12);
        font-size: 10px;
        color: #92b892;
    }}

    @media (max-width: 640px) {{
        .fr-footer {{
            padding: 14px 10px 10px;
        }}

        .brand {{
            font-size: 16px;
        }}

        .links a {{
            font-size: 11px;
            padding: 5px 9px;
            margin: 3px 2px;
        }}

        .features {{
            gap: 4px;
            font-size: 10px;
        }}

        .features span {{
            padding: 3px 6px;
        }}
    }}
</style>

<div class="fr-footer">
    <div class="brand">🌾 FR Excel Formatter</div>

    <div class="tagline">
        Smart farmer record formatting and village-wise Excel reporting
    </div>

    <div class="dev">
        Developed by <b>Sandeep Kumar Apkani</b>
    </div>

    <div class="links">
        <a href="{PORTFOLIO_URL}" target="_blank">🌐 Portfolio</a>
        <a href="{GITHUB_URL}" target="_blank">💻 GitHub</a>
        <a href="{LINKEDIN_URL}" target="_blank">🔗 LinkedIn</a>
        <a href="mailto:{CONTACT_EMAIL}">✉️ Contact</a>
    </div>

    <div class="features">
        <span>⚡ Fast formatting</span>
        <span>🧹 Duplicate removal</span>
        <span>📁 Village split reports</span>
        <span>📊 Excel ready</span>
    </div>

    <div class="bottom">
        © {datetime.now().year} FR Excel Formatter · {APP_VERSION} · Made with ❤️ for Users
    </div>
</div>
"""

components.html(
    footer_html,
    height=235,
    scrolling=False
)