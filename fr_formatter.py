import streamlit as st
import pandas as pd
import io
import zipfile

# Set a professional page configuration
st.set_page_config(page_title="Farmer Data Pro", page_icon="🌾", layout="wide")

# FIX: Custom CSS for visibility and card contrast
st.markdown("""
    <style>
    .main { background-color: #f0f2f6; }
    
    /* FIX: Dark background for metrics so white text is visible */
    [data-testid="metric-container"] {
        background-color: #1a1c24; /* Dark background */
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        border: 1px solid #3e4452;
    }

    /* FIX: Ensure Label and Value are bright and readable */
    [data-testid="stMetricLabel"] {
        color: #e0e0e0 !important;
        font-weight: 600;
    }
    [data-testid="stMetricValue"] {
        color: #ffffff !important;
    }

    footer {visibility: hidden;}
    .footer { position: fixed; bottom: 0; width: 100%; text-align: center; color: #6c757d; padding: 10px; }
    </style>
""", unsafe_allow_html=True)

st.title("🌾 FR Excel formatter")
st.markdown("Upload multiple Excel files to deduplicate and group records by Village.")

# 1. File Upload Section
uploaded_files = st.file_uploader("Upload Excel files", type="xlsx, csv", accept_multiple_files=True)

if uploaded_files:
    # Read and combine files
    dfs = [pd.read_excel(f) for f in uploaded_files]
    raw_df = pd.concat(dfs, ignore_index=True)
    
    # 2. Data Processing
    processed_df = raw_df.groupby(['Bucket ID', 'Village LGD Code']).agg({
        "Village Name": lambda x: ", ".join(map(str, pd.unique(x))),
        "Farmer Name": "last",
        "Identifier Name": "last",
        "Farmer Mobile Number": "last",
        "Survey Number": lambda x: ", ".join(map(str, pd.unique(x))),
        "Sub Survey Number": lambda x: ", ".join(map(str, pd.unique(x)))
    }).reset_index()

    # Rename for cleaner UI
    processed_df.columns = [
        'Bucket ID', 'Village LGD Code', 'Village_Name', 
        'Farmer_Name', 'Father_Name', 'Contact_No', 'Sy_No', 'Sub_Sy'
    ]

    # 3. Stats Dashboard (NOW VISIBLE)
    st.subheader("📊 Processing Overview")
    m1, m2, m3 = st.columns(3)
    m1.metric("Total Rows Processed", f"{len(raw_df)}")
    m2.metric("Unique Records Remaining", f"{len(processed_df)}")
    m3.metric("Duplicates Removed", f"{len(raw_df) - len(processed_df)}")

    st.divider()

    # 4. Preview (UPDATED TO TOP 10)
    st.subheader("🔍 Data Preview")
    # Using head(10) as requested
    st.dataframe(processed_df.head(10), use_container_width=True)

    # 5. Download Center
    st.subheader("💾 Download Center")
    col_full, col_split = st.columns(2)

    # Option A: Full Download
    with col_full:
        st.info("📦 **Combined File**")
        output_full = io.BytesIO()
        with pd.ExcelWriter(output_full, engine='xlsxwriter') as writer:
            processed_df.to_excel(writer, index=False, sheet_name='All_Villages')
        
        st.download_button(
            label="Download Full Excel",
            data=output_full.getvalue(),
            file_name="Full_Farmer_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    # Option B: Split by Village (ZIP)
    with col_split:
        st.success("📁 **Split by Village**")
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            unique_villages = processed_df['Village_Name'].unique()
            for village in unique_villages:
                v_df = processed_df[processed_df['Village_Name'] == village]
                v_output = io.BytesIO()
                with pd.ExcelWriter(v_output, engine='xlsxwriter') as writer:
                    v_df.to_excel(writer, index=False)
                
                clean_name = "".join([c for c in str(village) if c.isalnum() or c in (' ', '_')]).rstrip()
                zip_file.writestr(f"{clean_name}.xlsx", v_output.getvalue())

        st.download_button(
            label="Download All Villages (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="Villages_Split_Reports.zip",
            mime="application/zip",
            use_container_width=True
        )

else:
    st.info("Waiting for files to be uploaded...")

st.markdown('<div class="footer">Built with ❤️ by Sandeep Kumar</div>', unsafe_allow_html=True)