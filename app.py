import streamlit as st
import pandas as pd
from processor import process_files
from datetime import datetime

st.set_page_config(page_title="Excel Processor for SAP", layout="wide")

st.title("Excel Processor for SAP")
st.markdown("""
This tool automates the processing of employee lists.
1. Upload the **Master Employee List**.
2. Upload one or more **Updated Employee Lists** (from HR).
3. Click **Process Files** to generate the required SAP and Master files.
""")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Upload Master File")
    master_file = st.file_uploader("Choose Master Excel File", type=['xlsx'], key="master_uploader")

with col2:
    st.subheader("2. Upload Update Files")
    update_files = st.file_uploader("Choose HR Update Excel Files", type=['xlsx'], accept_multiple_files=True, key="update_uploader")

# Initialize session state for processed files and timestamp
if 'processed_outputs' not in st.session_state:
    st.session_state.processed_outputs = None
if 'process_timestamp' not in st.session_state:
    st.session_state.process_timestamp = None

if st.button("Process Files", type="primary"):
    if not master_file:
        st.error("Please upload a Master File.")
    elif not update_files:
        st.error("Please upload at least one Update File.")
    else:
        with st.spinner("Processing files..."):
            try:
                # Run processing
                output_files = process_files(master_file, update_files)
                
                # Store in session state
                st.session_state.processed_outputs = output_files
                st.session_state.process_timestamp = datetime.now().strftime("%Y%m%d_%H%M")
                
                st.success("Processing Complete!")
                    
            except Exception as e:
                st.error(f"An error occurred during processing: {e}")
                import traceback
                traceback.print_exc()

# Display download buttons if data is available in session state
if st.session_state.processed_outputs:
    st.divider()
    st.subheader(f"Download Results (Processed at {st.session_state.process_timestamp})")
    
    outputs = st.session_state.processed_outputs
    ts = st.session_state.process_timestamp
    
    d_col1, d_col2, d_col3 = st.columns(3)
    
    with d_col1:
        st.download_button(
            label="Download SAP File",
            data=outputs['sap_file.xlsx'],
            file_name=f"sap_file_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_sap"
        )
    
    with d_col2:
        st.download_button(
            label="Download Resigned Batch",
            data=outputs['resigned_batch_file.xlsx'],
            file_name=f"resigned_batch_file_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_resigned"
        )
    
    with d_col3:
        st.download_button(
            label="Download New Master",
            data=outputs['new_master_file.xlsx'],
            file_name=f"new_master_file_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_master"
        )
