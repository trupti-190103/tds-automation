import streamlit as st
from io import BytesIO

# IMPORT PROCESSORS
from processors.purchase_tds import process_purchase_tds
from processors.contractor_tds import process_contractor_tds
from processors.contractor_1024_tds import process_contractor_1024_tds
from processors.professional_tds import process_professional_tds

# PAGE CONFIG
st.set_page_config(
    page_title="TDS Automation",
    layout="wide"
)

# TITLE
st.title("TDS Automation System")

st.markdown("---")

# DROPDOWN
tds_type = st.selectbox(
    "Select TDS Type",
    [
        "TDS on Purchase of Goods 1031 Code",
        "TDS on Contractor 1023 Code",
        "TDS on Contractor ( Other Than Individual)1024 Code",
        "TDS on Professional Fee. Sec194J(B) 1027 Code"
    ]
)

# FILE UPLOAD
uploaded_file = st.file_uploader(
    "Upload Tally Excel File",
    type=['xlsx']
)


# PROCESS BUTTON
if uploaded_file is not None:

    if st.button("Process TDS"):

        try:

            # =========================
            # PURCHASE TDS
            # =========================

            if tds_type == "TDS on Purchase of Goods 1031":

                wb = process_purchase_tds(uploaded_file)

                output_filename = (
                    "Purchase_TDS_Processed.xlsx"
                )

            # =========================
            # CONTRACTOR TDS
            # =========================

            elif tds_type == "TDS on Contractor 1023":

                wb = process_contractor_tds(uploaded_file)

                output_filename = (
                    "Contractor_TDS_Processed.xlsx"
                )


            elif tds_type == "TDS on Contractor ( Other Than Individual)1024 Code":

                wb = process_contractor_1024_tds(
                    uploaded_file
                )

                output_filename = (
                    "Contractor_1024_TDS.xlsx"
                )

            elif tds_type == "TDS on Professional Fee. Sec194J(B) 1027 Code":
                wb = process_professional_tds(
                    uploaded_file
                )
                
                output_filename = (
                    "Professional_TDS.xlsx"
                )

                

                

            # =========================
            # DOWNLOAD
            # =========================

            output = BytesIO()

            wb.save(output)

            output.seek(0)


            st.success("TDS Processed Successfully")

            st.download_button(
                label="Download Processed Excel",
                data=output,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:

            st.error(f"Error: {e}")