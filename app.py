import streamlit as st
import pandas as pd
import io
import os
from invoice_processor import process_invoices
from excel_utils import create_output_excel

st.set_page_config(
    page_title="Invoice Data Extractor",
    page_icon="📊",
    layout="wide"
)

st.title("Invoice Data Extractor")
st.markdown("""
This application extracts specific invoice data from a messy Excel file with multiple sheets 
and transfers it to a structured template format.
""")

st.subheader("Upload Files")

# File uploads
col1, col2 = st.columns(2)
with col1:
    uploaded_messy_file = st.file_uploader("Upload your messy Excel file with multiple sheets", type=["xlsx", "xls"])
with col2:
    uploaded_template_file = st.file_uploader("Upload your template Excel file", type=["xlsx", "xls"])

if uploaded_messy_file and uploaded_template_file:
    st.success("Files uploaded successfully!")
    
    # Process button
    if st.button("Process Invoices"):
        with st.spinner("Processing invoices..."):
            try:
                # Process the uploaded files
                processed_data = process_invoices(uploaded_messy_file)
                
                if not processed_data:
                    st.error("No invoice data could be extracted from the provided file.")
                else:
                    # Create output file based on template
                    output_excel_bytes = create_output_excel(processed_data, uploaded_template_file)
                    
                    # Display success message with the number of invoices processed
                    st.success(f"Successfully processed {len(processed_data)} invoices!")
                    
                    # Prepare download button
                    st.download_button(
                        label="Download Processed Excel File",
                        data=output_excel_bytes,
                        file_name="processed_invoices.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    # Display extracted data overview
                    st.subheader("Extracted Data Overview")
                    
                    for i, invoice in enumerate(processed_data):
                        with st.expander(f"Invoice {i+1}: {invoice.get('invoice_number', 'Unknown')}"):
                            st.write(f"**Invoice Number:** {invoice.get('invoice_number', 'Not found')}")
                            st.write(f"**Customer Code:** {invoice.get('customer_code', 'Not found')}")
                            st.write(f"**Currency:** {invoice.get('currency', 'Not found')}")
                            
                            if 'products' in invoice and invoice['products']:
                                st.write("**Products:**")
                                st.dataframe(pd.DataFrame(invoice['products']))
                            else:
                                st.write("No product data found in this invoice.")
            
            except Exception as e:
                st.error(f"An error occurred during processing: {str(e)}")
else:
    st.info("Please upload both the messy Excel file and the template file to proceed.")

# Add information section
with st.expander("About this app"):
    st.markdown("""
    ### How it works
    
    1. Upload your messy Excel file containing multiple sheets with invoice data
    2. Upload your template Excel file with the desired output format
    3. Click "Process Invoices" to extract and format the data
    4. Download the resulting Excel file
    
    ### Data Extraction Details
    
    The app looks for the following information:
    
    - **Invoice Number:** Found near keywords like "INVOICE N:" or similar
    - **Customer Code:** Found near keywords like "partner code:" or similar
    - **Currency:** Detected if available near price fields
    - **Product Details:** Extracted from tables with headers in Arabic or English:
        - Description (ÇáÊÓãíÉ or Description)
        - Quantity (ÇáßãíÉ or Quantity)
        - Unit Price (ÓÚÑ ÇáæÍÏÉ or Unit price)
    """)
