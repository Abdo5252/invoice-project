import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime
from invoice_processor import process_invoices
from excel_utils import create_output_excel

st.set_page_config(page_title="Invoice Data Extractor",
                   page_icon="📊",
                   layout="wide")

st.title("Invoice Data Extractor")
st.markdown("""
This application extracts specific invoice data from a messy Excel file with multiple sheets 
and organizes it into two structured sheets: a Headers sheet and an Items sheet.
""")

st.subheader("Upload Files")

# File uploads
uploaded_messy_file = st.file_uploader(
    "Upload your messy Excel file with multiple sheets", type=["xlsx", "xls"])

if uploaded_messy_file:
    st.success("File uploaded successfully!")

    # Process button
    if st.button("Process Invoices"):
        with st.spinner("Processing invoices..."):
            try:
                # Process the uploaded file
                processed_data = process_invoices(uploaded_messy_file)

                if not processed_data:
                    st.error(
                        "No invoice data could be extracted from the provided file."
                    )
                else:
                    # Create output file with two sheets (Headers and Items)
                    output_excel_bytes = create_output_excel(processed_data)

                    # Display success message with the number of invoices processed
                    st.success(
                        f"Successfully processed {len(processed_data)} invoices!"
                    )

                    # Prepare download button
                    st.download_button(
                        label="Download Processed Excel File",
                        data=output_excel_bytes,
                        file_name="processed_invoices.xlsx",
                        mime=
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    # Create dataframes for display
                    headers_data = []
                    items_data = []

                    for invoice in processed_data:
                        # Add header row
                        # Set exchange rate based on currency: 52 for USD, 60 for EUR, 0 for others
                        currency = invoice.get('currency', '')
                        exchange_rate = "0"  # Default exchange rate
                        
                        if currency == 'USD':
                            exchange_rate = "52"
                        elif currency == 'EUR':
                            exchange_rate = "60"  # Exchange rate for Euro
                        
                        # Get current date for document date
                        current_date = datetime.now().strftime("%m/%d/%Y")
                        
                        headers_data.append({
                            "Document Type":
                            "I",
                            "Document Number":
                            invoice.get('invoice_number', ''),
                            "Document Date":
                            current_date,
                            "Customer Code":
                            invoice.get('customer_code', ''),
                            "Currency Code":
                            invoice.get('currency', ''),
                            "Exchange Rate":
                            exchange_rate,
                            "Extra Discount":
                            "0",
                            "Activity Code":
                            "",
                            "Total Amount":
                            invoice.get('total_amount', 0)
                        })

                        # Add product rows
                        if 'products' in invoice and invoice['products']:
                            for product in invoice['products']:
                                items_data.append({
                                    "Document Number":
                                    invoice.get('invoice_number', ''),
                                    "Internal Code":
                                    "1",
                                    "Description":
                                    product.get('description', ''),
                                    "Unit Type":
                                    "",
                                    "Quantity":
                                    product.get('quantity', ''),
                                    "Unit Price":
                                    product.get('unit_price', ''),
                                    "Discount Amount":
                                    "0",
                                    "Value Difference":
                                    "0",
                                    "Item Discount":
                                    "0"
                                })

                    # Display extracted data overview
                    st.subheader("Extracted Data Preview")

                    # Display Headers tab
                    tab1, tab2 = st.tabs(["Header", "Items"])

                    with tab1:
                        st.write("### Invoice Headers")
                        if headers_data:
                            st.dataframe(pd.DataFrame(headers_data))
                        else:
                            st.write("No header data extracted.")

                    with tab2:
                        st.write("### Invoice Line Items")
                        if items_data:
                            st.dataframe(pd.DataFrame(items_data))
                        else:
                            st.write("No product data extracted.")

                    # Also show raw data in expandable sections
                    st.subheader("Raw Extracted Data")
                    for i, invoice in enumerate(processed_data):
                        with st.expander(
                                f"Invoice {i+1}: {invoice.get('invoice_number', 'Unknown')}"
                        ):
                            st.write(
                                f"**Invoice Number:** {invoice.get('invoice_number', 'Not found')}"
                            )
                            st.write(
                                f"**Document Date:** {datetime.now().strftime('%m/%d/%Y')}"
                            )
                            st.write(
                                f"**Customer Code:** {invoice.get('customer_code', 'Not found')}"
                            )
                            st.write(
                                f"**Currency:** {invoice.get('currency', 'Not found')}"
                            )
                            st.write(
                                f"**Total Amount:** {invoice.get('total_amount', 0)} {invoice.get('currency', '')}"
                            )

                            if 'products' in invoice and invoice['products']:
                                st.write("**Products:**")
                                st.dataframe(pd.DataFrame(invoice['products']))
                            else:
                                st.write(
                                    "No product data found in this invoice.")

            except Exception as e:
                st.error(f"An error occurred during processing: {str(e)}")
else:
    st.info(
        "Please upload your messy Excel file with invoice data to proceed.")

# Add information section
with st.expander("About this app"):
    st.markdown("""
    ### How it works
    
    1. Upload your messy Excel file containing multiple sheets with invoice data
    2. Click "Process Invoices" to extract and format the data
    3. Download the resulting Excel file with two sheets:
       - **Header sheet**: Contains all invoice headers with Document Type, Number, Date, etc.
       - **Items sheet**: Contains all product items from all invoices linked by Document Number
    
    ### Data Extraction Details
    
    The app extracts information using exact patterns from your invoice files:
    
    - **Invoice Number:** Found after "INVOICE N:" keywords with format SIxxxxx
    - **Document Date:** Extracted in MM/DD/YYYY format (example: 4/16/2025)
    - **Customer Code:** Found near "partner code:" with format Cxxxx
    - **Currency:** Standardized as EGP or USD based on content
    - **Product Details:** Extracted from tables with both Arabic and English headers:
        - Description (التسمية or Description)
        - Quantity (الكمية or Quantity)
        - Unit Price (سعر الوحدة or Unit price)
    
    ### Field Handling
    
    - Numeric fields use 0 as default when not found (Exchange Rate, Extra Discount)
    - Text fields are left empty when not found (Activity Code)
    - All data is output in two sheets: "Header" sheet and "Items" sheet
    - Document Numbers link Header entries to their corresponding Item rows
    
    ### Made by Abdelrhman Adel
    """)
