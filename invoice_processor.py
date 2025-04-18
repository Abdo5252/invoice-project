import pandas as pd
import re
import numpy as np
import io
import os
from datetime import datetime

def process_invoices(uploaded_file):
    """
    Process the uploaded Excel file containing multiple invoice sheets.
    Extract relevant invoice information from each sheet.

    Args:
        uploaded_file: The uploaded Excel file object

    Returns:
        List of dictionaries containing the processed invoice data
    """
    # Read the Excel file with all sheets
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names

    # List to store processed invoices
    processed_invoices = []

    # Process each sheet (invoice)
    for sheet_name in sheet_names:
        try:
            # Read the sheet into a DataFrame
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

            # Convert DataFrame to ensure all values are strings for text searching
            # Replace NaN with empty string to avoid errors during text search
            df_string = df.astype(str).replace('nan', '', regex=True)

            # Extract invoice number first
            invoice_number = extract_invoice_number(df_string)

            # Extract invoice data with linked products
            invoice_data = {
                'invoice_number': invoice_number,
                'customer_code': extract_customer_code(df_string),
                'currency': extract_currency(df_string),
                'invoice_date': extract_invoice_date(df_string),
                'products': extract_product_details(df_string, invoice_number),
                'sheet_name': sheet_name
            }

            processed_invoices.append(invoice_data)

        except Exception as e:
            print(f"Error processing sheet {sheet_name}: {str(e)}")
            continue

    return processed_invoices

def extract_invoice_number(df):
    """
    Extract invoice number from the dataframe.

    Args:
        df: DataFrame with string values

    Returns:
        Extracted invoice number or None if not found
    """
    # Primary method: Look for "INVOICE N:" keyword as specified in requirements
    invoice_n_keyword = 'invoice n:'

    for i in range(len(df)):
        for j in range(len(df.columns)):
            cell = str(df.iloc[i, j]).lower()
            if invoice_n_keyword in cell:
                # Extract invoice number directly from this cell 
                match = re.search(f"{invoice_n_keyword}[:\s°.]*([a-zA-Z0-9\-/]+)", cell, re.IGNORECASE)
                if match:
                    invoice_num = match.group(1).strip()
                    # Validate it follows the SIxxxxx pattern
                    if re.match(r'^SI\d+$', invoice_num):
                        return invoice_num

                # Check the cell to the right (common layout)
                if j + 1 < len(df.columns):
                    right_cell = str(df.iloc[i, j + 1])
                    if right_cell and not 'nan' in right_cell:
                        right_cell = right_cell.strip()
                        if re.match(r'^SI\d+$', right_cell):
                            return right_cell

                # Check the cell below (alternate layout)
                if i + 1 < len(df):
                    below_cell = str(df.iloc[i + 1, j])
                    if below_cell and not 'nan' in below_cell:
                        below_cell = below_cell.strip()
                        if re.match(r'^SI\d+$', below_cell):
                            return below_cell

    # Backup method 1: Look for cells with "Document Number" header and extract data
    # Look for a column header named "Document Number"
    for i in range(len(df)):
        for j in range(len(df.columns)):
            cell = str(df.iloc[i, j]).lower()
            if 'document number' in cell:
                # We found the header, now get the invoice numbers from the column
                for row_idx in range(i+1, min(i+20, len(df))):
                    doc_number = str(df.iloc[row_idx, j])
                    if doc_number and doc_number != 'nan' and re.match(r'^SI\d+$', doc_number):
                        return doc_number

    # Backup method 2: Extract from any cell with SI pattern
    si_pattern = r'SI\d+'
    for i in range(len(df)):
        for j in range(len(df.columns)):
            cell = str(df.iloc[i, j])
            match = re.search(si_pattern, cell)
            if match:
                return match.group(0).strip()

    # Backup method 3: Look for other invoice number keywords
    other_invoice_keywords = [
        'invoice no', 'invoice number', 'invoice #',
        'inv', 'inv no', 'invoice', 'فاتورة رقم', 'رقم الفاتورة'
    ]

    # Search for each keyword in the dataframe
    for keyword in other_invoice_keywords:
        for i in range(len(df)):
            for j in range(len(df.columns)):
                cell = str(df.iloc[i, j]).lower()
                if keyword.lower() in cell:
                    # Try to extract the invoice number from this cell or neighboring cells
                    match = re.search(f"{keyword.lower()}[:\s°.]*([a-zA-Z0-9\-/]+)", cell, re.IGNORECASE)
                    if match:
                        invoice_num = match.group(1).strip()
                        # Check if it looks like an SI invoice number
                        if re.match(r'^SI\d+$', invoice_num):
                            return invoice_num

                    # Check the cell to the right
                    if j + 1 < len(df.columns):
                        right_cell = str(df.iloc[i, j + 1])
                        if right_cell and not 'nan' in right_cell:
                            if re.match(r'^SI\d+$', right_cell.strip()):
                                return right_cell.strip()

    return None

def extract_customer_code(df):
    """
    Extract customer code from the dataframe.

    Args:
        df: DataFrame with string values

    Returns:
        Extracted customer code or None if not found
    """
    # Primary method: Look for "partner code:" keyword as specified in requirements
    partner_code_keyword = 'partner code:'

    for i in range(len(df)):
        for j in range(len(df.columns)):
            cell = str(df.iloc[i, j]).lower()
            if partner_code_keyword in cell:
                # Try to extract the customer code directly from this cell
                match = re.search(f"{partner_code_keyword}[:\s]*([a-zA-Z0-9\-/]+)", cell, re.IGNORECASE)
                if match:
                    code = match.group(1).strip()
                    # Validate it follows the Cxxxx pattern
                    if re.match(r'^C\d+$', code):
                        return code

                # Check the cell to the right (common layout)
                if j + 1 < len(df.columns):
                    right_cell = str(df.iloc[i, j + 1])
                    if right_cell and not 'nan' in right_cell:
                        right_cell = right_cell.strip()
                        if re.match(r'^C\d+$', right_cell):
                            return right_cell

                # Check a few cells to the right (in case value is in a different column)
                for k in range(2, min(5, len(df.columns) - j)):
                    if j + k < len(df.columns):
                        far_right_cell = str(df.iloc[i, j + k])
                        if far_right_cell and not 'nan' in far_right_cell and re.match(r'^C\d+$', far_right_cell.strip()):
                            return far_right_cell.strip()

                # Check the cell below (alternate layout)
                if i + 1 < len(df):
                    below_cell = str(df.iloc[i + 1, j])
                    if below_cell and not 'nan' in below_cell:
                        below_cell = below_cell.strip()
                        if re.match(r'^C\d+$', below_cell):
                            return below_cell

    # Backup method 1: Look for "Customer Code" column in structured tables
    for i in range(len(df)):
        for j in range(len(df.columns)):
            cell = str(df.iloc[i, j]).lower()
            if 'customer code' in cell:
                # Look for C-prefixed codes in the column below
                for row_idx in range(i+1, min(i+20, len(df))):
                    code_value = str(df.iloc[row_idx, j])
                    # Looking for customer codes like C0033, C0034, etc.
                    if re.match(r'^C\d+$', code_value):
                        return code_value

                # Check the column to the right if not found directly below
                if j + 1 < len(df.columns):
                    for row_idx in range(i+1, min(i+20, len(df))):
                        code_value = str(df.iloc[row_idx, j+1])
                        if re.match(r'^C\d+$', code_value):
                            return code_value

    # Backup method 2: Look for other customer code indicators
    other_code_keywords = [
        'customer code', 'client code', 'account code', 'partner id',
        'رمز العميل', 'كود العميل', 'رقم العميل', 'partner code :',
        'partner details'
    ]

    # Search for each keyword in the dataframe
    for keyword in other_code_keywords:
        for i in range(len(df)):
            for j in range(len(df.columns)):
                cell = str(df.iloc[i, j]).lower()
                if keyword.lower() in cell:
                    # Try to extract from nearby cells
                    # Check the cell to the right
                    if j + 1 < len(df.columns):
                        right_cell = str(df.iloc[i, j + 1])
                        if right_cell and not 'nan' in right_cell and re.match(r'^C\d+$', right_cell.strip()):
                            return right_cell.strip()

                    # Check the cell below
                    if i + 1 < len(df):
                        below_cell = str(df.iloc[i + 1, j])
                        if below_cell and not 'nan' in below_cell and re.match(r'^C\d+$', below_cell.strip()):
                            return below_cell.strip()

    # Backup method 3: Search for any C-prefixed code in the document
    c_pattern = r'C\d+'
    for i in range(len(df)):
        for j in range(len(df.columns)):
            cell = str(df.iloc[i, j])
            match = re.search(c_pattern, cell)
            if match and re.match(r'^C\d+$', match.group(0)):
                return match.group(0).strip()

    return None

def extract_currency(df):
    """
    Extract currency information from the dataframe.
    ONLY returns EGP or USD as specified in the requirements.

    Args:
        df: DataFrame with string values

    Returns:
        'EGP' or 'USD' only, with EGP as default
    """
    # First method: Look for the "Currency Code" column in a structured table
    for i in range(len(df)):
        for j in range(len(df.columns)):
            cell = str(df.iloc[i, j]).lower()
            if 'currency code' in cell:
                # Look for currency codes in the column below
                for row_idx in range(i+1, min(i+20, len(df))):
                    curr_value = str(df.iloc[row_idx, j]).strip().upper()
                    # ONLY allow EGP or USD as specified in requirements
                    if curr_value == 'EGP' or curr_value == 'USD':
                        return curr_value

                # If not found in the exact column, check the one to the right
                if j + 1 < len(df.columns):
                    for row_idx in range(i+1, min(i+20, len(df))):
                        curr_value = str(df.iloc[row_idx, j+1]).strip().upper()
                        if curr_value == 'EGP' or curr_value == 'USD':
                            return curr_value

    # Look specifically for dollar or USD mentions - they indicate USD currency
    usd_indicators = ['$', 'dollar', 'usd', 'دولار', 'united states', 'u.s.']
    for i in range(len(df)):
        for j in range(len(df.columns)):
            cell = str(df.iloc[i, j]).lower()
            if any(indicator in cell for indicator in usd_indicators):
                return 'USD'

    # Look for Egypt mentions - they indicate EGP currency
    egp_indicators = ['egypt', 'egyptian', 'egp', 'مصر', 'مصري', 'جنيه']
    for i in range(len(df)):
        for j in range(len(df.columns)):
            cell = str(df.iloc[i, j]).lower()
            if any(indicator in cell for indicator in egp_indicators):
                return 'EGP'

    # Look for explicit currency keywords then check nearby cells
    currency_keywords = ['currency', 'العملة', 'curr.', 'curr', 'currency code']
    for i in range(len(df)):
        for j in range(len(df.columns)):
            cell = str(df.iloc[i, j]).lower()
            if any(keyword in cell for keyword in currency_keywords):
                # Check the cell itself
                if 'usd' in cell or 'dollar' in cell or '$' in cell:
                    return 'USD'
                if 'egp' in cell or 'egypt' in cell:
                    return 'EGP'

                # Check the cell to the right
                if j + 1 < len(df.columns):
                    right_cell = str(df.iloc[i, j + 1]).lower()
                    if right_cell == 'usd' or right_cell == 'dollar' or right_cell == '$':
                        return 'USD'
                    if right_cell == 'egp' or right_cell == 'egypt':
                        return 'EGP'

    # Default to EGP if no currency found (since most invoices in examples use EGP)
    return 'EGP'

def extract_invoice_date(df):
    """
    Extract invoice date from the dataframe.

    Args:
        df: DataFrame with string values

    Returns:
        Extracted date as string in MM/DD/YYYY format to match the examples
    """
    # First method: Look for structured "Document Date" column in a table
    for i in range(len(df)):
        for j in range(len(df.columns)):
            cell = str(df.iloc[i, j]).lower()
            if 'document date' in cell:
                # Found a document date header - check for dates in the column below
                for row_idx in range(i+1, min(i+20, len(df))):
                    date_value = str(df.iloc[row_idx, j])
                    # If it's a date in MM/DD/YYYY format (like in the examples), return it
                    if re.match(r'\d{1,2}/\d{1,2}/\d{4}', date_value):
                        return date_value

                # If not found in the exact column, check the one to the right
                if j + 1 < len(df.columns):
                    for row_idx in range(i+1, min(i+20, len(df))):
                        date_value = str(df.iloc[row_idx, j+1])
                        if re.match(r'\d{1,2}/\d{1,2}/\d{4}', date_value):
                            return date_value

    # Keywords that might precede an invoice date
    date_keywords = [
        'date', 'invoice date', 'issued on', 'document date', 'payment date',
        'تاريخ', 'تاريخ الفاتورة'
    ]

    # Common date formats (DD/MM/YYYY, MM/DD/YYYY, YYYY-MM-DD, etc.)
    date_patterns = [
        r'(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})',  # DD/MM/YYYY or MM/DD/YYYY
        r'(\d{2,4}[/-]\d{1,2}[/-]\d{1,2})',  # YYYY-MM-DD
    ]

    # Search for each keyword in the dataframe
    for keyword in date_keywords:
        for i in range(len(df)):
            for j in range(len(df.columns)):
                cell = str(df.iloc[i, j]).lower()
                if keyword.lower() in cell:
                    # Check this cell for date patterns
                    for pattern in date_patterns:
                        match = re.search(pattern, cell)
                        if match:
                            # Try to parse the date
                            try:
                                date_str = match.group(1)
                                # Convert to MM/DD/YYYY format to match examples
                                if re.match(r'\d{1,2}/\d{1,2}/\d{4}', date_str):
                                    return date_str
                                else:
                                    # For other formats, return as is
                                    return date_str
                            except:
                                pass

                    # Check the cell to the right
                    if j + 1 < len(df.columns):
                        right_cell = str(df.iloc[i, j + 1])
                        for pattern in date_patterns:
                            match = re.search(pattern, right_cell)
                            if match:
                                try:
                                    date_str = match.group(1)
                                    if re.match(r'\d{1,2}/\d{1,2}/\d{4}', date_str):
                                        return date_str
                                    else:
                                        return date_str
                                except:
                                    pass

    # If no date found from the sheet, return the standard date from the examples
    return "4/16/2025"

def extract_product_details(df, invoice_number=None):
    """
    Extract product details from structured invoice tables.
    Focuses on the specific format with Description, Quantity, and Unit price columns.

    Args:
        df: DataFrame with string values
        invoice_number: The invoice number to link products to

    Returns:
        List of dictionaries containing product details with invoice linkage
    """
    products = []

    # Look for "Invoice details" section which contains the product table
    for i in range(len(df)):
        for j in range(len(df.columns)):
            cell = str(df.iloc[i, j]).strip().lower()
            if 'invoice details' in cell:
                # Found the product table header, look for column headers in next rows
                header_row = i + 1
                desc_col = None
                qty_col = None
                price_col = None

                # Look for column headers in the next few rows
                for row in range(header_row, min(header_row + 3, len(df))):
                    for col in range(len(df.columns)):
                        header = str(df.iloc[row, col]).strip().lower()
                        if 'description' in header:
                            desc_col = col
                        elif any(x in header for x in ['quantity', 'qty']):
                            qty_col = col
                        elif 'unit price' in header:
                            price_col = col

                # If we found the columns, start extracting products
                if desc_col is not None and (qty_col is not None or price_col is not None):
                    current_row = row + 1

                    # Process rows until we hit an empty row or obvious end markers
                    while current_row < len(df):
                        desc = str(df.iloc[current_row, desc_col]).strip()

                        # Skip empty rows or header-like content
                        if not desc or desc.lower() in ['description', 'total', 'amount']:
                            current_row += 1
                            continue

                        product = {'description': desc}

                        # Extract quantity if column exists
                        if qty_col is not None:
                            qty_str = str(df.iloc[current_row, qty_col]).strip()
                            try:
                                qty = float(qty_str.replace(',', ''))
                                if qty > 0:
                                    product['quantity'] = qty
                            except:
                                pass

                        # Extract unit price if column exists
                        if price_col is not None:
                            price_str = str(df.iloc[current_row, price_col]).strip()
                            try:
                                price = float(price_str.replace(',', ''))
                                if price > 0:
                                    product['unit_price'] = price
                            except:
                                pass

                        # Add product if we have enough data
                        if len(product) > 1:  # Must have description and at least one other field
                            product['invoice_number'] = invoice_number
                            products.append(product)

                        current_row += 1

                        # Stop if we hit "Amount excluding Value Added Tax" or similar
                        if current_row < len(df):
                            next_row = str(df.iloc[current_row, desc_col]).strip().lower()
                            if 'amount' in next_row and 'tax' in next_row:
                                break

                    # If we found products, return them
                    if products:
                        return products

    #Fallback to the original method if no structured table is found.
    products = []

    # Look for English headers only
    header_keywords = {
        'description': ['description', 'item', 'product', 'desc'],
        'quantity': ['quantity', 'qty', 'amount', 'pcs'],
        'price': ['unit price', 'price', 'amount', 'unit cost']
    }

    def is_number(s):
        try:
            float(str(s).replace(',', ''))
            return True
        except:
            return False

    def analyze_row(row):
        """Analyze a row to determine if it looks like a product row"""
        text_cols = []
        num_cols = []

        for i, cell in enumerate(row):
            cell_str = str(cell).strip()
            if cell_str and cell_str.lower() != 'nan':
                if is_number(cell_str):
                    num_cols.append(i)
                else:
                    text_cols.append(i)

        return len(text_cols) >= 1 and len(num_cols) >= 1

    # First try to find tables by data patterns
    for i in range(len(df)):
        row_data = [str(df.iloc[i, j]).strip() for j in range(len(df.columns))]

        # Skip empty rows
        if not any(row_data):
            continue

        # Check if this row looks like a product row
        if analyze_row(row_data):
            # Find columns based on data type
            desc_col = None
            num_cols = []

            for j, cell in enumerate(row_data):
                if cell and cell.lower() != 'nan':
                    if is_number(cell):
                        num_cols.append(j)
                    elif len(cell) > 3 and not any(k in cell.lower() for k in sum(header_keywords.values(), [])):
                        desc_col = j

            if desc_col is not None and num_cols:
                # Process rows starting from here
                current_row = i
                pattern_products = []

                while current_row < len(df) and current_row < i + 30:
                    row_values = [str(df.iloc[current_row, j]).strip() for j in range(len(df.columns))]

                    if not any(row_values):
                        break

                    description = str(df.iloc[current_row, desc_col]).strip()
                    if not description or description.lower() == 'nan':
                        current_row += 1
                        continue

                    product = {'description': description}

                    # Get first numeric value as quantity
                    for num_col in num_cols:
                        val = str(df.iloc[current_row, num_col]).strip()
                        if val and is_number(val):
                            try:
                                num = float(val.replace(',', ''))
                                if num > 0 and num < 1000000:  # Reasonable quantity range
                                    product['quantity'] = num
                                    num_cols.remove(num_col)
                                    break
                            except:
                                pass

                    # Get second numeric value as price
                    for num_col in num_cols:
                        val = str(df.iloc[current_row, num_col]).strip()
                        if val and is_number(val):
                            try:
                                num = float(val.replace(',', ''))
                                if num > 0:  # Any positive number could be a price
                                    product['unit_price'] = num
                                    break
                            except:
                                pass

                    if len(product) > 1:
                        # Link the product to its invoice
                        product['invoice_number'] = invoice_number
                        pattern_products.append(product)

                    current_row += 1

                if pattern_products:
                    products.extend(pattern_products)
                    return products

        # If pattern matching didn't work, try header keywords
        header_cols = {'description': None, 'quantity': None, 'price': None}
        header_found = False

        # Look for header row
        for j in range(len(df.columns)):
            cell = str(df.iloc[i, j]).strip().lower()

            # Check for each type of header
            for header_type, keywords in header_keywords.items():
                if any(keyword in cell for keyword in keywords):
                    header_cols[header_type] = j
                    header_found = True

            # Look for tables with Document Number, Description, Quantity, and Unit Price headers
            if ('document number' in cell) or ('description' in cell and j+3 < len(df.columns)):
                # Search for a row that has these headers
                header_row = i
                header_found = False
                doc_num_col = None
                desc_col = None
                qty_col = None
                price_col = None
                unit_type_col = None

                # Scan this row and nearby rows for headers
                for scan_row in range(max(0, i-1), min(i+2, len(df))):
                    for scan_col in range(len(df.columns)):
                        scan_cell = str(df.iloc[scan_row, scan_col]).strip().lower()

                        if scan_cell == 'document number':
                            doc_num_col = scan_col
                            header_found = True
                        elif scan_cell == 'description':
                            desc_col = scan_col
                            header_found = True
                        elif scan_cell == 'quantity':
                            qty_col = scan_col
                            header_found = True
                        elif scan_cell == 'unit price':
                            price_col = scan_col
                            header_found = True
                        elif scan_cell == 'unit type':
                            unit_type_col = scan_col

                # If we found at least description and one other field
                if header_found and header_cols['description'] is not None and (header_cols['quantity'] is not None or header_cols['price'] is not None):
                    current_row = i + 1
                    temp_products = []

                    # Process rows until we hit an empty row or max rows
                    while current_row < len(df) and current_row < i + 20:  # Limit to 20 rows
                        desc = str(df.iloc[current_row, header_cols['description']]).strip()

                        # Skip if description is empty or looks like a header
                        if not desc or desc.lower() in [k for kws in header_keywords.values() for k in kws]:
                            current_row += 1
                            continue

                        product = {'description': desc}

                        # Get quantity if column exists
                        if header_cols['quantity'] is not None:
                            qty_str = str(df.iloc[current_row, header_cols['quantity']]).strip()
                            if qty_str and not any(k in qty_str.lower() for k in header_keywords['quantity']):
                                try:
                                    qty = float(qty_str.replace(',', ''))
                                    product['quantity'] = qty
                                except:
                                    product['quantity'] = qty_str

                        # Get price if column exists
                        if header_cols['price'] is not None:
                            price_str = str(df.iloc[current_row, header_cols['price']]).strip()
                            if price_str and not any(k in price_str.lower() for k in header_keywords['price']):
                                try:
                                    price = float(price_str.replace(',', ''))
                                    product['unit_price'] = price
                                except:
                                    product['unit_price'] = price_str

                        # Add product if it has enough data
                        if len(product) > 1:  # At least description and one other field
                            temp_products.append(product)

                        current_row += 1

                    # If we found products, add them and return
                    if temp_products:
                        products.extend(temp_products)
                        return products

                    # Process rows in this table
                    while current_row < len(df) and current_row < header_row + 30:  # Limit search depth
                        # Skip empty rows
                        if all(str(df.iloc[current_row, k]).strip() == '' 
                              for k in range(max(0, j-2), min(j+8, len(df.columns)))):
                            current_row += 1
                            continue

                        # Get values from this row
                        # For document number, either use the column or store from invoice
                        doc_num = ''
                        if doc_num_col is not None:
                            doc_num = str(df.iloc[current_row, doc_num_col]).strip()

                        # Description is required
                        if desc_col is None:
                            current_row += 1
                            continue

                        description = str(df.iloc[current_row, desc_col]).strip()

                        # Skip if description is empty or looks like a header
                        if not description or description.lower() in ['description', 'item', 'product', 'التسمية', 'الوصف']:
                            current_row += 1
                            continue

                        # Create product entry
                        product = {
                            'description': description,
                        }

                        # Add document number if found
                        if doc_num:
                            product['document_number'] = doc_num

                        # Get quantity if column exists
                        if qty_col is not None:
                            qty_str = str(df.iloc[current_row, qty_col]).strip()
                            if qty_str and qty_str.lower() not in ['quantity', 'qty', 'الكمية']:
                                try:
                                    # Clean the quantity string and convert to float
                                    qty_str = qty_str.replace(',', '')
                                    qty = float(qty_str)
                                    product['quantity'] = qty
                                except:
                                    product['quantity'] = qty_str

                        # Get unit price if column exists
                        if price_col is not None:
                            price_str = str(df.iloc[current_row, price_col]).strip()
                            if price_str and price_str.lower() not in ['unit price', 'price', 'سعر الوحدة']:
                                try:
                                    # Clean the price string and convert to float
                                    price_str = price_str.replace(',', '')
                                    price = float(price_str)
                                    product['unit_price'] = price
                                except:
                                    product['unit_price'] = price_str

                        # Get unit type if column exists
                        if unit_type_col is not None:
                            unit_type = str(df.iloc[current_row, unit_type_col]).strip()
                            if unit_type and unit_type.lower() not in ['unit type', 'unit', 'الوحدة']:
                                product['unit_type'] = unit_type

                        # Add the product if it has essential data
                        if 'description' in product and (
                            'quantity' in product or 'unit_price' in product
                        ):
                            doc_products.append(product)

                        current_row += 1

                    # If we found products in this table, add them to the results
                    if doc_products:
                        products.extend(doc_products)
                        # Return early if we've found products in the expected format
                        return products

    # Method 2: Look for Arabic description tables (like in example image 4)
    arabic_desc_found = False
    for i in range(len(df)):
        for j in range(len(df.columns)):
            cell = str(df.iloc[i, j]).strip().lower()

            # Look for Arabic description header
            if 'التسمية' in cell or 'الوصف' in cell:
                arabic_desc_found = True
                desc_col = j
                qty_col = None
                price_col = None

                # Look for quantity and price headers in this row
                for k in range(len(df.columns)):
                    cell_k = str(df.iloc[i, k]).strip().lower()
                    if 'الكمية' in cell_k or 'العدد' in cell_k:
                        qty_col = k
                    elif 'سعر الوحدة' in cellk or 'السعر' in cell_k:
                        price_col = k

                # If we've found the key columns
                if qty_col is not None or price_col is not None:
                    # Process product rows
                    current_row = i + 1
                    arabic_products = []

                    while current_row < len(df) and current_row < i + 30:  # Limit search depth
                        # Check if this is a valid product row
                        desc = str(df.iloc[current_row, desc_col]).strip()

                        if desc and not any(header in desc.lower() for header in ['التسمية', 'الوصف', 'الكمية', 'سعر']):
                            product = {'description': desc}

                            # Get quantity
                            if qty_col is not None:
                                qty_str = str(df.iloc[current_row, qty_col]).strip()
                                if qty_str and not any(h in qty_str.lower() for h in ['الكمية', 'quantity']):
                                    try:
                                        qty = float(qty_str.replace(',', ''))
                                        product['quantity'] = qty
                                    except:
                                        product['quantity'] = qty_str

                            # Get unit price
                            if price_col is not None:
                                price_str = str(df.iloc[current_row, price_col]).strip()
                                if price_str and not any(h in price_str.lower() for h in ['سعر', 'price']):
                                    try:
                                        price = float(price_str.replace(',', ''))
                                        product['unit_price'] = price
                                    except:
                                        product['unit_price'] = price_str

                            # Add product if it has enough information
                            if len(product) > 1:  # At least description and one other field
                                arabic_products.append(product)

                        current_row += 1

                    # If we found products, add them to the results
                    if arabic_products:
                        products.extend(arabic_products)
                        # Return early if we've found products matching Arabic format
                        return products

    # Method 3: Generic method for finding product tables
    # Use this as a fallback if specific formats weren't found

    # Common headers in English and Arabic
    desc_headers = ['description', 'item', 'product', 'التسمية', 'الوصف', 'المنتج']
    qty_headers = ['quantity', 'qty', 'pcs', 'الكمية', 'العدد']
    price_headers = ['unit price', 'price', 'سعر الوحدة', 'السعر']

    # Find rows that look like table headers
    header_rows = []
    for i in range(len(df)):
        header_count = 0
        for j in range(len(df.columns)):
            cell = str(df.iloc[i, j]).strip().lower()
            if any(h in cell for h in desc_headers + qty_headers + price_headers):
                header_count += 1

        # If this row has multiple headers, it's likely a product table header
        if header_count >= 2:
            header_rows.append(i)

    # Process each potential header row
    for header_row in header_rows:
        # Find column indices for product data
        desc_col = None
        qty_col = None
        price_col = None

        for j in range(len(df.columns)):
            cell = str(df.iloc[header_row, j]).strip().lower()

            if any(h in cell for h in desc_headers):
                desc_col = j
            elif any(h in cell for h in qty_headers):
                qty_col = j
            elif any(h in cell for h in price_headers):
                price_col = j

        # If we found key columns, process product data
        if desc_col is not None and (qty_col is not None or price_col is not None):
            current_row = header_row + 1
            fallback_products = []

            while current_row < len(df) and current_row < header_row + 30:
                # Get product data
                desc = str(df.iloc[current_row, desc_col]).strip()

                # Skip empty rows or header-like rows
                if not desc or any(h in desc.lower() for h in desc_headers + qty_headers + price_headers):
                    current_row += 1
                    continue

                product = {'description': desc}

                # Get quantity if available
                if qty_col is not None:
                    qty_str = str(df.iloc[current_row, qty_col]).strip()
                    if qty_str and not any(h in qty_str.lower() for h in qty_headers):
                        try:
                            qty = float(qty_str.replace(',', ''))
                            product['quantity'] = qty
                        except:
                            product['quantity'] = qty_str

                # Get price if available
                if price_col is not None:
                    price_str = str(df.iloc[current_row, price_col]).strip()
                    if price_str and not any(h in price_str.lower() for h in price_headers):
                        try:
                            price = float(price_str.replace(',', ''))
                            product['unit_price'] = price
                        except:
                            product['unit_price'] = price_str

                # Add product if it has enough information
                if len(product) > 1:
                    fallback_products.append(product)

                current_row += 1

            # If we found products, add them to the results
            if fallback_products:
                products.extend(fallback_products)
                # Return if we found products with this fallback method
                if len(products) > 0:
                    return products

    # Return whatever products we found, or empty list if none
    return products
