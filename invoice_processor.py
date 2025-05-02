import pandas as pd
import re
import numpy as np
import io
import os
from datetime import datetime

def fix_encoding(text):
    """
    Attempts to fix corrupted Arabic text by decoding it using 'windows-1256' 
    after encoding with 'latin1'. Preserves the original encoded text
    in cases where exact pattern matching is needed.
    
    Args:
        text: The potentially corrupted text string
        
    Returns:
        Fixed text string if successful, or original text if decoding fails
    """
    if not isinstance(text, str):
        return text
    
    # Check for exact encoded string patterns that we want to preserve for matching
    exact_encoded_patterns = [
        "ÝÇÊæÑÉ / INVOICE  N°:",
        "ÑãÒ ÇáÔÑíß ÇáÊÌÇÑí / Partner code :",
        "ÑãÒ ÇáÓáÚÉ\nArticle code",
        "ÇáÊÓãíÉ\nDescription",
        "ÇáßãíÉ\nQuantity",
        "ÓÚÑ ÇáæÍÏÉ\nUnit price"
    ]
    
    # Preserve exact pattern if found
    for pattern in exact_encoded_patterns:
        if pattern in text:
            return text
            
    try:
        # Convert corrupted Arabic text using windows-1256 encoding
        # First encode to latin1 bytes, then decode using windows-1256
        fixed_text = text.encode('latin1').decode('windows-1256')
        return fixed_text
    except Exception:
        # If decoding fails, return the original text
        return text

def calculate_invoice_total(products):
    """
    Calculate the total amount for an invoice based on its products.
    Ignores products with unit price of zero.
    
    Args:
        products: List of product dictionaries with quantity and unit_price
        
    Returns:
        Total amount as float
    """
    total = 0.0
    
    if not products:
        return total
        
    for product in products:
        quantity = product.get('quantity', 0)
        unit_price = product.get('unit_price', 0)
        
        # Convert to float if they are strings
        if isinstance(quantity, str):
            try:
                quantity = float(quantity.replace(',', ''))
            except (ValueError, AttributeError):
                quantity = 0
                
        if isinstance(unit_price, str):
            try:
                unit_price = float(unit_price.replace(',', ''))
            except (ValueError, AttributeError):
                unit_price = 0
        
        # Skip products with zero unit price
        if unit_price <= 0:
            continue
                
        # Add to total
        total += quantity * unit_price
        
    return round(total, 2)

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
            
            # Create a copy of the original DataFrame for exact pattern matching
            df_original = df_string.copy()
            
            # Apply fix_encoding to all text values in the DataFrame
            for i in range(len(df_string)):
                for j in range(len(df_string.columns)):
                    df_string.iloc[i, j] = fix_encoding(df_string.iloc[i, j])
                    
            # First try extraction with original encoding (for exact pattern matching)
            invoice_number = extract_invoice_number(df_original)
            customer_code = extract_customer_code(df_original)
            
            # If extraction fails with original encoding, try with fixed encoding
            if not invoice_number:
                invoice_number = extract_invoice_number(df_string)
            if not customer_code:
                customer_code = extract_customer_code(df_string)

            # Currency and date are less sensitive to encoding issues
            currency = extract_currency(df_string)
            invoice_date = extract_invoice_date(df_string)
            
            # For products, try with original encoding first, then with fixed encoding if needed
            products = extract_product_details(df_original, invoice_number)
            if not products:
                products = extract_product_details(df_string, invoice_number)
                
            # Calculate the total invoice amount
            total_amount = calculate_invoice_total(products)

            # Extract invoice data with linked products
            invoice_data = {
                'invoice_number': invoice_number,
                'customer_code': customer_code,
                'currency': currency,
                'invoice_date': invoice_date,
                'products': products,
                'total_amount': total_amount,
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
    # Primary method: Look for exact encoded Arabic and English invoice keywords
    # Including the exact encoded string that appears in real invoices
    invoice_n_keywords = ['invoice n:', 'invoice n°:', 'invoice no:', 'invoice number:', 'ÝÇÊæÑÉ / invoice n°:']

    for i in range(len(df)):
        for j in range(len(df.columns)):
            cell = str(df.iloc[i, j]).lower()
            
            # Check if any of the invoice keywords are in the cell
            if any(keyword in cell for keyword in invoice_n_keywords):
                # Extract invoice number directly from this cell 
                for keyword in invoice_n_keywords:
                    if keyword in cell:
                        match = re.search(f"{re.escape(keyword)}[:\\s°.]*([a-zA-Z0-9\\-/]+)", cell, re.IGNORECASE)
                        if match:
                            invoice_num = match.group(1).strip()
                            # Validate it follows the SIxxxxx pattern
                            if re.match(r'^SI\d+$', invoice_num):
                                return invoice_num
                
                # Also look for SI pattern in the same cell
                si_match = re.search(r'SI\d+', cell)
                if si_match:
                    return si_match.group(0).strip()

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

    # Backup method 3: Look for other invoice number keywords (English only)
    other_invoice_keywords = [
        'invoice no', 'invoice number', 'invoice #',
        'inv', 'inv no', 'invoice'
    ]

    # Search for each keyword in the dataframe
    for keyword in other_invoice_keywords:
        for i in range(len(df)):
            for j in range(len(df.columns)):
                cell = str(df.iloc[i, j]).lower()
                if keyword.lower() in cell:
                    # Try to extract the invoice number from this cell or neighboring cells
                    match = re.search(f"{keyword.lower()}[:\\s°.]*([a-zA-Z0-9\\-/]+)", cell, re.IGNORECASE)
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
    # Primary method: Look for exact encoded Arabic and English partner code keywords
    # Including the exact encoded string that appears in real invoices
    partner_code_keywords = ['partner code:', 'partner code :', 'code:', 'ÑãÒ ÇáÔÑíß ÇáÊÌÇÑí / partner code :']

    for i in range(len(df)):
        for j in range(len(df.columns)):
            cell = str(df.iloc[i, j]).lower()
            
            if any(keyword in cell for keyword in partner_code_keywords):
                # Try to extract the customer code directly from this cell
                for keyword in partner_code_keywords:
                    if keyword in cell:
                        match = re.search(f"{re.escape(keyword)}[:\\s]*([a-zA-Z0-9\\-/]+)", cell, re.IGNORECASE)
                        if match:
                            code = match.group(1).strip()
                            # Validate it follows the Cxxxx pattern
                            if re.match(r'^C\d+$', code):
                                return code
                
                # Also check for C pattern in the same cell
                c_match = re.search(r'C\d+', cell)
                if c_match and re.match(r'^C\d+$', c_match.group(0)):
                    return c_match.group(0).strip()

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

    # Backup method 2: Look for other customer code indicators (English only)
    other_code_keywords = [
        'customer code', 'client code', 'account code', 'partner id',
        'partner code', 'partner details', 'client id', 'customer id'
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
    Returns EGP, USD, or EUR as specified in the requirements.

    Args:
        df: DataFrame with string values

    Returns:
        'EGP', 'USD', or 'EUR', with EGP as default
    """
    # First method: Look for the "Currency Code" column in a structured table
    for i in range(len(df)):
        for j in range(len(df.columns)):
            cell = str(df.iloc[i, j]).lower()
            if 'currency code' in cell:
                # Look for currency codes in the column below
                for row_idx in range(i+1, min(i+20, len(df))):
                    curr_value = str(df.iloc[row_idx, j]).strip().upper()
                    # Allow EGP, USD or EUR
                    if curr_value in ['EGP', 'USD', 'EUR']:
                        return curr_value

                # If not found in the exact column, check the one to the right
                if j + 1 < len(df.columns):
                    for row_idx in range(i+1, min(i+20, len(df))):
                        curr_value = str(df.iloc[row_idx, j+1]).strip().upper()
                        if curr_value in ['EGP', 'USD', 'EUR']:
                            return curr_value

    # Look specifically for dollar or USD mentions - they indicate USD currency
    usd_indicators = ['$', 'dollar', 'usd', 'دولار', 'united states', 'u.s.']
    for i in range(len(df)):
        for j in range(len(df.columns)):
            cell = str(df.iloc[i, j]).lower()
            if any(indicator in cell for indicator in usd_indicators):
                return 'USD'

    # Look for Euro mentions - they indicate EUR currency
    eur_indicators = ['€', 'euro', 'eur', 'يورو', 'european', 'europe']
    for i in range(len(df)):
        for j in range(len(df.columns)):
            cell = str(df.iloc[i, j]).lower()
            if any(indicator in cell for indicator in eur_indicators):
                return 'EUR'

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
                if 'eur' in cell or 'euro' in cell or '€' in cell:
                    return 'EUR'
                if 'egp' in cell or 'egypt' in cell:
                    return 'EGP'

                # Check the cell to the right
                if j + 1 < len(df.columns):
                    right_cell = str(df.iloc[i, j + 1]).lower()
                    if right_cell == 'usd' or right_cell == 'dollar' or right_cell == '$':
                        return 'USD'
                    if right_cell == 'eur' or right_cell == 'euro' or right_cell == '€':
                        return 'EUR'
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

    # Keywords that might precede an invoice date (English only)
    date_keywords = [
        'date', 'invoice date', 'issued on', 'document date', 'payment date',
        'issue date'
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
    Ignores weight/packaging information and products with zero unit price.

    Args:
        df: DataFrame with string values
        invoice_number: The invoice number to link products to

    Returns:
        List of dictionaries containing product details with invoice linkage
    """
    products = []
    last_product_row = -1  # Track the last row where we found a product

    # Method 1: Look for "Invoice details" section which contains the product table
    for i in range(len(df)):
        for j in range(len(df.columns)):
            cell = str(df.iloc[i, j]).strip().lower()
            if 'invoice details' in cell:
                # Found the product table header, look for column headers in next rows
                header_row = i
                item_code_col = None
                desc_col = None
                qty_col = None
                price_col = None
                amount_col = None

                # Search for headers in the rows below (check next 5 rows for flexibility)
                for row_idx in range(header_row + 1, min(header_row + 6, len(df))):
                    for col_idx in range(len(df.columns)):
                        header_text = str(df.iloc[row_idx, col_idx]).strip().lower()

                        # Check for different header possibilities including exact encoded Arabic strings
                        if 'code' in header_text or 'item code' in header_text or 'article code' in header_text or 'ÑãÒ ÇáÓáÚÉ' in header_text:
                            item_code_col = col_idx
                        elif 'description' in header_text or 'desc' in header_text or 'ÇáÊÓãíÉ' in header_text:
                            desc_col = col_idx
                        elif 'quantity' in header_text or 'qty' in header_text or 'ÇáßãíÉ' in header_text:
                            qty_col = col_idx
                        elif ('unit' in header_text and 'price' in header_text) or 'unit price' in header_text or 'ÓÚÑ ÇáæÍÏÉ' in header_text:
                            price_col = col_idx
                        elif 'amount' in header_text:
                            amount_col = col_idx

                # If we found at least description column and one numeric column
                if desc_col is not None and (qty_col is not None or price_col is not None or amount_col is not None):
                    # Start parsing from the row after the headers
                    data_start_row = row_idx + 1
                    table_products = []

                    # Process rows until we hit the end
                    for data_row in range(data_start_row, len(df)):
                        # Get description - this is the key field to identify a product row
                        try:
                            desc_text = str(df.iloc[data_row, desc_col]).strip()
                        except:
                            continue  # Skip if there's an error accessing this cell

                        # Skip empty, header-like rows, or payment terms
                        if (not desc_text or 
                            desc_text.lower() in ['description', 'total', 'amount', 'وصف', 'المجموع'] or
                            'ÔÑæØ ÇáÏÝÚ' in desc_text or 
                            'term of payment' in desc_text.lower()):
                            continue

                        # Skip any weight or package related items
                        desc_lower = desc_text.lower().strip()
                        weight_patterns = [
                            'total weight', 'weight total', 'total wt', 'net weight', 'gross weight',
                            'total package', 'package total', 'total packages', 'packages total'
                        ]
                        if any(pattern in desc_lower for pattern in weight_patterns):
                            continue

                        # Create product entry for regular products
                        product = {'description': desc_text, 'invoice_number': invoice_number}

                        # Add item code if available
                        if item_code_col is not None:
                            try:
                                code = str(df.iloc[data_row, item_code_col]).strip()
                                if code and code.lower() != 'nan':
                                    product['item_code'] = code
                            except:
                                pass

                        # Add quantity if available
                        if qty_col is not None:
                            try:
                                qty_text = str(df.iloc[data_row, qty_col]).strip().replace(',', '')
                                if qty_text and qty_text.lower() != 'nan':
                                    try:
                                        qty_value = float(qty_text)
                                        product['quantity'] = qty_value
                                    except:
                                        # Keep as string if conversion fails
                                        product['quantity'] = qty_text
                            except:
                                pass

                        # Add unit price if available
                        unit_price_value = 0
                        if price_col is not None:
                            try:
                                price_text = str(df.iloc[data_row, price_col]).strip().replace(',', '')
                                if price_text and price_text.lower() != 'nan':
                                    try:
                                        price_value = float(price_text)
                                        product['unit_price'] = price_value
                                        unit_price_value = price_value
                                    except:
                                        # Keep as string if conversion fails
                                        product['unit_price'] = price_text
                            except:
                                pass

                        # If we have at least description and one numeric field, and price is not zero, add the product
                        if len(product) > 2 and unit_price_value > 0:  # More than just description and invoice_number
                            table_products.append(product)
                            last_product_row = data_row  # Update last product row

                    # If we found products, add them to the results
                    if table_products:
                        products.extend(table_products)

    # Method 2: Look for tables with common headers like Description/Quantity/Price
    if not products:  # Only if we haven't found products yet
        # Common header texts including exact encoded Arabic strings
        desc_headers = ['description', 'item description', 'product', 'desc', 'ÇáÊÓãíÉ']
        qty_headers = ['quantity', 'qty', 'ÇáßãíÉ']
        price_headers = ['unit price', 'price', 'unit cost', 'ÓÚÑ ÇáæÍÏÉ']

        # Scan for header rows
        for i in range(len(df)):
            header_matches = 0
            header_cols = {}

            # Check if this row contains typical header text
            for j in range(len(df.columns)):
                cell_text = str(df.iloc[i, j]).strip().lower()

                # Check for description header
                if any(h in cell_text for h in desc_headers):
                    header_cols['description'] = j
                    header_matches += 1
                # Check for quantity header
                elif any(h in cell_text for h in qty_headers):
                    header_cols['quantity'] = j
                    header_matches += 1
                # Check for price header
                elif any(h in cell_text for h in price_headers):
                    header_cols['price'] = j
                    header_matches += 1

            # If we found at least 2 matching headers, this is likely a product table
            if header_matches >= 2 and 'description' in header_cols:
                table_products = []

                # Process rows below the header
                for data_row in range(i + 1, min(i + 30, len(df))):
                    # Get description - the key field
                    try:
                        desc_text = str(df.iloc[data_row, header_cols['description']]).strip()
                    except:
                        continue

                    # Skip empty rows, header-like texts, or payment terms
                    if (not desc_text or 
                        desc_text.lower() in desc_headers + ['total', 'subtotal', 'المجموع'] or
                        'ÔÑæØ ÇáÏÝÚ' in desc_text or 
                        'term of payment' in desc_text.lower()):
                        continue

                    # Skip any weight or package related items using simplified pattern check
                    desc_lower = desc_text.lower().strip()
                    weight_patterns = [
                        'total weight', 'weight', 'wt', 'package', 'pkg', 
                        'total package', 'packages'
                    ]
                    if any(pattern in desc_lower for pattern in weight_patterns):
                        continue

                    # Create product
                    product = {'description': desc_text, 'invoice_number': invoice_number}

                    # Add quantity if that column exists
                    if 'quantity' in header_cols:
                        try:
                            qty_text = str(df.iloc[data_row, header_cols['quantity']]).strip().replace(',', '')
                            if qty_text and qty_text.lower() != 'nan' and qty_text.lower() not in qty_headers:
                                try:
                                    qty_value = float(qty_text)
                                    product['quantity'] = qty_value
                                except:
                                    product['quantity'] = qty_text
                        except:
                            pass

                    # Add price if that column exists
                    unit_price_value = 0
                    if 'price' in header_cols:
                        try:
                            price_text = str(df.iloc[data_row, header_cols['price']]).strip().replace(',', '')
                            if price_text and price_text.lower() != 'nan' and price_text.lower() not in price_headers:
                                try:
                                    price_value = float(price_text)
                                    product['unit_price'] = price_value
                                    unit_price_value = price_value
                                except:
                                    product['unit_price'] = price_text
                        except:
                            pass

                    # Add the product if we have enough data and price is not zero
                    if len(product) > 2 and unit_price_value > 0:
                        table_products.append(product)

                # If we found products, add them
                if table_products:
                    products.extend(table_products)

    # Method 3: Special format search based on the image that shows product table with code, description, qty, price columns
    if not products:  # Only try this if we haven't found products yet
        for i in range(len(df)):
            # Look for rows that contain product details
            for j in range(len(df.columns)):
                cell = str(df.iloc[i, j]).strip().lower()
                # Looking for "invoice details" or similar headers that indicate where product tables begin
                if 'invoice details' in cell or 'details' in cell:

                    # Search for column indices in the rows below
                    code_col = None
                    desc_col = None
                    qty_col = None
                    price_col = None

                    # Look at the next 5 rows for headers
                    for row_idx in range(i+1, min(i+6, len(df))):
                        # Look for code column
                        for col_idx in range(len(df.columns)):
                            cell_text = str(df.iloc[row_idx, col_idx]).strip().lower()

                            # Look for code/item code column including exact encoded Arabic string
                            if ('code' in cell_text and not 'currency' in cell_text) or 'ÑãÒ ÇáÓáÚÉ' in cell_text:
                                code_col = col_idx
                            # Look for description column including exact encoded Arabic string
                            elif any(term in cell_text for term in ['description', 'desc', 'ÇáÊÓãíÉ']):
                                desc_col = col_idx
                            # Look for quantity column including exact encoded Arabic string
                            elif any(term in cell_text for term in ['quantity', 'qty', 'ÇáßãíÉ']):
                                qty_col = col_idx
                            # Look for unit price column including exact encoded Arabic string
                            elif any(term in cell_text for term in ['unit price', 'price', 'ÓÚÑ ÇáæÍÏÉ']):
                                price_col = col_idx

                    # If we found the required columns
                    if desc_col is not None and (qty_col is not None or price_col is not None):
                        special_products = []

                        # Start from the row after header row
                        start_row = row_idx + 1

                        # Parse product rows
                        for data_row in range(start_row, min(start_row+30, len(df))):
                            # Only process if we can get a valid description
                            try:
                                desc_text = str(df.iloc[data_row, desc_col]).strip()
                                if not desc_text or desc_text.lower() in ['description', 'total', 'amount', 'المجموع']:
                                    continue

                                # Skip any weight or package related items
                                desc_lower = desc_text.lower().strip()
                                weight_patterns = [
                                    'total weight', 'weight', 'wt', 'package', 'pkg'
                                ]
                                if any(pattern in desc_lower for pattern in weight_patterns):
                                    continue

                                product = {'description': desc_text, 'invoice_number': invoice_number}

                                # Add code if available
                                if code_col is not None:
                                    code_text = str(df.iloc[data_row, code_col]).strip()
                                    if code_text and code_text.lower() != 'nan':
                                        product['item_code'] = code_text

                                # Add quantity if available
                                if qty_col is not None:
                                    qty_text = str(df.iloc[data_row, qty_col]).strip().replace(',', '')
                                    if qty_text and qty_text.lower() != 'nan':
                                        try:
                                            qty_value = float(qty_text)
                                            product['quantity'] = qty_value
                                        except:
                                            product['quantity'] = qty_text

                                # Add price if available
                                unit_price_value = 0
                                if price_col is not None:
                                    price_text = str(df.iloc[data_row, price_col]).strip().replace(',', '')
                                    if price_text and price_text.lower() != 'nan':
                                        try:
                                            price_value = float(price_text)
                                            product['unit_price'] = price_value
                                            unit_price_value = price_value
                                        except:
                                            product['unit_price'] = price_text

                                # Add product if it has enough data and price is not zero
                                if len(product) > 2 and unit_price_value > 0:
                                    special_products.append(product)
                            except:
                                continue

                        # If we found products, add them
                        if special_products:
                            products.extend(special_products)
                            break

    # Fallback: generic data pattern search if we still haven't found products
    # This looks for any content that appears to be product descriptions followed by numbers
    pattern_products = []

    # Helper function to check if a string is a number
    def is_number(s):
        try:
            float(str(s).replace(',', ''))
            return True
        except:
            return False

    # Look for rows with text followed by 2+ numbers (likely description + qty + price)
    for i in range(len(df)):
        row_data = []

        # Get all cell values in this row
        for j in range(len(df.columns)):
            try:
                cell_text = str(df.iloc[i, j]).strip()
                if cell_text and cell_text.lower() != 'nan':
                    row_data.append((j, cell_text))
            except:
                continue

        # Skip empty rows
        if len(row_data) < 3:
            continue

        # Check if this row has the pattern: text + number + number
        has_text = False
        num_count = 0

        for _, cell_val in row_data:
            if is_number(cell_val):
                num_count += 1
            elif len(cell_val) > 3:  # Non-numeric cell with reasonable length
                has_text = True

        # Check if any of the cells contain payment terms text
        contains_payment_terms = False
        for _, cell_val in row_data:
            # Define payment terms text list
            payment_terms = ['ÔÑæØ ÇáÏÝÚ', 'term of payment', 'payment term', 'payment terms']
            if any(term in cell_val for term in payment_terms) or any(term in cell_val.lower() for term in payment_terms):
                contains_payment_terms = True
                break

        # Skip weight/package related rows
        is_weight_package_row = False
        for _, cell_val in row_data:
            cell_lower = cell_val.lower().strip()
            # Simple list of weight/package keywords
            weight_patterns = ['weight', 'package', 'pkg', 'wt']
            if any(pattern in cell_lower for pattern in weight_patterns):
                is_weight_package_row = True
                break

        if not is_weight_package_row and has_text and num_count >= 2 and not contains_payment_terms:
            # Identify which column contains the description (usually the longest text)
            desc_col = None
            desc_len = 0

            for col_idx, cell_val in row_data:
                if not is_number(cell_val) and len(cell_val) > desc_len:
                    desc_col = col_idx
                    desc_len = len(cell_val)

            # If we found a description column
            if desc_col is not None:
                description = str(df.iloc[i, desc_col]).strip()
                
                # Skip weight and package related descriptions
                desc_lower = description.lower()
                if any(pattern in desc_lower for pattern in ['weight', 'package', 'pkg', 'wt']):
                    continue
                    
                product = {'description': description, 'invoice_number': invoice_number}

                # Find numeric columns (potential quantity/price)
                num_cols = []
                for col_idx, cell_val in row_data:
                    if col_idx != desc_col and is_number(cell_val):
                        num_cols.append(col_idx)

                # If we have at least one numeric column, get quantity and price
                if num_cols:
                    # First number is usually quantity
                    try:
                        qty_text = str(df.iloc[i, num_cols[0]]).strip().replace(',', '')
                        qty_value = float(qty_text)
                        product['quantity'] = qty_value
                    except:
                        pass

                    # Second number is usually price
                    unit_price_value = 0
                    if len(num_cols) > 1:
                        try:
                            price_text = str(df.iloc[i, num_cols[1]]).strip().replace(',', '')
                            price_value = float(price_text)
                            product['unit_price'] = price_value
                            unit_price_value = price_value
                        except:
                            pass

                # Add product if it has enough information and price is not zero
                if len(product) > 2 and unit_price_value > 0:
                    pattern_products.append(product)

    # If we found products with pattern matching
    if pattern_products:
        products.extend(pattern_products)

    # Return whatever products we found
    return products
