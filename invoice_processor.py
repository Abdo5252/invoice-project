import pandas as pd
import re
import numpy as np
import io
import os

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
            
            # Extract invoice data
            invoice_data = {
                'invoice_number': extract_invoice_number(df_string),
                'customer_code': extract_customer_code(df_string),
                'currency': extract_currency(df_string),
                'products': extract_product_details(df_string),
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
    # Keywords that might precede an invoice number
    invoice_keywords = [
        'invoice n', 'invoice no', 'invoice number', 'invoice #',
        'inv', 'inv no', 'invoice', 'فاتورة رقم', 'رقم الفاتورة'
    ]
    
    # Search for each keyword in the dataframe
    for keyword in invoice_keywords:
        for i in range(len(df)):
            for j in range(len(df.columns)):
                cell = df.iloc[i, j].lower()
                if keyword.lower() in cell:
                    # Try to extract the invoice number from this cell or neighboring cells
                    # First check if the number is in the same cell after the keyword
                    match = re.search(f"{keyword.lower()}[:\s]*([a-zA-Z0-9\-/]+)", cell, re.IGNORECASE)
                    if match:
                        return match.group(1).strip()
                    
                    # Check the cell to the right
                    if j + 1 < len(df.columns):
                        right_cell = df.iloc[i, j + 1]
                        if right_cell and not 'nan' in right_cell:
                            return right_cell.strip()
                    
                    # Check the cell below
                    if i + 1 < len(df):
                        below_cell = df.iloc[i + 1, j]
                        if below_cell and not 'nan' in below_cell:
                            return below_cell.strip()
    
    return None

def extract_customer_code(df):
    """
    Extract customer code from the dataframe.
    
    Args:
        df: DataFrame with string values
        
    Returns:
        Extracted customer code or None if not found
    """
    # Keywords that might precede a customer code
    customer_keywords = [
        'partner code', 'customer code', 'client code', 'account code',
        'partner id', 'customer id', 'client id', 'customer',
        'رمز العميل', 'كود العميل', 'رقم العميل'
    ]
    
    # Search for each keyword in the dataframe
    for keyword in customer_keywords:
        for i in range(len(df)):
            for j in range(len(df.columns)):
                cell = df.iloc[i, j].lower()
                if keyword.lower() in cell:
                    # Try to extract the customer code from this cell or neighboring cells
                    # First check if the code is in the same cell after the keyword
                    match = re.search(f"{keyword.lower()}[:\s]*([a-zA-Z0-9\-/]+)", cell, re.IGNORECASE)
                    if match:
                        return match.group(1).strip()
                    
                    # Check the cell to the right
                    if j + 1 < len(df.columns):
                        right_cell = df.iloc[i, j + 1]
                        if right_cell and not 'nan' in right_cell:
                            return right_cell.strip()
                    
                    # Check the cell below
                    if i + 1 < len(df):
                        below_cell = df.iloc[i + 1, j]
                        if below_cell and not 'nan' in below_cell:
                            return below_cell.strip()
    
    return None

def extract_currency(df):
    """
    Extract currency information from the dataframe.
    
    Args:
        df: DataFrame with string values
        
    Returns:
        Extracted currency or None if not found
    """
    # Common currency symbols and names
    currency_patterns = [
        r'\$', r'€', r'£', r'¥', 
        r'USD', r'EUR', r'GBP', r'JPY', r'AED', r'SAR',
        r'dollar', r'euro', r'dirham', r'riyal',
        r'دولار', r'يورو', r'درهم', r'ريال'
    ]
    
    # Currency keywords that might be used
    currency_keywords = [
        'currency', 'العملة', 'curr.', 'curr', 
        'عملة', 'بعملة'
    ]
    
    # First look for explicit currency indicators
    for keyword in currency_keywords:
        for i in range(len(df)):
            for j in range(len(df.columns)):
                cell = df.iloc[i, j].lower()
                if keyword.lower() in cell:
                    # Try to extract the currency from this cell or neighboring cells
                    for pattern in currency_patterns:
                        match = re.search(pattern, cell, re.IGNORECASE)
                        if match:
                            return match.group(0).strip()
                    
                    # Check the cell to the right
                    if j + 1 < len(df.columns):
                        right_cell = df.iloc[i, j + 1]
                        for pattern in currency_patterns:
                            match = re.search(pattern, right_cell, re.IGNORECASE)
                            if match:
                                return match.group(0).strip()
    
    # If no explicit currency indicator found, look for currency symbols in price columns
    for i in range(len(df)):
        for j in range(len(df.columns)):
            cell = df.iloc[i, j]
            for pattern in currency_patterns:
                match = re.search(pattern, cell, re.IGNORECASE)
                if match:
                    return match.group(0).strip()
    
    return None

def extract_product_details(df):
    """
    Extract product details from tables in the dataframe.
    
    Args:
        df: DataFrame with string values
        
    Returns:
        List of dictionaries containing product details
    """
    # English header variations
    description_headers_en = ['description', 'product', 'item', 'service', 'detail', 'product description']
    quantity_headers_en = ['quantity', 'qty', 'pcs', 'amount', 'count']
    price_headers_en = ['unit price', 'price', 'rate', 'unit cost', 'price per unit']
    
    # Arabic header variations
    description_headers_ar = ['التسمية', 'الوصف', 'المنتج', 'البند', 'الخدمة', 'التفاصيل']
    quantity_headers_ar = ['الكمية', 'الكميه', 'العدد', 'قطع']
    price_headers_ar = ['سعر الوحدة', 'السعر', 'التكلفة', 'سعر الوحده']
    
    # Combined headers
    description_headers = description_headers_en + description_headers_ar
    quantity_headers = quantity_headers_en + quantity_headers_ar
    price_headers = price_headers_en + price_headers_ar
    
    # Find potential header rows
    header_rows = []
    for i in range(len(df)):
        row = df.iloc[i]
        desc_match = False
        qty_match = False
        price_match = False
        
        for j in range(len(row)):
            cell = str(row[j]).lower()
            
            if any(header in cell for header in description_headers):
                desc_match = True
            if any(header in cell for header in quantity_headers):
                qty_match = True
            if any(header in cell for header in price_headers):
                price_match = True
        
        # If row has at least two types of headers, consider it a potential header row
        if sum([desc_match, qty_match, price_match]) >= 2:
            header_rows.append(i)
    
    products = []
    
    # Process each potential table
    for header_row in header_rows:
        if header_row + 1 >= len(df):
            continue
            
        # Find column indices for each field
        desc_col = None
        qty_col = None
        price_col = None
        
        for j in range(len(df.columns)):
            cell = str(df.iloc[header_row, j]).lower()
            
            if any(header in cell for header in description_headers):
                desc_col = j
            if any(header in cell for header in quantity_headers):
                qty_col = j
            if any(header in cell for header in price_headers):
                price_col = j
        
        # If we found at least description and one other column
        if desc_col is not None and (qty_col is not None or price_col is not None):
            # Extract products from rows following the header
            current_row = header_row + 1
            
            while current_row < len(df):
                # If row is empty or contains header-like content, stop extraction
                if all(str(cell) == '' or 'nan' in str(cell) for cell in df.iloc[current_row]):
                    break
                    
                # If row has potential headers, stop extraction
                if any(header in str(df.iloc[current_row, j]).lower() 
                       for j in range(len(df.columns)) 
                       for header in description_headers + quantity_headers + price_headers):
                    break
                
                # Extract product data
                product = {}
                
                if desc_col is not None:
                    product['description'] = str(df.iloc[current_row, desc_col]).strip()
                
                if qty_col is not None:
                    qty = str(df.iloc[current_row, qty_col]).strip()
                    # Try to convert to numeric if possible
                    try:
                        qty = float(qty.replace(',', ''))
                    except:
                        pass
                    product['quantity'] = qty
                
                if price_col is not None:
                    price = str(df.iloc[current_row, price_col]).strip()
                    # Extract numeric part if it contains currency symbols
                    price_match = re.search(r'([0-9,.]+)', price)
                    if price_match:
                        try:
                            price = float(price_match.group(1).replace(',', ''))
                        except:
                            pass
                    product['unit_price'] = price
                
                # Only add product if it has a description and at least one other field
                if product.get('description') and (product.get('quantity') is not None or product.get('unit_price') is not None):
                    # Skip if description is just a placeholder or seems like a header
                    if not any(header in product['description'].lower() 
                              for header in description_headers + quantity_headers + price_headers):
                        products.append(product)
                
                current_row += 1
    
    return products
