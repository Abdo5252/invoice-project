import pandas as pd
import io
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

def create_output_excel(processed_invoices, template_file):
    """
    Create a new Excel file based on the template and fill it with processed invoice data.
    
    Args:
        processed_invoices: List of dictionaries containing processed invoice data
        template_file: The uploaded template Excel file object
        
    Returns:
        BytesIO object containing the Excel file
    """
    # Read the template file
    template_workbook = openpyxl.load_workbook(template_file)
    
    # Create a new workbook for output
    output = io.BytesIO()
    
    # If there are no invoices, return an empty file with just the template
    if not processed_invoices:
        template_workbook.save(output)
        output.seek(0)
        return output
    
    # Get the first sheet of the template as reference
    template_sheet = template_workbook.active
    
    # For each processed invoice, create a new sheet based on the template
    for i, invoice in enumerate(processed_invoices):
        # If this is the first invoice, use the existing sheet
        if i == 0:
            sheet = template_sheet
            sheet.title = f"Invoice_{invoice.get('invoice_number', i+1)}"
        else:
            # Create a new sheet for subsequent invoices
            sheet = template_workbook.copy_worksheet(template_sheet)
            sheet.title = f"Invoice_{invoice.get('invoice_number', i+1)}"
        
        # Fill in the invoice data into the template
        populate_template_sheet(sheet, invoice)
    
    # Save the workbook to the BytesIO object
    template_workbook.save(output)
    output.seek(0)
    
    return output

def populate_template_sheet(sheet, invoice_data):
    """
    Populate a template sheet with invoice data.
    
    Args:
        sheet: The openpyxl worksheet object to populate
        invoice_data: Dictionary containing the invoice data
    """
    # Scan the sheet for placeholders or fields to populate
    invoice_number_keywords = ['invoice number', 'invoice no', 'inv #', 'فاتورة رقم', 'رقم الفاتورة']
    customer_code_keywords = ['customer code', 'client code', 'partner code', 'رمز العميل', 'كود العميل']
    currency_keywords = ['currency', 'العملة', 'curr']
    
    # Look through all cells for relevant keywords to replace data
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value:
                cell_value = str(cell.value).lower()
                
                # Check for invoice number field
                if any(keyword in cell_value for keyword in invoice_number_keywords):
                    # Check if cell already has a value format like "Invoice #: "
                    if ':' in cell_value or '：' in cell_value:
                        # Replace after the colon
                        prefix = cell.value.split(':', 1)[0] + ': '
                        cell.value = prefix + str(invoice_data.get('invoice_number', ''))
                    else:
                        # Place value in adjacent cell
                        col_idx = cell.column + 1
                        sheet.cell(row=cell.row, column=col_idx).value = invoice_data.get('invoice_number', '')
                
                # Check for customer code field
                elif any(keyword in cell_value for keyword in customer_code_keywords):
                    # Check if cell already has a value format like "Customer Code: "
                    if ':' in cell_value or '：' in cell_value:
                        # Replace after the colon
                        prefix = cell.value.split(':', 1)[0] + ': '
                        cell.value = prefix + str(invoice_data.get('customer_code', ''))
                    else:
                        # Place value in adjacent cell
                        col_idx = cell.column + 1
                        sheet.cell(row=cell.row, column=col_idx).value = invoice_data.get('customer_code', '')
                
                # Check for currency field
                elif any(keyword in cell_value for keyword in currency_keywords):
                    # Check if cell already has a value format like "Currency: "
                    if ':' in cell_value or '：' in cell_value:
                        # Replace after the colon
                        prefix = cell.value.split(':', 1)[0] + ': '
                        cell.value = prefix + str(invoice_data.get('currency', ''))
                    else:
                        # Place value in adjacent cell
                        col_idx = cell.column + 1
                        sheet.cell(row=cell.row, column=col_idx).value = invoice_data.get('currency', '')
    
    # Find product table area in the template
    product_table_start_row = None
    description_col = None
    quantity_col = None
    unit_price_col = None
    
    # English and Arabic header variations
    description_headers = ['description', 'product', 'item', 'التسمية', 'الوصف', 'المنتج']
    quantity_headers = ['quantity', 'qty', 'الكمية', 'الكميه', 'العدد']
    price_headers = ['unit price', 'price', 'سعر الوحدة', 'السعر']
    
    # Search for the product table headers
    for row_idx, row in enumerate(sheet.iter_rows(), 1):
        for col_idx, cell in enumerate(row, 1):
            if cell.value:
                cell_value = str(cell.value).lower()
                
                if any(header in cell_value for header in description_headers):
                    description_col = col_idx
                    product_table_start_row = row_idx
                
                elif any(header in cell_value for header in quantity_headers):
                    quantity_col = col_idx
                    if not product_table_start_row:
                        product_table_start_row = row_idx
                
                elif any(header in cell_value for header in price_headers):
                    unit_price_col = col_idx
                    if not product_table_start_row:
                        product_table_start_row = row_idx
        
        # If we found at least two columns, consider it a valid product table
        if product_table_start_row and sum([bool(description_col), bool(quantity_col), bool(unit_price_col)]) >= 2:
            break
    
    # If we found a product table, populate it with product data
    if product_table_start_row and 'products' in invoice_data and invoice_data['products']:
        # Start from the row after the header
        current_row = product_table_start_row + 1
        
        for product in invoice_data['products']:
            # Add description
            if description_col and 'description' in product:
                sheet.cell(row=current_row, column=description_col).value = product['description']
            
            # Add quantity
            if quantity_col and 'quantity' in product:
                sheet.cell(row=current_row, column=quantity_col).value = product['quantity']
            
            # Add unit price
            if unit_price_col and 'unit_price' in product:
                sheet.cell(row=current_row, column=unit_price_col).value = product['unit_price']
            
            current_row += 1
