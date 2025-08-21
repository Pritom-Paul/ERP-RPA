import fitz  # PyMuPDF
import re
import sys
import pdfplumber
from datetime import datetime, timedelta

def replace_shipment_date(doc):
    """
    Replace shipment date in a PDF document with a date 15 days before the original date
    
    Args:
        doc: PyMuPDF document object
    
    Returns:
        bool: True if successful, False otherwise
    """
    # Extract text from the first page to find the shipment date
    text = doc[0].get_text("text")
    
    # Regex pattern to find the shipment date (MM-DD-YYYY format)
    # This pattern looks for "Shipment Date ETD" followed by the date
    date_pattern = r"Shipment Date ETD\s*(\d{2}-\d{2}-\d{4})"
    match = re.search(date_pattern, text)
    
    # If not found with PyMuPDF, try pdfplumber as fallback
    if not match:
        print("Shipment date not found with PyMuPDF, trying pdfplumber...")
        try:
            # Extract text using pdfplumber
            with pdfplumber.open(doc.name) as pdf:
                all_text = ""
                for page in pdf.pages:
                    all_text += page.extract_text() + "\n"
            
            # Try the pattern with pdfplumber extracted text
            match = re.search(r"Shipment Date ETD\s*([0-9]{1,2}-[0-9]{2}-[0-9]{4})", all_text)
            
            if not match:
                print("Shipment date not found in the PDF with pdfplumber either")
                return False
        except Exception as e:
            print(f"Error using pdfplumber: {e}")
            return False
    
    old_shipment_date = match.group(1)
    print(f"Found shipment date: {old_shipment_date}")
    
    # Parse the original date and subtract 15 days
    original_date = datetime.strptime(old_shipment_date, "%m-%d-%Y")
    new_date = original_date - timedelta(days=15)
    new_shipment_date = new_date.strftime("%m-%d-%Y")
    print(f"New shipment date (15 days before): {new_shipment_date}")
    
    # Iterate through each page to find and replace the date
    for page_num in range(len(doc)):
        page = doc[page_num]
        
        # Search for the old shipment date text
        text_instances = page.search_for(old_shipment_date)
        
        # Remove each found instance
        for inst in text_instances:
            # Add a white rectangle to cover the old date
            page.add_redact_annot(inst, fill=(1, 1, 1))  # White fill
            
        # Apply the redactions (remove the text)
        page.apply_redactions()
        
        # Now add the new shipment date at the first found location
        if text_instances:
            # Use the position of the first found instance
            rect = text_instances[0]
            
            # Insert the new text at the same position
            page.insert_text(
                (rect.x0, rect.y0+8.5),  # Position
                new_shipment_date,   # New text
                fontsize=7.5,          # Adjust font size to match original
                color=(0, 0, 0)      # Black text
            )
    
    return True

def replace_price(doc):
    """
    Replace all prices in the PDF document by reducing each by 0.75
    
    Args:
        doc: PyMuPDF document object
    
    Returns:
        bool: True if successful, False otherwise
    """
    # Extract text from all pages to find all prices
    all_text = ""
    for page_num in range(len(doc)):
        all_text += doc[page_num].get_text("text") + "\n"
    
    # Regex pattern to find all prices (looks for "Price" followed by the value)
    price_pattern = r"Price\s*([\d.,]+)"
    matches = re.findall(price_pattern, all_text, re.IGNORECASE)
    
    if not matches:
        print("No prices found in the PDF")
        return False
    
    print(f"Found {len(matches)} prices: {matches}")
    
    # Create a list of price conversions
    price_conversions = []
    
    for old_price_text in matches:
        # Clean and convert the price
        cleaned_price = old_price_text.replace(',', '.')
        dot_count = cleaned_price.count('.')
        
        # If there's more than one dot, remove all but the last one
        if dot_count > 1:
            parts = [part for part in cleaned_price.split('.') if part]
            if len(parts) > 1:
                cleaned_price = ''.join(parts[:-1]) + '.' + parts[-1]
            else:
                cleaned_price = ''.join(parts)
        
        # Convert to float and reduce by 0.75
        try:
            old_price_float = float(cleaned_price)
            new_price_float = old_price_float - 0.75
            # Format back to European style (comma as decimal separator)
            new_price_text = f"{new_price_float:.2f}".replace('.', ',')
            price_conversions.append((old_price_text, new_price_text))
            print(f"Price {old_price_text} → {new_price_text}")
        except ValueError:
            print(f"Could not convert price to float: {cleaned_price}")
            continue
    
    # Sort by length (longest first) to avoid partial replacements
    price_conversions.sort(key=lambda x: len(x[0]), reverse=True)
    
    # Iterate through each page to find and replace all prices
    for page_num in range(len(doc)):
        page = doc[page_num]
        
        # First pass: find all instances and create redaction annotations
        annotations = []
        for old_price_text, new_price_text in price_conversions:
            # Search for the old price text on this page
            text_instances = page.search_for(old_price_text)
            
            for inst in text_instances:
                annotations.append((inst, new_price_text))
        
        # Apply all redactions at once
        for rect, new_text in annotations:
            page.add_redact_annot(rect, fill=(1, 1, 1))  # White fill
        
        page.apply_redactions()
        
        # Second pass: insert new text
        for rect, new_text in annotations:
            page.insert_text(
                (rect.x0, rect.y0+9),  # Position
                new_text,            # New text
                fontsize=7.5,          # Adjust font size to match original
                color=(0, 0, 0)      # Black text
            )
    
    return True

def remove_amount_values_only(doc):
    """
    Remove only the numeric values following the 'Amount' text 
    (keep 'Amount' word and currency symbol)
    
    Args:
        doc: PyMuPDF document object
    
    Returns:
        bool: True if successful, False otherwise
    """
    # Extract text from all pages to find all amounts
    all_text = ""
    for page_num in range(len(doc)):
        all_text += doc[page_num].get_text("text") + "\n"
    
    # Regex pattern to find the numeric values after "Amount"
    # This captures only the numeric part (digits, commas, dots) before the currency
    amount_pattern = r"Amount\s*([\d.,]+)\s*([A-Z]{3})"
    matches = re.findall(amount_pattern, all_text, re.IGNORECASE)
    
    if not matches:
        print("No amounts found in the PDF")
        return False
    
    print(f"Found {len(matches)} amounts to process:")
    for numeric_part, currency in matches:
        print(f"  {numeric_part} {currency}")
    
    # Extract just the numeric parts to remove (remove duplicates)
    numeric_values_to_remove = list(set([match[0] for match in matches]))
    numeric_values_to_remove.sort(key=len, reverse=True)
    
    # Iterate through each page to find and remove only the numeric values
    for page_num in range(len(doc)):
        page = doc[page_num]
        
        redaction_rects = []
        for numeric_value in numeric_values_to_remove:
            # Search for the numeric value text on this page
            text_instances = page.search_for(numeric_value)
            redaction_rects.extend(text_instances)
        
        # Apply all redactions at once
        for rect in redaction_rects:
            page.add_redact_annot(rect, fill=(1, 1, 1))  # White fill
        
        if redaction_rects:
            page.apply_redactions()
            print(f"Removed {len(redaction_rects)} numeric values from page {page_num + 1}")
    
    return True

def remove_total_usd_value(doc):
    """
    Remove only the numerical value following the first 'Total USD' text
    
    Args:
        doc: PyMuPDF document object
    
    Returns:
        bool: True if successful, False otherwise
    """
    # Extract text from all pages to find the Total USD value
    all_text = ""
    for page_num in range(len(doc)):
        all_text += doc[page_num].get_text("text") + "\n"
    
    # Regex pattern to find the numeric value after "Total USD"
    total_pattern = r"Total USD\s*([\d.,]+)"
    match = re.search(total_pattern, all_text, re.IGNORECASE)
    
    if not match:
        print("Total USD not found in the PDF")
        return False
    
    numeric_value = match.group(1)
    print(f"Found Total USD value: {numeric_value}")
    
    # Iterate through each page to find and remove the numeric value
    found = False
    for page_num in range(len(doc)):
        if found:
            break  # Stop after finding and processing the first occurrence
            
        page = doc[page_num]
        
        # Search for the numeric value on this page
        text_instances = page.search_for(numeric_value)
        
        if text_instances:
            # Apply redaction to remove the numeric value
            for rect in text_instances:
                page.add_redact_annot(rect, fill=(1, 1, 1))  # White fill
            
            page.apply_redactions()
            print(f"Removed Total USD value '{numeric_value}' from page {page_num + 1}")
            found = True
    
    if not found:
        print("Could not locate the Total USD value on any page for removal")
        return False
    
    return True

def process_pdf(input_pdf, output_pdf):
    """
    Main function to process PDF - apply all modifications
    
    Args:
        input_pdf (str): Path to input PDF file
        output_pdf (str): Path to output PDF file
    
    Returns:
        bool: True if successful, False otherwise
    """
    # Open the PDF document
    doc = fitz.open(input_pdf)
    
    # Apply all modifications in sequence
    success1 = replace_shipment_date(doc)
    success2 = replace_price(doc)
    success3 = remove_amount_values_only(doc)
    success4 = remove_total_usd_value(doc)
    
    # Check if all operations were successful
    all_success = success1 and success2 and success3 and success4
    
    if all_success:
        print(f"Successfully processed PDF with all modifications. Saved as: {output_pdf}")
        # Save the document
        doc.save(output_pdf)
        doc.close()
        return True
    else:
        # Identify which operations failed
        failed_operations = []
        if not success1:
            failed_operations.append("shipment date replacement")
        if not success2:
            failed_operations.append("price replacement")
        if not success3:
            failed_operations.append("amount values removal")
        if not success4:
            failed_operations.append("total USD value removal")
        
        doc.close()
        error_message = f"Failed to process PDF: The following operations were not successful - {', '.join(failed_operations)}"
        raise Exception(error_message)

if __name__ == "__main__":
    # Example usage
    input_file = r"C:\Users\Altersense\Desktop\ERP-RPA\Sample\Other Samples\PO_CG-05B41981-NORMA AOP polo_250527 - Copy.pdf"
    output_file = r"C:\Users\Altersense\Desktop\ERP-RPA\Sample\Other Samples\PO_CG-05B41981-NORMA AOP polo_250527 - Copy_fully_processed.pdf"
    
    # Process PDF with all modifications
    process_pdf(input_file, output_file)