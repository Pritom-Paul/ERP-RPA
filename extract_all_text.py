import os
import re
import pdfplumber
import pandas as pd
import camelot
from babel.numbers import parse_decimal

pdf_dir = r"C:\Users\Altersense\Desktop\ERP-RPA\Sample\Main"

for filename in os.listdir(pdf_dir):
    if filename.lower().endswith(".pdf"):
        pdf_path = os.path.join(pdf_dir, filename)

        with pdfplumber.open(pdf_path) as pdf:
            all_text = ""           # Raw extracted text from all pages
            all_pages_body = []     # Store processed body text per page

            # Header pattern like "*05B00041981*"
            header_pattern = re.compile(r'^\*\d{2}B\d+\*$')

            # === PAGE LOOP ===
            for page in pdf.pages:
                # Append raw text to all_text (full PDF content, unprocessed)
                page_text = page.extract_text() or ""
                all_text += page_text + "\n"

                # Clean body text for this page
                lines = page_text.split('\n')
                in_header = False
                page_body_lines = []

                for line in lines:
                    if header_pattern.match(line.strip()):
                        in_header = True
                        continue
                    if in_header and line.strip().startswith('Order '):
                        in_header = False
                        continue
                    if re.match(r'^Page \d+ of \d+$', line.strip()):
                        continue
                    page_body_lines.append(line)

                cleaned_body = '\n'.join(page_body_lines).strip()
                all_pages_body.append(cleaned_body)

            # === Combine cleaned page texts ===
            body_text = "\n".join(all_pages_body)

            # === Camelot table extraction (run ONCE per file) ===
            tables = camelot.read_pdf(pdf_path, pages='all', flavor='stream')
            camelot_tables = [table.df for table in tables]

            # === Extract order details ===
            order_details_lines = []
            capture = False

            for line in body_text.split("\n"):
                if "This document contains certified products. Please see table below for details." in line:
                    capture = True
                    continue
                if "Applicable Certifications" in line:
                    capture = False
                    continue
                if capture and line.strip():
                    order_details_lines.append(line)

            order_details_text = "\n".join(order_details_lines)

            # print(f"All Text:\n{all_text}")  # Raw text
            # print(f"Body Text:\n{body_text}")  # Cleaned text without headers/footers
            # print(f"Order Details:\n{order_details_text}")  # Extracted order details texts             



#* COMMON/SHARED FIELDS

extract_order_number = re.search(r"Order\s*([A-Za-z0-9]+)", all_text)
# print(f"Extracted Order Number: {extract_order_number.group(1) if extract_order_number else 'Not found'}")
order_number = extract_order_number.group(1) if extract_order_number else "Not found"

extract_buy_from_vendor_no = re.search(r"Buy-from Vendor No\.?\s*([0-9]+)", all_text)
# print(f"Extracted Buy-from Vendor No.: {extract_buy_from_vendor_no.group(1) if extract_buy_from_vendor_no else 'Not found'}")
buy_from_vendor_no = extract_buy_from_vendor_no.group(1) if extract_buy_from_vendor_no else "Not found"

extract_order_date = re.search(r"Order Date\s*([0-9]{2}-[0-9]{2}-[0-9]{4})", all_text)
# print(f"Extracted Order Date: {extract_order_date.group(1) if extract_order_date else 'Not found'}")
order_date = extract_order_date.group(1) if extract_order_date else "Not found"

extract_purchaser = re.search(r"Purchaser\s*([^\n\d]+)", all_text)
# print(f"Extracted Purchaser: {extract_purchaser.group(1) if extract_purchaser else 'Not found'}")
purchaser = extract_purchaser.group(1).strip() if extract_purchaser else "Not found"

# Extracting Email
extract_email_match = re.search(r"E-Mail\s*([^\n@]+@)", all_text)
extract_email = None
if extract_email_match:
    email_prefix = extract_email_match.group(1).strip().rstrip('@')
    extract_email = email_prefix + "@brands-fashion.com"
# print(f"Extracted Email: {extract_email if extract_email else 'Not found'}")
email = extract_email if extract_email else "Not found"

extract_phone_no = re.search(r"Phone No\.?\s*([+\d\-\(\)\s]+)", all_text)
# print(f"Extracted Phone No.: {extract_phone_no.group(1) if extract_phone_no else 'Not found'}")
phone_no = extract_phone_no.group(1) if extract_phone_no else "Not found"

extract_payment_terms = re.search(r"Payment Terms\s*([^\n]+)", all_text)
# print(f"Extracted Payment Terms: {extract_payment_terms.group(1) if extract_payment_terms else 'Not found'}")
payment_terms = extract_payment_terms.group(1) if extract_payment_terms else "Not found"

extract_payment_method = re.search(r"Payment Method\s*([^\n]+)", all_text)
# print(f"Extracted Payment Method: {extract_payment_method.group(1) if extract_payment_method else 'Not found'}")
payment_method = extract_payment_method.group(1) if extract_payment_method else "Not found"

extract_shipment_method = re.search(r"Shipment Method\s*([^\n]+)", all_text)
# print(f"Extracted Shipment Method: {extract_shipment_method.group(1) if extract_shipment_method else 'Not found'}")
shipment_method = extract_shipment_method.group(1) if extract_shipment_method else "Not found"

extract_transport_method = re.search(r"Transport Method\s*([^\n]+)", all_text)
# print(f"Extracted Transport Method: {extract_transport_method.group(1) if extract_transport_method else 'Not found'}")
transport_method = extract_transport_method.group(1) if extract_transport_method else "Not found"

extract_shipping_agent_code = re.search(r"Shipping Agent Code\s*([^\n]+)", all_text)
# print(f"Extracted Shipping Agent Code: {extract_shipping_agent_code.group(1) if extract_shipping_agent_code else 'Not found'}")
shipping_agent_code = extract_shipping_agent_code.group(1) if extract_shipping_agent_code else "Not found"

extract_port_of_departure = re.search(r"Port of Departure\s*([^\n]+)", all_text)
# print(f"Extracted Port of Departure: {extract_port_of_departure.group(1) if extract_port_of_departure else 'Not found'}")
port_of_departure = extract_port_of_departure.group(1) if extract_port_of_departure else "Not found"

extract_shipment_date_etd = re.search(r"Shipment Date ETD\s*([0-9]{1,2}-[0-9]{2}-[0-9]{4})", all_text)
# print(f"Extracted Shipment Date ETD: {extract_shipment_date_etd.group(1) if extract_shipment_date_etd else 'Not found'}")
shipment_date_etd = extract_shipment_date_etd.group(1) if extract_shipment_date_etd else "Not found"

extract_cost_center = re.search(r"Cost Center\s*([^\n]+)", all_text)
# print(f"Extracted Cost Center: {extract_cost_center.group(1) if extract_cost_center else 'Not found'}")
cost_center = extract_cost_center.group(1) if extract_cost_center else "Not found"

extract_unit_match = re.search(r"Total\s*[\d\.,]+\s*([A-Za-z]+)", all_text)
# print(f"Extracted Unit: {unit if unit else 'Not found'}")
unit = extract_unit_match.group(1) if extract_unit_match else "Not found"

extract_currency_match = re.search(r"Amount\s*[\d\.,]+\s*([A-Za-z]+)", all_text)
# print(f"Extracted Currency: {currency if currency else 'Not found'}")
currency = extract_currency_match.group(1) if extract_currency_match else None

extract_total_pieces_match = re.search(r"Total PIECES\s*([\d\.,]+)", all_text)
if extract_total_pieces_match:
    total_pieces = str(parse_decimal(extract_total_pieces_match.group(1), locale='de_DE'))
    # print(f"Extracted Total Pieces: {total_pieces}")
else:
    total_pieces = "Not found"

extract_total_usd_match = re.search(r"Total USD\s*([\d\.,]+)", all_text)
if extract_total_usd_match:
    total_usd = str(parse_decimal(extract_total_usd_match.group(1), locale='de_DE'))
    # print(f"Extracted Total USD: {total_usd}")
else:
    total_usd = "Not found"

#* Camelot Extractions 
ship_to_address = None
agency_to_address = None
buying_house_address = None
agency_full_address = None

for df in camelot_tables:
    if df.shape[1] >= 2 and "Ship-to Address" in str(df.iloc[0, 0]):
        ship_to_address = df.iloc[1, 0].strip()
        agency_to_address = df.iloc[1, 1].strip()
        # Extract FULL addresses (all lines under headers)
        buying_house_address = "\n".join(
            [row[0].strip() for row in df.iloc[2:].values if row[0].strip()]
        )
        agency_full_address = "\n".join(
            [row[1].strip() for row in df.iloc[2:].values if row[1].strip()]
        )
        break  # Stop after first match

# print(f"Extracted Ship-to Address: {ship_to_address if ship_to_address else 'Not found'}")
# print(f"Extracted Agency to Address: {agency_to_address if agency_to_address else 'Not found'}")
# print(f"Extracted Buying House Address: {buying_house_address if buying_house_address else 'Not found'}")
# print(f"Extracted Agency Full Address: {agency_full_address if agency_full_address else 'Not found'}")


#* GROUPED FIELDS

# Applicable Certifications
applicable_certifications=[]
extract_applicable_certifications = re.findall(r"Certifications:\s*(.+)", all_text)
if extract_applicable_certifications:
    for cert in extract_applicable_certifications:
        cert = cert.strip()
        if cert:
            applicable_certifications.append(cert)
    # print(f"Extracted Applicable Certifications: {applicable_certifications}")

# Capture each order block: from the first line until the line starting with Amount
order_blocks = re.findall(r'^(.*?Amount\s[\d\.,]+\s*\w*)', order_details_text, re.DOTALL | re.MULTILINE)

orders = []

# Check if we have matching certifications before processing
if len(applicable_certifications) != len(order_blocks):
    print(f"Warning: Length mismatch! Order blocks: {len(order_blocks)}, Certifications: {len(applicable_certifications)}")
    raise ValueError("The number of order blocks and applicable certifications do not match")

for block_idx, block in enumerate(order_blocks):
    block = block.strip()
    lines = block.splitlines()
    # print(f"Processing block:\n{block}\n")

    # Get certification for this block (same for all sizes in this block)
    current_certification = applicable_certifications[block_idx] if block_idx < len(applicable_certifications) else "Not found"

    # Extract first word as Style No
    style_no = lines[0].split()[0] if lines else "Not found"
    
    # Extract Amount value
    amount_match = re.search(r'Amount\s([\d\.,]+)', block)
    amount = amount_match.group(1) if amount_match else "0"
    # Convert European format to US float
    amount_usd = str(parse_decimal(amount, locale='de_DE'))
    

    # Extract Customs no. and Color from the same line
    customs_line_match = re.search(r'([^\n]*?)\bCustoms\s+no\.?\s*:\s*([\d ]+)', block, re.IGNORECASE)
    if customs_line_match:
        color = customs_line_match.group(1).strip()
        customs_no = customs_line_match.group(2).strip()
    else:
        color = "Not found"
        customs_no = "Not found"


    #Shipment FOB Date
    shipment_fob_date_match = re.search(r'shipment\s+FOB\s+date\s+(\d{1,2}\.\s+\w+\s+\d{4})', block)
    if shipment_fob_date_match:
        shipment_fob_date = shipment_fob_date_match.group(1).strip()
    else:
        shipment_fob_date = "Not found"

    #Price
    price_match = re.search(r'Price\s*:?\s*([\d\.,]+)', block, re.IGNORECASE)
    if price_match:
        price_str = price_match.group(1).strip()
        price = str(parse_decimal(price_str, locale='de_DE'))
    else:
        price = "Not found"

    #Total Quantity
    total_quantity_match = re.search(r'Total\s([\d\.,]+)', block)
    if total_quantity_match:
        total_quantity_str = total_quantity_match.group(1).strip()
        total_quantity = str(parse_decimal(total_quantity_str, locale='de_DE'))
    else:
        total_quantity = "Not found"


    # Style Description
    style_description_parts = []

    # first line (remove style no and cut at "shipment FOB date")
    first_line = lines[0]
    first_line_words = first_line.split()[1:]  # skip style no
    first_line_str = " ".join(first_line_words)
    shipment_match = re.search(r'(.*?)\bshipment\s+FOB\s+date', first_line_str, re.IGNORECASE)
    if shipment_match:
        style_description_parts.append(shipment_match.group(1).strip())
    else:
        style_description_parts.append(first_line_str.strip())

    # subsequent lines until Customs no
    for line in lines[1:]:
        if re.search(r'Customs\s+no', line, re.IGNORECASE):
            break
        style_description_parts.append(line.strip())

    style_description = " ".join(part for part in style_description_parts if part)
    
    # Find the line starting with "Size" and the next line starting with "Qty"
    size_line = ""
    qty_line = ""

    for line in lines:
        if line.startswith("Size"):
            size_line = line
        elif line.startswith("Qty"):
            qty_line = line

    if size_line and qty_line:
        # Extract sizes (skip the first word "Size")
        sizes = size_line.split()[1:]
        
        # Extract quantities (skip the first word "Qty")
        quantities = qty_line.split()[1:]

    # Store as dict - one row per size
    for size, qty in zip(sizes, quantities):
        orders.append({
            'Style No': style_no,
            'Style Description': style_description,
            'Applicable Certifications': current_certification,
            'Customs No': customs_no,
            'Amount USD': amount_usd,
            'Shipment FOB Date': shipment_fob_date,
            'Color': color,
            'Total Quantity': total_quantity,
            'Price': price,
            'Size': size,
            'Qty': qty
        })

# Convert to DataFrame
orders_df = pd.DataFrame(orders)
print(f"Extracted Orders DataFrame:\n{orders_df}")

