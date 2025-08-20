import os
import re
import pdfplumber
import pandas as pd
import camelot
from babel.numbers import parse_decimal
from pypdf import PdfReader

pdf_dir = r"C:\Users\Altersense\Desktop\ERP-RPA\Sample\Other Samples"
# pdf_dir = r"C:\Users\Altersense\Desktop\ERP-RPA\Sample\Main"

#PYPDF HELPER FUNCTIONS
# === Helpers for size/qty extraction ===
def extract_pdf_text_preserve_layout(pdf_path):
    reader = PdfReader(pdf_path)
    pages_text = {}
    for page_num, page in enumerate(reader.pages, start=1):
        page_text = page.extract_text(
            extraction_mode="layout",
            keep_blank_chars=True,
        )
        if page_text:
            pages_text[page_num] = page_text
    return pages_text


def tokens_with_cols(line):
    """Return tokens with their starting column positions"""
    return [(m.group(), m.start()) for m in re.finditer(r"\S+", line)]


def extract_size_qty_mappings(text):
    """Extracts size-qty tables from text block"""
    pattern = r"(Size[^\n]+)\n\s*(Qty[^\n]+)"
    matches_iter = re.finditer(pattern, text)
    mappings = []

    for m in matches_iter:
        size_line = m.group(1)
        qty_line = m.group(2)

        size_cols = [(tok, col) for tok, col in tokens_with_cols(size_line) if tok.lower() != "size"]
        qty_cols = [(tok, col) for tok, col in tokens_with_cols(qty_line) if tok.lower() != "qty"]

        mapping = {s: 0 for s, _ in size_cols}
        size_tokens = [s for s, _ in size_cols]
        qty_tokens = [q for q, _ in qty_cols]

        if len(size_tokens) == len(qty_tokens):  # direct mapping
            for s_tok, q_tok in zip(size_tokens, qty_tokens):
                norm = q_tok.replace('.', '').replace(',', '')
                mapping[s_tok] = int(norm) if norm.isdigit() else 0
        else:  # align by column
            size_ranges = []
            for i, (s_tok, s_col) in enumerate(size_cols):
                s_start = -9999 if i == 0 else (size_cols[i - 1][1] + s_col) // 2
                s_end = (s_col + size_cols[i + 1][1]) // 2 if i + 1 < len(size_cols) else 9999
                size_ranges.append((s_tok, s_start, s_end))

            for tok, q_col in qty_cols:
                norm = tok.replace('.', '').replace(',', '')
                if norm.isdigit():
                    q_val = int(norm)
                    for s_tok, s_start, s_end in size_ranges:
                        if s_start <= q_col < s_end:
                            mapping[s_tok] += q_val
                            break

        mappings.append(mapping)

    return mappings

#* MAIN PROCESSING LOOP
for filename in os.listdir(pdf_dir):
    if filename.lower().endswith(".pdf"):
        pdf_path = os.path.join(pdf_dir, filename)

        with pdfplumber.open(pdf_path) as pdf:
            all_text = ""           # Raw extracted text from all pages
            all_pages_body = []     # Store processed body text per page
            # preserved_layout_text = ""  # Text with preserved layout (spaces, etc.)

            # Header pattern like "*05B00041981*"
            header_pattern = re.compile(r'^\*\d{2}B\d+\*$')

            # === PAGE LOOP ===
            for page in pdf.pages:
                # Append raw text to all_text (full PDF content, unprocessed)
                page_text = page.extract_text() or ""
                all_text += page_text + "\n"
                # # Use preserved_text with extra parameters to preserve spacing
                # preserved_text = "" 
                # preserved_text += page.extract_text(
                #     layout=True,      # Preserves layout including spaces
                #     keep_blank_chars=True,  # Keeps blank characters
                #     x_tolerance=1,    # Smaller tolerance preserves spacing better
                #     y_tolerance=1
                # ) or ""
                # preserved_layout_text += preserved_text + "\n"
                # # print(f"-----All text:\n{all_text}\n-----")
                # print(f"-----Preserved Layout Text:\n{preserved_layout_text}\n-----")

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
            last_total_usd_index = None
            lines = body_text.split("\n")

            # Pre-calc last "Total USD" if "Applicable Certifications" not found
            if "Applicable Certifications" not in body_text:
                for i, line in enumerate(lines):
                    if "Total USD" in line:
                        last_total_usd_index = i

            for i, line in enumerate(lines):
                if "This document contains certified products. Please see table below for details." in line:
                    capture = True
                    continue
                
                if not capture and "shipment FOB date" in line:
                    capture = True

                if "Applicable Certifications" in line:
                    capture = False
                    continue

                # Stop continuing if we hit the last "Total USD" line
                if last_total_usd_index is not None and i > last_total_usd_index:
                    capture = False

                if capture and line.strip():
                    order_details_lines.append(line)

            order_details_text = "\n".join(order_details_lines)

        # === Extract sizes & quantities with PyPDF ===
        pages_text = extract_pdf_text_preserve_layout(pdf_path)

        sizes_list = []
        quantities_list = []

        for page_num, text in pages_text.items():
            mappings = extract_size_qty_mappings(text)
            for mapping in mappings:
                for size, qty in mapping.items():
                    sizes_list.append(size)
                    quantities_list.append(str(qty))

        # print(len(sizes_list), len(quantities_list))
        if len(sizes_list) != len(quantities_list):
            # print(f"Warning: Sizes and Quantities lists have different lengths! Sizes: {len(sizes_list)}, Quantities: {len(quantities_list)}")
            # Handle the mismatch by raising an error
            raise ValueError("Sizes and Quantities lists have different lengths!")
        # else:
            # print(f"Sizes: {sizes_list}, Quantities: {quantities_list}")
            # print("LENGTH:",len(sizes_list))


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
        ship_to_address = 'Not found'
        agency_to_address = "Not found"
        buying_house_address = "Not found"
        agency_full_address = "Not found"

        found_address = False

        for i, df in enumerate(camelot_tables):
            if df.shape[1] >= 2:
                first_cell = str(df.iloc[0, 0])
                if "Ship-to Address" in first_cell:
                    ship_to_address = df.iloc[1, 0].strip()
                    agency_to_address = df.iloc[1, 1].strip()
                    buying_house_address = "\n".join(
                        [row[0].strip() for row in df.iloc[2:].values if row[0].strip()]
                    )
                    agency_full_address = "\n".join(
                        [row[1].strip() for row in df.iloc[2:].values if row[1].strip()]
                    )
                    found_address = True
                    break

        # Fallback search if not found in standard position
        if not found_address:
            for i, df in enumerate(camelot_tables):
                for row_idx in range(df.shape[0]):
                    for col_idx in range(df.shape[1]):
                        cell_value = str(df.iloc[row_idx, col_idx]).strip()
                        if "Ship-to Address" in cell_value:
                            # print(f"Found 'Ship-to Address' at Table {i+1}, Row {row_idx+1}, Column {col_idx+1}")
                            
                            # Extract ship-to address (next row in same column)
                            if row_idx + 1 < df.shape[0]:
                                ship_to_address = str(df.iloc[row_idx + 1, col_idx]).strip()
                                
                                # Extract buying house address (next 3 rows max)
                                buying_house_lines = []
                                for r in range(row_idx + 2, df.shape[0]):
                                    line = str(df.iloc[r, col_idx]).strip()
                                    if line: buying_house_lines.append(line)
                                buying_house_address = "\n".join(buying_house_lines)
                            
                            # Try to find agency in same row
                            for agency_col in range(df.shape[1]):
                                if agency_col != col_idx and "Agency" in str(df.iloc[row_idx, agency_col]).strip():
                                    agency_to_address = str(df.iloc[row_idx + 1, agency_col]).strip() if row_idx + 1 < df.shape[0] else "Not found"
                                    
                                    # Extract agency address (next 3 rows max)
                                    agency_lines = []
                                    for r in range(row_idx + 2, df.shape[0]):
                                        line = str(df.iloc[r, agency_col]).strip()
                                        if line: agency_lines.append(line)
                                    agency_full_address = "\n".join(agency_lines)
                                    break
                            
                            found_address = True
                            break
                    if found_address: break
                if found_address: break

        # print("\nFinal Address Values:")
        # print(f"Ship-to: {ship_to_address}")
        # print(f"Agency: {agency_to_address}")
        # print(f"Buying House: {buying_house_address}")
        # print(f"Full Agency: {agency_full_address}")


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

        # Also capture the style description line above each certification
        cert_map = {}
        cert_pattern = re.finditer(r'([^\n]+)\nCertifications:\s*(.+)', all_text)
        for match in cert_pattern:
            style_desc_full = match.group(1).strip()
            style_desc_parts = style_desc_full.split()
            style_desc = " ".join(style_desc_parts[1:]) if style_desc_parts else style_desc_full
            certs = match.group(2).strip()
            cert_map[style_desc] = certs
        # Capture each order block: from the first line until the line starting with Amount
        order_blocks = re.findall(r'^(.*?Amount\s[\d\.,]+\s*\w*)', order_details_text, re.DOTALL | re.MULTILINE)

        orders = []

        # Check if we have matching certifications before processing
        if len(applicable_certifications) != len(order_blocks):
            # print(f"Warning: Length mismatch! Order blocks: {len(order_blocks)}, Certifications: {len(applicable_certifications)}")
            use_cert_map = True
        else:
            use_cert_map = False
        for block_idx, block in enumerate(order_blocks):
            block = block.strip()
            lines = block.splitlines()
            # print(f"Processing block:\n{block}\n")

            # Extract first word as Style No
            style_no = lines[0].split()[0] if lines else "Not found"
            
            # Extract Amount value
            amount_match = re.search(r'Amount\s([\d\.,]+)', block)
            amount_usd = amount_match.group(1) if amount_match else "Not found"
            # Convert European format to US float
            amount = str(parse_decimal(amount_usd, locale='de_DE'))
            

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
            
            # Get certification
            if not use_cert_map:
                current_certification = applicable_certifications[block_idx] if block_idx < len(applicable_certifications) else "Not found"
                # print(f"[DEBUG] Block {block_idx}: Using index-based cert → {current_certification}")
            else:
                current_certification = cert_map.get(style_description, "Not found")
                # print(f"[DEBUG] Block {block_idx}: Style Description = '{style_description}' | Mapped Cert = '{current_certification}'")

                # if current_certification == "Not found":
                #     # fallback: partial match
                #     for desc, certs in cert_map.items():
                #         if desc in style_description:
                #             current_certification = certs
                #             print(f"For Block {block_idx}: Fallback partial match → '{desc}' → '{certs}'")
                #             break
                # print(f"[DEBUG] Final mapping for Block {block_idx}: Style No = {style_no}, Style Desc = '{style_description}', Certification = '{current_certification}'")
            
            # Get Size length for this block(For Loop count for the block)
            size_line = ""
            for line in lines:
                if line.startswith("Size"):
                    size_line = line
            if size_line:
                # Extract sizes (skip the first word "Size")
                sizes = size_line.split()[1:]

            # Store as dict - one row per size
            # for size, qty in zip(sizes_list, quantities_list):
            for i in range(len(sizes)):
                orders.append({
                    'Order Number': order_number,
                    'Buy-from Vendor No': buy_from_vendor_no,
                    'Order Date': order_date,
                    'Purchaser': purchaser,
                    'Email': email,
                    'Phone No': phone_no,
                    'Payment Terms': payment_terms,
                    'Payment Method': payment_method,
                    'Shipment Method': shipment_method,
                    'Transport Method': transport_method,
                    'Shipping Agent Code': shipping_agent_code,
                    'Port of Departure': port_of_departure,
                    'Shipment Date ETD': shipment_date_etd,
                    'Cost Center': cost_center,
                    'Ship-to Address': ship_to_address,
                    'Buying House Address': buying_house_address,
                    'Agency to Address': agency_to_address,
                    'Agency to Address Details': agency_full_address,
                    'Unit': unit,
                    'Currency': currency,
                    'Style No': style_no,
                    'Style Description': style_description,
                    'Applicable Certifications': current_certification,
                    'Customs No': customs_no,
                    'Shipment FOB Date': shipment_fob_date,
                    'Color': color,
                    'Size': None,
                    'Qty': None,
                    'Price': price,
                    'Total Quantity': total_quantity,
                    'Amount': amount,
                    "Total Pieces": total_pieces,
                    "Total USD": total_usd,
                })

        # Convert to DataFrame
        orders_df = pd.DataFrame(orders)
        # Check lengths match
        if len(orders_df) != len(sizes_list) or len(orders_df) != len(quantities_list):
            raise ValueError(f"Length mismatch: orders({len(orders_df)}), sizes({len(sizes_list)}), quantities({len(quantities_list)})")

        if sizes_list and quantities_list:
            # Ensure sizes_list and quantities_list are the same length as orders_df
            if len(sizes_list) != len(orders_df) or len(quantities_list) != len(orders_df):
                raise ValueError(f"Sizes and Quantities lists do not match orders DataFrame length: {len(sizes_list)} vs {len(orders_df)}, {len(quantities_list)} vs {len(orders_df)}")
            # Assign the global size and quantity lists to the pre-initialized columns
            orders_df['Size'] = sizes_list
            orders_df['Qty'] = quantities_list
        
        if orders_df.empty:
            raise ValueError(f"No valid order data extracted from {filename}")
        else:
            print(f"Extracted Orders DataFrame:\n{orders_df}")
            # # Save to Excel
            output_excel = os.path.join(pdf_dir, f"{os.path.splitext(filename)[0]}_orders.xlsx")
            orders_df.to_excel(output_excel, index=False)
            print(f"Extracted orders saved to {output_excel}")

