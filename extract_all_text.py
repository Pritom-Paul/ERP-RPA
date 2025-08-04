import os
import re
import pdfplumber
import pandas as pd

pdf_dir = r"C:\Users\Altersense\Desktop\ERP-RPA\Sample"

for filename in os.listdir(pdf_dir):
    if filename.lower().endswith(".pdf"):
        pdf_path = os.path.join(pdf_dir, filename)
        with pdfplumber.open(pdf_path) as pdf:
            all_text = ""
            for page in pdf.pages:
                all_text += page.extract_text() + "\n" or ""
        
        ## COMMON FIELDS EXTRACTION ## 
        # PO Number extraction
        po_match = re.search(r"Purchase Order\s+(PO\d+)", all_text, re.IGNORECASE)
        po_no = po_match.group(1) if po_match else None

        # Extract Buying House (first non-empty line)
        lines = all_text.splitlines()
        buying_house = next((line.strip() for line in lines if line.strip()), None)

        # Buying House Address: from line 2 until "Purchase Order" line
        buying_house_add_lines = []
        for line in lines[1:]:
            if "purchase order" in line.lower():
                break
            buying_house_add_lines.append(line)
        buying_house_add = "\n".join(buying_house_add_lines) if buying_house_add_lines else None

        # Ship-to Address
        ship_to_address = None
        for idx, line in enumerate(lines):
            if "ship-to address" in line.lower():
                # Get next non-empty line
                for next_line in lines[idx + 1:]:
                    if next_line.strip():
                        ship_to_address = next_line.strip()
                        break
                break

        # Vendor No
        vendor_no_match = re.search(r"Vendor No.\s+(\d+)", all_text, re.IGNORECASE)
        vendor_no = vendor_no_match.group(1) if vendor_no_match else None

        #Payment Terms
        payment_terms_match = re.search(r"Payment Terms\s+(.+?)(?=\s+Prices)", all_text, re.IGNORECASE)
        payment_terms = payment_terms_match.group(1) if payment_terms_match else None

        #Document Date
        document_date_match = re.search(r"Document Date\s+(\d{2}[-/]\d{2}[-/]\d{2,4})", all_text, re.IGNORECASE)        
        document_date = document_date_match.group(1) if document_date_match else None

        # Shipment Method
        shipment_method_match = re.search(r"Shipment Method\s+(.+?)(?=\s+Order Type)", all_text, re.IGNORECASE)
        shipment_method = shipment_method_match.group(1) if shipment_method_match else ""

        # Order Type
        order_type_match = re.search(r"Order Type\s+(.*)", all_text, re.IGNORECASE)
        order_type = order_type_match.group(1).strip() if order_type_match else None

        #Shipping Agent
        shipping_agent_match = re.search(r"Shipping Agent\s+(.*)", all_text, re.IGNORECASE)
        shipping_agent = shipping_agent_match.group(1).strip() if shipping_agent_match else None
        
        #Order For
        order_for_match = re.search(r"Order For\s+(.*)", all_text, re.IGNORECASE)
        order_for = order_for_match.group(1).strip() if order_for_match else None

        #Style No
        style_no = None
        for i, line in enumerate(lines):
            if re.search(r"Style No\.", line, re.IGNORECASE):
                if i + 1 < len(lines):
                    next_line = lines[i + 1].strip()
                    style_no = next_line.split()[0] if next_line else None
                break

        # Description
        description = None
        for i, line in enumerate(lines):
            if re.search(r"Style No\.", line, re.IGNORECASE):
                if i + 1 < len(lines):
                    next_line = lines[i + 1].strip()
                    words = next_line.split()
                    if len(words) > 1:
                        description = " ".join(words[1:])
                    else:
                        description = None
                break
        
        # HS Code
        hs_code_match = re.search(r"HS CODE:\s+(\d+)", all_text, re.IGNORECASE)
        hs_code = hs_code_match.group(1) if hs_code_match else None

        #Total Quantity
        total_qty_match = re.search(r"Total Qty.\s+(\d+)", all_text, re.IGNORECASE)
        total_qty = total_qty_match.group(1) if total_qty_match else None

        #Prices Including VAT
        prices_including_vat_match = re.search(r"Prices Including VAT\s+(\w+)", all_text, re.IGNORECASE)
        prices_including_vat = prices_including_vat_match.group(1) if prices_including_vat_match else None


        ## GROUP FIELDS ##
        # Extract group fields: Color Code + Description, Size, Quantity
        group_entries = []
        i = 0
        while i < len(lines):
            line = lines[i].strip()

            # Skip HS Code lines
            if "HS Code:" in line:
                i += 1
                continue

            # Detect a new product block (starts with Style No.)
            if re.match(r"^\d{6}\b", line):  # e.g., "101423 USPA T-Shirt Arjun Men"
                i += 1  # Skip the Style No. line
                
                if i >= len(lines):
                    break

                # Extract sizes (next line)
                size_line = lines[i].strip()
                sizes = size_line.split()  # e.g., ["XXS", "XS", "S", "M", ...]
                i += 1

                if i >= len(lines):
                    break

                # Get the quantity line (contains color code + quantities + batch qty)
                qty_line = lines[i].strip()

                # Handle Greymelange case (no color code)
                greymelange_match = re.match(r"^([A-Za-z]+)\s+([\d\s]+)$", qty_line)
                if greymelange_match:
                    color_desc = greymelange_match.group(1)  # "Greymelange"
                    qty_parts = greymelange_match.group(2).split()
                    
                    # The last number is the batch quantity
                    batch_qty = int(qty_parts[-1]) if qty_parts else 0
                    qtys = [int(q) for q in qty_parts[:-1]]  # Individual quantities
                    
                    for sz, qty in zip(sizes, qtys):
                        group_entries.append((color_desc, sz, qty, batch_qty))
                    i += 1
                    continue

                # Normal case (with color code)
                color_code_match = re.match(r"^([A-Za-z0-9\-]+)", qty_line)
                if color_code_match:
                    color_code = color_code_match.group(1)  # e.g., "11-0601TCX"
                    qty_parts = qty_line.split()
                    
                    # The last number is the batch quantity
                    batch_qty = int(qty_parts[-1]) if qty_parts else 0
                    qtys = [int(q) for q in qty_parts[1:-1] if q.isdigit()]  # Skip color code & batch qty
                    
                    # Get color description from next line
                    if i + 1 < len(lines):
                        color_desc = lines[i+1].strip()  # e.g., "White"
                        i += 1
                    
                    color_key = f"{color_code}\n{color_desc}"
                    
                    for sz, qty in zip(sizes, qtys):
                        group_entries.append((color_key, sz, qty, batch_qty))
                    i += 1
                else:
                    i += 1
            else:
                i += 1
        

        #BARCODES
        barcodes = []
        barcode_started = False
        for line in lines:
            stripped_line = line.strip()
            if not barcode_started:
                if "barcode" in stripped_line.lower():
                    barcode_started = True
                continue  # don't process this line yet
            if not stripped_line:
                break  # stop on empty line
            match = re.search(r"(\d{13})\s*$", stripped_line)
            if match:
                barcodes.append(match.group(1))
            else:
                break  # stop if it doesn't end with 13 digits


        # Print the extracted information
        # print(f"--- File: {filename} ---")
        # print(f"-----All text:\n{all_text}\n-----")
        # print(f"\nColor Code + Description\tSize\tQuantity")
        # for entry in group_entries:
        #     print(f"{entry[0]}\t{entry[1]}\t{entry[2]}")
        # print(f"PO Number: {po_no}")
        # print(f"Buying House: {buying_house}")
        # print(f"Buying House Address:\n{buying_house_add}")
        # print(f"Ship-to Address: {ship_to_address}")
        # print(f"Vendor No: {vendor_no}")
        # print(f"Payment Terms: {payment_terms}")
        # print(f"Document Date: {document_date}")
        # print(f"Shipment Method: {shipment_method}")
        # print(f"Order Type: {order_type}")
        # print(f"Shipping Agent: {shipping_agent}")
        # print(f"Order For: {order_for}")
        # print(f"Style No: {style_no}")
        # print(f"Description: {description}")
        # print(f"HS Code: {hs_code}")
        # # Print the extracted group entries
        # print(f"Barcodes: {barcodes}")
        # print(f"Total Quantity: {total_qty}")

        
        # print("="*80)

        # Make sure barcode count matches group entries
        if len(barcodes) != len(group_entries):
            print(f"⚠️ Warning: Number of barcodes ({len(barcodes)}) does not match group entries ({len(group_entries)}).")


        # # Create dataframe rows
        rows = []
        for (color_desc, size, qty, batch_qty), barcode in zip(group_entries, barcodes):
            rows.append({
                "PO Number": po_no,
                "Buying House": buying_house,
                "Buying House Address": buying_house_add,
                "Ship-to Address": ship_to_address,
                "Vendor No": vendor_no,
                "Payment Terms": payment_terms,
                "Document Date": document_date,
                "Shipment Method": shipment_method,
                "Order Type": order_type,
                "Shipping Agent": shipping_agent,
                "Order For": order_for,
                "Style No": style_no,
                "Description": description,
                "HS Code": hs_code,
                "Color Code + Description": color_desc,
                "Size": size,
                "Quantity": qty,
                "Barcode": barcode,
                "Prepack Code": "",
                "Prepacks":"",
                "Total Quantity": total_qty,
                "Batch Quantity": batch_qty,
                "Prices Including VAT": prices_including_vat
            })

        # Create DataFrame
        df = pd.DataFrame(rows)

        # Print the DataFrame (optional)
        print("\nFinal DataFrame:")
        print(df.to_string(index=False))

        # Save to Excel
        output_file = os.path.join(pdf_dir, f"{os.path.splitext(filename)[0]}.xlsx")
        df.to_excel(output_file, index=False)
