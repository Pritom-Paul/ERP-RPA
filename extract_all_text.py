import os
import re
import pdfplumber

pdf_dir = r"C:\Users\Altersense\Desktop\ERP-RPA\Sample"

for filename in os.listdir(pdf_dir):
    if filename.lower().endswith(".pdf"):
        pdf_path = os.path.join(pdf_dir, filename)
        with pdfplumber.open(pdf_path) as pdf:
            all_text = ""
            for page in pdf.pages:
                all_text += page.extract_text() + "\n" or ""
        
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


        print(f"--- File: {filename} ---")
        print(f"-----All text:\n{all_text}\n-----")
        print(f"PO Number: {po_no}")
        print(f"Buying House: {buying_house}")
        print(f"Buying House Address:\n{buying_house_add}")
        print(f"Ship-to Address: {ship_to_address}")
        print(f"Vendor No: {vendor_no}")
        print(f"Payment Terms: {payment_terms}")
        print(f"Document Date: {document_date}")
        print(f"Shipment Method: {shipment_method}")
        print(f"Order Type: {order_type}")
        print(f"Shipping Agent: {shipping_agent}")
        print(f"Order For: {order_for}")
        print(f"Style No: {style_no}")
        print(f"Description: {description}")
        print(f"HS Code: {hs_code}")

        
        print("="*80)
