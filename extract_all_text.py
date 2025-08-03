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
        po_no = po_match.group(1) if po_match else "NOT FOUND"

        # Extract Buying House (first non-empty line)
        lines = all_text.splitlines()
        buying_house = next((line.strip() for line in lines if line.strip()), "NOT FOUND")

        # Buying House Address: from line 2 until "Purchase Order" line
        buying_house_add_lines = []
        for line in lines[1:]:
            if "purchase order" in line.lower():
                break
            buying_house_add_lines.append(line)
        buying_house_add = "\n".join(buying_house_add_lines) if buying_house_add_lines else "NOT FOUND"

        print(f"--- File: {filename} ---")
        print(f"-----All text:\n{all_text}\n-----")
        print(f"PO Number: {po_no}")
        print(f"Buying House: {buying_house}")
        print(f"Buying House Address:\n{buying_house_add}")
        print("="*80)
