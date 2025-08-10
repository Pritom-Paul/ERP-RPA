import os
import re
import pdfplumber
import pandas as pd

pdf_dir = r"C:\Users\Altersense\Desktop\ERP-RPA\Format 2 Test"

for filename in os.listdir(pdf_dir):
    if filename.lower().endswith(".pdf"):
        pdf_path = os.path.join(pdf_dir, filename)
        with pdfplumber.open(pdf_path) as pdf:
            all_text = ""
            body_text_lines = []
            last_header_line = None
            
            # First pass to find the last header line from first page
            first_page = pdf.pages[0]
            first_page_text = first_page.extract_text() or ""
            first_page_lines = first_page_text.splitlines()
            
            for idx, line in enumerate(first_page_lines):
                if "order for" in line.lower():
                    if idx > 0:
                        last_header_line = first_page_lines[idx-1].strip().lower()
                    break
            
            if not last_header_line:
                raise ValueError("Invalid file format: Header not found.")
            

            # Process all pages using the found last header line
            for page in pdf.pages:
                page_text = page.extract_text() or ""
                all_text += page_text + "\n"
                page_lines = page_text.splitlines()
                found_header_in_page = False
                if last_header_line:
                    for idx, line in enumerate(page_lines):
                        if last_header_line in line.strip().lower():
                            found_header_in_page = True
                            if idx + 1 < len(page_lines):
                                body_text_lines.extend(page_lines[idx + 1:])
                            break
                if not found_header_in_page:
                    raise ValueError("Invalid file format: Header missing on one or more pages.")
        # print(f"last_header_line: {last_header_line}")
        
        # Extract group fields: Color Code + Description, Size, Quantity (from body text)
        group_entries = []
        i = 0
        while i < len(body_text_lines):
            line = body_text_lines[i].strip()

            # Skip HS Code lines
            if "HS Code:" in line:
                i += 1
                continue

            # Detect a new product block (starts with Style No.)
            if re.match(r"^\d{6}\b", line):  # e.g., "101423 USPA T-Shirt Arjun Men"
                i += 1  # Skip the Style No. line
                
                if i >= len(body_text_lines):
                    break

                # Extract sizes (next line)
                size_line = body_text_lines[i].strip()
                sizes = size_line.split()  # e.g., ["XXS", "XS", "S", "M", ...]
                i += 1

                if i >= len(body_text_lines):
                    break

                # Get the quantity line (contains color code + quantities + batch qty)
                qty_line = body_text_lines[i].strip()

                # Handle Greymelange case (no color code)
                greymelange_match = re.match(r"^([A-Za-z]+)\s+([\d\s]+)$", qty_line)
                if greymelange_match:
                    color_desc = greymelange_match.group(1)  # "Greymelange"
                    qty_parts = greymelange_match.group(2).split()
                    
                    # The last value is the batch quantity
                    batch_qty = qty_parts[-1] if qty_parts else ""
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
                    
                    # The last value is the batch quantity
                    batch_qty = qty_parts[-1] if qty_parts else ""
                    qtys = [int(q) for q in qty_parts[1:-1] if q.isdigit()]  # Skip color code & batch qty
                    
                    if i + 1 < len(body_text_lines):
                        color_desc = body_text_lines[i+1].strip()
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
        for line in body_text_lines:
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


        # # Print the extracted information
        # print(f"--- File: {filename} ---")
        # print(f"\nColor Code + Description\tSize\tQuantity")
        # for entry in group_entries:
        #     print(f"{entry[0]}\t{entry[1]}\t{entry[2]}")
        # print(f"Barcodes: {barcodes}")
        # print("="*80)
        # Make sure barcode count matches group entries
        if len(barcodes) != len(group_entries):
            print(f"⚠️ Warning: For the file: {filename} Number of barcodes ({len(barcodes)}) does not match group entries ({len(group_entries)}).")
            # print(f"-----All text:\n{all_text}\n-----")
            print(f"body_text_lines: {body_text_lines}")
        else:
            print(f"✅ Number of barcodes matches group entries.")
