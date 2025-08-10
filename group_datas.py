import os
import re
import pdfplumber

pdf_dir = r"C:\Users\Altersense\Desktop\ERP-RPA\Format 2 Test\test"

for filename in os.listdir(pdf_dir):
    if filename.lower().endswith(".pdf"):
        pdf_path = os.path.join(pdf_dir, filename)
        with pdfplumber.open(pdf_path) as pdf:
            # Find last header line on first page
            first_page_text = pdf.pages[0].extract_text() or ""
            first_page_lines = first_page_text.splitlines()
            last_header_line = None
            
            for idx, line in enumerate(first_page_lines):
                if "order for" in line.lower():
                    if idx > 0:
                        last_header_line = first_page_lines[idx-1].strip().lower()
                    break

            if not last_header_line:
                print(f"File {filename}: Header not found. Skipping.")
                continue
            
            # Collect body lines after last header line from all pages
            body_text_lines = []
            for page in pdf.pages:
                page_text = page.extract_text() or ""
                page_lines = page_text.splitlines()
                found_header = False
                for idx, line in enumerate(page_lines):
                    if last_header_line in line.strip().lower():
                        found_header = True
                        body_text_lines.extend(page_lines[idx + 1:])
                        break
                if not found_header:
                    print(f"File {filename}: Header missing on page. Skipping.")
                    body_text_lines = []
                    break

            if not body_text_lines:
                continue

            # Extract group entries
            group_entries = []
            i = 0
            while i < len(body_text_lines):
                line = body_text_lines[i].strip()

                # Skip HS Code lines
                if "HS Code:" in line:
                    i += 1
                    continue

                # Detect product block start (style no.)
                if re.match(r"^\d{6}\b", line):
                    i += 1
                    if i >= len(body_text_lines):
                        break

                    sizes_line = body_text_lines[i].strip()
                    sizes = sizes_line.split()
                    i += 1
                    if i >= len(body_text_lines):
                        break

                    qty_line = body_text_lines[i].strip()

                    # Greymelange case (no color code)
                    greymelange_match = re.match(r"^([A-Za-z]+)\s+([\d\s]+)$", qty_line)
                    if greymelange_match:
                        color_desc = greymelange_match.group(1)
                        qty_parts = greymelange_match.group(2).split()
                        batch_qty = qty_parts[-1] if qty_parts else ""
                        qtys = [int(q) for q in qty_parts[:-1]]
                        for sz, qty in zip(sizes, qtys):
                            group_entries.append((color_desc, sz, qty, batch_qty))
                        i += 1
                        continue

                    # Normal color code case
                    color_code_match = re.match(r"^([A-Za-z0-9\-]+)", qty_line)
                    if color_code_match:
                        color_code = color_code_match.group(1)
                        qty_parts = qty_line.split()
                        batch_qty = qty_parts[-1] if qty_parts else ""
                        qtys = [int(q) for q in qty_parts[1:-1] if q.isdigit()]

                        if i + 1 < len(body_text_lines):
                            color_desc = body_text_lines[i + 1].strip()
                            i += 1

                        color_key = f"{color_code}\n{color_desc}"

                        for sz, qty in zip(sizes, qtys):
                            group_entries.append((color_key, sz, qty, batch_qty))
                        i += 1
                    else:
                        i += 1
                else:
                    i += 1

            # Print extracted group data only for debug
            print(f"--- File: {filename} ---")
            print("Color Code + Description\tSize\tQuantity\tBatch Qty")
            for entry in group_entries:
                print(f"{entry[0]}\t{entry[1]}\t{entry[2]}\t{entry[3]}")
            print("="*60)
