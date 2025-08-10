# RUN THE SCRIPT WITH THE LINE BELOW
# #  python -m uvicorn main:app --reload
##-----------------------------------------------------
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import JSONResponse
import pdfplumber
import pandas as pd
import re
import tempfile
from fastapi import HTTPException

app = FastAPI()


@app.post("/extract")
async def extract_pdf_data(file: UploadFile = File(...)):
    # Save uploaded file to a temp location
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        contents = await file.read()
        tmp.write(contents)
        tmp_path = tmp.name

    with pdfplumber.open(tmp_path) as pdf:
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
            raise HTTPException(status_code=400, detail="Invalid file format: Header not found.")
        
        # Process all pages using the found last header line
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            all_text += page_text + "\n"
            page_lines = page_text.splitlines()
            # Check if header line exists in this page
            found_header_in_page = False
            for idx, line in enumerate(page_lines):
                if last_header_line in line.strip().lower():
                    found_header_in_page = True
                    if idx + 1 < len(page_lines):
                        body_text_lines.extend(page_lines[idx + 1:])
                    break

            if not found_header_in_page:
                raise HTTPException(status_code=400, detail="Invalid file format: Header missing on one or more pages.")

    # All field extraction logic remains the same
    rows = extract_fields(all_text, body_text_lines)

    return JSONResponse(content=rows)


def extract_fields(all_text, body_text_lines):
    lines = all_text.splitlines()

    po_match = re.search(r"Purchase Order\s+(PO\d+)", all_text, re.IGNORECASE)
    po_no = po_match.group(1) if po_match else None

    buying_house = next((line.strip() for line in lines if line.strip()), None)

    buying_house_add_lines = []
    for line in lines[1:]:
        if "purchase order" in line.lower():
            break
        buying_house_add_lines.append(line)
    buying_house_add = "\n".join(buying_house_add_lines) if buying_house_add_lines else None

    ship_to_address = None
    for idx, line in enumerate(lines):
        if "ship-to address" in line.lower():
            for next_line in lines[idx + 1:]:
                if next_line.strip():
                    ship_to_address = next_line.strip()
                    break
            break

    vendor_no_match = re.search(r"Vendor No.\s+(\d+)", all_text, re.IGNORECASE)
    vendor_no = vendor_no_match.group(1) if vendor_no_match else None

    payment_terms_match = re.search(r"Payment Terms\s+(.+?)(?=\s+Prices)", all_text, re.IGNORECASE)
    payment_terms = payment_terms_match.group(1) if payment_terms_match else None

    document_date_match = re.search(r"Document Date\s+(\d{2}[-/]\d{2}[-/]\d{2,4})", all_text, re.IGNORECASE)
    document_date = document_date_match.group(1) if document_date_match else None

    shipment_method_match = re.search(r"Shipment Method\s+(.+?)(?=\s+Order Type)", all_text, re.IGNORECASE)
    shipment_method = shipment_method_match.group(1) if shipment_method_match else ""

    order_type_match = re.search(r"Order Type\s+(.*)", all_text, re.IGNORECASE)
    order_type = order_type_match.group(1).strip() if order_type_match else None

    shipping_agent_match = re.search(r"Shipping Agent\s+(.*)", all_text, re.IGNORECASE)
    shipping_agent = shipping_agent_match.group(1).strip() if shipping_agent_match else None

    order_for_match = re.search(r"Order For\s+(.*)", all_text, re.IGNORECASE)
    order_for = order_for_match.group(1).strip() if order_for_match else None

    style_no = None
    description = None
    for i, line in enumerate(lines):
        if re.search(r"Style No\.", line, re.IGNORECASE):
            if i + 1 < len(lines):
                next_line = lines[i + 1].strip()
                words = next_line.split()
                if words:
                    style_no = words[0]
                    if len(words) > 1:
                        description = " ".join(words[1:])
            break

    hs_code_match = re.search(r"HS CODE:\s+(\d+)", all_text, re.IGNORECASE)
    hs_code = hs_code_match.group(1) if hs_code_match else None

    total_qty_match = re.search(r"Total Qty.\s+(\d+)", all_text, re.IGNORECASE)
    total_qty = total_qty_match.group(1) if total_qty_match else None

    prices_including_vat_match = re.search(r"Prices Including VAT\s+(\w+)", all_text, re.IGNORECASE)
    prices_including_vat = prices_including_vat_match.group(1) if prices_including_vat_match else None

    group_entries = []
    i = 0
    while i < len(body_text_lines):
        line = body_text_lines[i].strip()
        if "HS Code:" in line:
            i += 1
            continue

        if re.match(r"^\d{6}\b", line):
            i += 1
            if i >= len(body_text_lines): break
            size_line = body_text_lines[i].strip()
            sizes = size_line.split()
            i += 1
            if i >= len(body_text_lines): break
            qty_line = body_text_lines[i].strip()

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

    barcodes = []
    barcode_started = False
    for line in body_text_lines:
        stripped_line = line.strip()
        if not barcode_started:
            if "barcode" in stripped_line.lower():
                barcode_started = True
            continue
        if not stripped_line:
            break
        match = re.search(r"(\d{13})\s*$", stripped_line)
        if match:
            barcodes.append(match.group(1))
        else:
            break
    
    if len(barcodes) != len(group_entries):
        error_msg = f"Warning: Number of barcodes ({len(barcodes)}) does not match group entries ({len(group_entries)})"
        # You could either:
        # 1) Return an error response immediately
        # return JSONResponse(
        #     content={"error": error_msg, "detail": body_text_lines},
        #     status_code=400
        # )
        
        # OR 2) Include the warning in the response (current implementation)
        print(error_msg)  # This will log to your server console
        print(f"Body text lines: {body_text_lines}")

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
            "Prepacks": "",
            "Total Quantity": total_qty,
            "Batch Quantity": batch_qty,
            "Prices Including VAT": prices_including_vat
        })

    return rows