import os
import re
import pdfplumber
import pandas as pd

pdf_dir = r"C:\Users\Altersense\Desktop\ERP-RPA\Sample\Main"

for filename in os.listdir(pdf_dir):
    if filename.lower().endswith(".pdf"):
        pdf_path = os.path.join(pdf_dir, filename)
        with pdfplumber.open(pdf_path) as pdf:
            all_text = ""
            for page in pdf.pages:
                all_text += page.extract_text() + "\n" or ""
        
                print(f"-----All text:\n{all_text}\n-----")
