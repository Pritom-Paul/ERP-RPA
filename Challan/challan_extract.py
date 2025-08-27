import os
import camelot
import pandas as pd

pdf_dir = r"C:\Users\Altersense\Desktop\ERP-RPA\Challan"

for filename in os.listdir(pdf_dir):
    if filename.lower().endswith(".pdf"):
        pdf_path = os.path.join(pdf_dir, filename)
        print(f"\n{'='*80}")
        print(f"FILE: {filename}")
        print(f"{'='*80}")
        
        try:
            # Extract tables from PDF using Camelot
            tables = camelot.read_pdf(pdf_path, pages='all', flavor='lattice')
            
            print(f"Number of tables found: {tables.n}")
            
            if tables.n == 0:
                print("No tables found in this PDF.")
                continue
            
            for i, table in enumerate(tables):
                print(f"\n--- Table {i+1} ---")
                print(f"Accuracy: {table.accuracy:.2f}")
                print(f"Whitespace: {table.whitespace:.2f}")
                print(f"Order: {table.order}")
                print(f"Page: {table.page}")
                
                # Get the DataFrame
                df = table.df
                
                print("\nDataFrame:")
                print(df)
                print(f"\nShape: {df.shape}")
                print("-" * 40)
                
                # Optional: Save each table to CSV
                # csv_filename = f"{os.path.splitext(filename)[0]}_table_{i+1}.csv"
                # df.to_csv(os.path.join(pdf_dir, csv_filename), index=False)
                        
        except Exception as e:
            print(f"Error processing {filename}: {str(e)}")