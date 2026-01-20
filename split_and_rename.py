import os
import pandas as pd
from pypdf import PdfReader, PdfWriter
import tkinter as tk
from tkinter import filedialog, messagebox
import re

def clean_filename_part(val):
    if pd.isna(val): return "MissingData"
    s = str(val).replace('/', '')
    # Removes non-breaking spaces (\xa0) and multiple spaces
    s = s.replace('\xa0', ' ')
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def split_and_rename_pdf():
    # Hide the main tkinter window
    root = tk.Tk()
    root.withdraw()

    # 1. Ask for files
    excel_path = filedialog.askopenfilename(title="1. Select the RT Excel file", 
                                            filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv")])
    if not excel_path: return
        
    pdf_path = filedialog.askopenfilename(title="2. Select the RM PDF file", 
                                          filetypes=[("PDF files", "*.pdf")])
    if not pdf_path: return

    output_folder = "Splitted_PDFs"

    try:
        # 2. Load Data
        if excel_path.endswith('.csv'):
            df = pd.read_csv(excel_path)
        else:
            df = pd.read_excel(excel_path)
       
        # --- CRITICAL FIX: Clean the column names ---
        # This removes the invisible spaces from "  Año / Nº..."
        df.columns = [str(col).strip().replace('\xa0', ' ') for col in df.columns]    
        
        reader = PdfReader(pdf_path)
        num_pages = len(reader.pages)
        num_rows = len(df)
        
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        # 3. Process Updated check using cleaned column names
        required_cols = ["Año / Nº de justificante", "Asociado a Año / Nº", "Importe a pagar"]
        for col in required_cols:
            if col not in df.columns:
                messagebox.showerror("Column Error", 
                    f"Required column '{col}' not found.\n\n"
                    f"Found columns: {list(df.columns)}")
                return

        count = 0
        limit = min(num_pages, num_rows)
        
        for i in range(limit):
            row = df.iloc[i]
            
            # Extract values based on your column names
            p1 = clean_filename_part(row["Año / Nº de justificante"])
            p2 = clean_filename_part(row["Asociado a Año / Nº"])
            p3 = clean_filename_part(row["Importe a pagar"])
            
            # Construct filename
            new_filename = f"{p1}_{p2}_{p3}.pdf"
            # Remove any remaining invalid characters for Windows filenames
            new_filename = re.sub(r'[<>:"/\\|?*]', '', new_filename)

            # Split and Save
            writer = PdfWriter()
            writer.add_page(reader.pages[i])
            
            output_path = os.path.join(output_folder, new_filename)
            with open(output_path, "wb") as f:
                writer.write(f)
            count += 1

        messagebox.showinfo("Success", f"Process complete!\n\nFiles created: {count}\nFolder: {output_folder}")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

if __name__ == "__main__":
    split_and_rename_pdf()
