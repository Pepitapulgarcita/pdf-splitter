import os
import pandas as pd
from pypdf import PdfReader, PdfWriter
import tkinter as tk
from tkinter import filedialog, messagebox
import re

def clean_filename_part(val):
    if pd.isna(val): return "MissingData"
    s = str(val).replace('/', '')
    s = s.replace('\xa0', ' ')
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def normalize_account(account):
    """Remove all spaces and IBAN prefix from account number for comparison"""
    if pd.isna(account): return ""
    # Remove spaces, non-breaking spaces, and IBAN prefix
    s = str(account).replace(' ', '').replace('\xa0', '').strip()
    # Remove 'IBAN' prefix if present (case insensitive)
    s = re.sub(r'^IBAN', '', s, flags=re.IGNORECASE)
    return s

def normalize_amount(amount):
    """Normalize amount for comparison (remove EUR, spaces, convert to float)"""
    if pd.isna(amount): return None
    # Convert to string, remove EUR, spaces, replace comma with dot
    s = str(amount).replace(' ', '').replace('\xa0', '').replace('EUR', '').replace('€', '').strip()
    # Replace comma with dot for decimal
    s = s.replace(',', '.')
    try:
        return float(s)
    except:
        return None

def extract_pdf_data(page):
    """Extract Importe a liquidar and Cuenta from PDF page"""
    text = page.extract_text()
    
    # Search for "Importe a liquidar" or "Importe a Liquidar"
    importe = None
    importe_match = re.search(r'Importe\s+a\s+liquidar[:\s]*([0-9.,\s]+)', text, re.IGNORECASE)
    if importe_match:
        importe = normalize_amount(importe_match.group(1))
    
    # Search for "Cuenta" (account number, typically format ES__ ____ ____ ____ ____)
    cuenta = None
    # Look for Spanish IBAN format or account number
    cuenta_match = re.search(r'Cuenta[:\s]*(ES[0-9\s]{22,}|[0-9\s]{20,})', text, re.IGNORECASE)
    if cuenta_match:
        cuenta = normalize_account(cuenta_match.group(1))
    
    return importe, cuenta, text

def split_and_rename_pdf():
    root = tk.Tk()
    root.withdraw()
    
    excel_path = filedialog.askopenfilename(title="1. Select the RT Excel file", 
                                            filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv")])
    if not excel_path: return
        
    pdf_path = filedialog.askopenfilename(title="2. Select the RM PDF file", 
                                          filetypes=[("PDF files", "*.pdf")])
    if not pdf_path: return
    
    # Use Downloads folder
    downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
    output_folder = os.path.join(downloads_path, "Splitted_PDFs")
    
    try:
        # Load Data
        if excel_path.endswith('.csv'):
            df = pd.read_csv(excel_path)
        else:
            df = pd.read_excel(excel_path)
       
        df.columns = [str(col).strip().replace('\xa0', ' ') for col in df.columns]    
        
        reader = PdfReader(pdf_path)
        num_pages = len(reader.pages)
        num_rows = len(df)
        
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        
        # Check required columns
        required_cols = ["Año / Nº de justificante", "Asociado a Año / Nº", "Importe a pagar"]
        
        # Try to find account column (might have slight variations)
        account_col = None
        for col in df.columns:
            if 'cuenta' in col.lower() and 'tercer' in col.lower():
                account_col = col
                break
        
        for col in required_cols:
            if col not in df.columns:
                messagebox.showerror("Column Error", 
                    f"Required column '{col}' not found.\n\n"
                    f"Found columns: {list(df.columns)}")
                return
        
        if not account_col:
            messagebox.showwarning("Warning", 
                "Could not find 'Cuenta del tercer/cesionario' column.\n"
                "Validation will be skipped.\n\n"
                f"Available columns: {list(df.columns)}")
        
        count = 0
        mismatches = []
        limit = min(num_pages, num_rows)
        
        for i in range(limit):
            row = df.iloc[i]
            page = reader.pages[i]
            
            # Extract values for filename
            p1 = clean_filename_part(row["Año / Nº de justificante"])
            p2 = clean_filename_part(row["Asociado a Año / Nº"])
            p3 = clean_filename_part(row["Importe a pagar"])
            
            # Validate data if account column exists
            if account_col:
                excel_amount = normalize_amount(row["Importe a pagar"])
                excel_account = normalize_account(row[account_col])
                
                pdf_amount, pdf_account, pdf_text = extract_pdf_data(page)
                
                # Check for mismatches
                amount_match = (pdf_amount is not None and excel_amount is not None and 
                               abs(pdf_amount - excel_amount) < 0.01)
                account_match = (pdf_account and excel_account and pdf_account == excel_account)
                
                if not amount_match or not account_match:
                    mismatches.append({
                        'page': i + 1,
                        'row': i + 1,
                        'excel_amount': excel_amount,
                        'pdf_amount': pdf_amount,
                        'excel_account': excel_account,
                        'pdf_account': pdf_account,
                        'filename': f"{p1}_{p2}_{p3}.pdf"
                    })
            
            # Construct filename
            new_filename = f"{p1}_{p2}_{p3}.pdf"
            new_filename = re.sub(r'[<>:"/\\|?*]', '', new_filename)
            
            # Split and Save
            writer = PdfWriter()
            writer.add_page(page)
            
            output_path = os.path.join(output_folder, new_filename)
            with open(output_path, "wb") as f:
                writer.write(f)
            count += 1
        
        # Show results
        if mismatches:
            # Create detailed mismatch report file
            report_path = os.path.join(output_folder, "MISMATCH_REPORT.txt")
            with open(report_path, "w", encoding="utf-8") as report:
                report.write("=" * 80 + "\n")
                report.write("DATA MISMATCH REPORT\n")
                report.write("=" * 80 + "\n\n")
                report.write(f"Total files created: {count}\n")
                report.write(f"Total mismatches found: {len(mismatches)}\n")
                report.write(f"Validation success rate: {((count - len(mismatches)) / count * 100):.1f}%\n\n")
                report.write("=" * 80 + "\n\n")
                
                for m in mismatches:
                    report.write(f"MISMATCH #{mismatches.index(m) + 1}\n")
                    report.write(f"{'-' * 80}\n")
                    report.write(f"Row/Page Number: {m['row']}\n")
                    report.write(f"Filename: {m['filename']}\n\n")
                    
                    report.write(f"AMOUNT COMPARISON:\n")
                    report.write(f"  Excel (Importe a pagar): {m['excel_amount']}\n")
                    report.write(f"  PDF (Importe a liquidar): {m['pdf_amount']}\n")
                    amount_match = "✓ MATCH" if (m['pdf_amount'] and m['excel_amount'] and 
                                                 abs(m['pdf_amount'] - m['excel_amount']) < 0.01) else "✗ MISMATCH"
                    report.write(f"  Status: {amount_match}\n\n")
                    
                    report.write(f"ACCOUNT COMPARISON:\n")
                    report.write(f"  Excel (Cuenta del tercer/cesionario): {m['excel_account']}\n")
                    report.write(f"  PDF (Cuenta): {m['pdf_account']}\n")
                    account_match = "✓ MATCH" if (m['pdf_account'] and m['excel_account'] and 
                                                  m['pdf_account'] == m['excel_account']) else "✗ MISMATCH"
                    report.write(f"  Status: {account_match}\n\n")
                    report.write("=" * 80 + "\n\n")
            
            # Show summary in messagebox
            mismatch_msg = "⚠️ DATA MISMATCHES FOUND!\n\n"
            for m in mismatches[:5]:  # Show first 5
                mismatch_msg += f"Row/Page {m['row']}:\n"
                mismatch_msg += f"  Amount: Excel={m['excel_amount']}, PDF={m['pdf_amount']}\n"
                mismatch_msg += f"  Account: Excel={m['excel_account']}, PDF={m['pdf_account']}\n\n"
            
            if len(mismatches) > 5:
                mismatch_msg += f"... and {len(mismatches) - 5} more mismatches\n\n"
            
            mismatch_msg += f"Total mismatches: {len(mismatches)} out of {count}\n"
            mismatch_msg += f"Files created in: {output_folder}\n\n"
            mismatch_msg += f"📄 Detailed report saved to:\nMISMATCH_REPORT.txt\n\n"
            mismatch_msg += "⚠️ Please verify the data alignment!"
            
            messagebox.showwarning("Validation Warning", mismatch_msg)
        else:
            messagebox.showinfo("Success", 
                f"✓ Process complete!\n\n"
                f"Files created: {count}\n"
                f"All data validated successfully!\n"
                f"Folder: {output_folder}")
    
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

if __name__ == "__main__":
    split_and_rename_pdf()
    split_and_rename_pdf()
