import os
import pandas as pd
from pypdf import PdfReader, PdfWriter
import tkinter as tk
from tkinter import filedialog, messagebox
import re

def clean_filename_part(val, allow_blank_underscore=False):
    if pd.isna(val) or str(val).strip() == '':
        return "_" if allow_blank_underscore else "MissingData"
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
    """Normalize amount for comparison - handles European format (1.000,50)"""
    if pd.isna(amount): return None
    
    # Convert to string and clean
    s = str(amount).replace(' ', '').replace('\xa0', '').replace('EUR', '').replace('€', '').strip()
    
    # Handle European number format: 1.000,50 or 1000,50
    # Strategy: Remove thousand separators (.), then replace decimal comma with dot
    # First check if it contains both . and , to determine format
    if ',' in s and '.' in s:
        # European format like 1.000,50 - remove dots (thousand sep), keep comma
        s = s.replace('.', '')
        s = s.replace(',', '.')
    elif ',' in s:
        # Only comma, likely European decimal: 1000,50
        s = s.replace(',', '.')
    # else: only dots or neither - likely already correct format (1000.50 or 1000)
    
    try:
        return float(s)
    except:
        return None

def extract_pdf_data(page):
    """Extract Importe a liquidar and Cuenta from PDF page"""
    text = page.extract_text()
    
    # Search for "Importe a liquidar" or "Importe a Liquidar"
    importe = None
    importe_match = re.search(r'Importe\s+a\s+liquidar[:\s]*([0-9.,\s]+(?:EUR)?)', text, re.IGNORECASE)
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
        
        # Try to find account column
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
        matches = []
        limit = min(num_pages, num_rows)
        
        # STEP 1: Split all pages with temporary page numbers
        temp_files = []
        for i in range(limit):
            page = reader.pages[i]
            
            # Create temporary filename with page number
            temp_filename = f"Page_{i+1:03d}_TEMP.pdf"
            temp_path = os.path.join(output_folder, temp_filename)
            
            # Split and Save with temporary name
            writer = PdfWriter()
            writer.add_page(page)
            
            with open(temp_path, "wb") as f:
                writer.write(f)
            
            temp_files.append((i, temp_path))
        
        # STEP 2: Validate and rename each file
        for i, temp_path in temp_files:
            row = df.iloc[i]
            page = reader.pages[i]
            
            # Extract values for filename
            p1 = clean_filename_part(row["Año / Nº de justificante"])
            p2 = clean_filename_part(row["Asociado a Año / Nº"], allow_blank_underscore=True)
            p3 = clean_filename_part(row["Importe a pagar"])
            
            # Proposed final filename (prefixed with EXTRACTO SANTANDER)
            final_filename = f"EXTRACTO_SANTANDER_{p1}_{p2}_{p3}.pdf"
            final_filename = re.sub(r'[<>:"/\\|?*]', '', final_filename)
            final_path = os.path.join(output_folder, final_filename)
            
            # Validate data if account column exists
            validation_passed = True
            if account_col:
                excel_amount = normalize_amount(row["Importe a pagar"])
                excel_account = normalize_account(row[account_col])
                
                pdf_amount, pdf_account, pdf_text = extract_pdf_data(page)
                
                # Check for mismatches
                amount_match = (pdf_amount is not None and excel_amount is not None and 
                               abs(pdf_amount - excel_amount) < 0.01)
                account_match = (pdf_account and excel_account and pdf_account == excel_account)
                
                if amount_match and account_match:
                    matches.append({
                        'page': i + 1,
                        'row': i + 1,
                        'filename': final_filename
                    })
                    # Rename to final filename
                    os.rename(temp_path, final_path)
                else:
                    validation_passed = False
                    mismatches.append({
                        'page': i + 1,
                        'row': i + 1,
                        'excel_amount': excel_amount,
                        'pdf_amount': pdf_amount,
                        'excel_account': excel_account,
                        'pdf_account': pdf_account,
                        'proposed_filename': final_filename,
                        'kept_filename': f"Page_{i+1:03d}_MISMATCH.pdf"
                    })
                    # Keep with page number but mark as mismatch
                    mismatch_path = os.path.join(output_folder, f"Page_{i+1:03d}_MISMATCH.pdf")
                    os.rename(temp_path, mismatch_path)
            else:
                # No validation, just rename
                os.rename(temp_path, final_path)
            
            count += 1
        
        # STEP 3: Create detailed report
        if mismatches:
            # Create detailed mismatch report file
            report_path = os.path.join(output_folder, "MISMATCH_REPORT.txt")
            with open(report_path, "w", encoding="utf-8") as report:
                report.write("=" * 80 + "\n")
                report.write("DATA MISMATCH REPORT\n")
                report.write("=" * 80 + "\n\n")
                report.write(f"Total files created: {count}\n")
                report.write(f"Successfully matched: {len(matches)}\n")
                report.write(f"Mismatches found: {len(mismatches)}\n")
                report.write(f"Validation success rate: {(len(matches) / count * 100):.1f}%\n\n")
                report.write("=" * 80 + "\n\n")
                
                for m in mismatches:
                    report.write(f"MISMATCH #{mismatches.index(m) + 1}\n")
                    report.write(f"{'-' * 80}\n")
                    report.write(f"Excel Row Number: {m['row']}\n")
                    report.write(f"PDF Page Number: {m['page']}\n")
                    report.write(f"File kept as: {m['kept_filename']}\n")
                    report.write(f"Would have been named: {m['proposed_filename']}\n\n")
                    
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
            mismatch_msg += f"Successfully matched: {len(matches)} files\n"
            mismatch_msg += f"Mismatches: {len(mismatches)} files\n\n"
            mismatch_msg += "Mismatched files kept as: Page_XXX_MISMATCH.pdf\n\n"
            
            for m in mismatches[:3]:  # Show first 3
                mismatch_msg += f"Page {m['page']} / Row {m['row']}:\n"
                mismatch_msg += f"  Amount: Excel={m['excel_amount']}, PDF={m['pdf_amount']}\n"
                mismatch_msg += f"  Account: Excel={m['excel_account'][:15]}..., PDF={m['pdf_account'][:15]}...\n\n"
            
            if len(mismatches) > 3:
                mismatch_msg += f"... and {len(mismatches) - 3} more mismatches\n\n"
            
            mismatch_msg += f"📄 Detailed report: MISMATCH_REPORT.txt\n"
            mismatch_msg += f"📁 Folder: {output_folder}"
            
            messagebox.showwarning("Validation Warning", mismatch_msg)
        else:
            messagebox.showinfo("Success", 
                f"✓ Process complete!\n\n"
                f"Files created: {count}\n"
                f"All data validated successfully!\n"
                f"All files properly named.\n\n"
                f"Folder: {output_folder}")
    
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

if __name__ == "__main__":
    split_and_rename_pdf()
