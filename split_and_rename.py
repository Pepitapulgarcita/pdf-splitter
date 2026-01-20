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
    s = str(account).replace(' ', '').replace('\xa0', '').strip()
    s = re.sub(r'^IBAN', '', s, flags=re.IGNORECASE)
    return s

def normalize_amount(amount):
    """Normalize amount for comparison (remove EUR, spaces, convert to float)"""
    if pd.isna(amount): return None
    s = str(amount).replace(' ', '').replace('\xa0', '').replace('EUR', '').replace('€', '').strip()
    s = s.replace(',', '.')
    try:
        return float(s)
    except:
        return None

def extract_pdf_data(page):
    """Extract Importe a liquidar and Cuenta from PDF page"""
    text = page.extract_text()
    
    # Search for "Importe a liquidar"
    importe = None
    importe_match = re.search(r'Importe\s+a\s+liquidar[:\s]*([0-9.,\s]+(?:EUR)?)', text, re.IGNORECASE)
    if importe_match:
        importe = normalize_amount(importe_match.group(1))
    
    # Search for "Cuenta"
    cuenta = None
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
        
        # Find account column
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
            response = messagebox.askyesno("Warning", 
                "Could not find 'Cuenta del tercer/cesionario' column.\n"
                "Files will be named based on row order without validation.\n\n"
                f"Available columns: {list(df.columns)}\n\n"
                "Continue anyway?")
            if not response:
                return
        
        # Step 1: Extract data from all PDF pages
        pdf_data = []
        for i in range(num_pages):
            page = reader.pages[i]
            importe, cuenta, text = extract_pdf_data(page)
            pdf_data.append({
                'page_num': i,
                'page': page,
                'importe': importe,
                'cuenta': cuenta
            })
        
        # Step 2: Process based on whether we can validate
        count = 0
        mismatches = []
        unmatched_pages = []
        
        if account_col:
            # VALIDATION MODE: Match pages to rows based on data
            matched_pages = set()
            
            for row_idx in range(num_rows):
                row = df.iloc[row_idx]
                excel_amount = normalize_amount(row["Importe a pagar"])
                excel_account = normalize_account(row[account_col])
                
                # Find matching PDF page
                matching_page = None
                for pdf_item in pdf_data:
                    if pdf_item['page_num'] in matched_pages:
                        continue
                    
                    amount_match = (pdf_item['importe'] and excel_amount and 
                                   abs(pdf_item['importe'] - excel_amount) < 0.01)
                    account_match = (pdf_item['cuenta'] and excel_account and 
                                    pdf_item['cuenta'] == excel_account)
                    
                    if amount_match and account_match:
                        matching_page = pdf_item
                        matched_pages.add(pdf_item['page_num'])
                        break
                
                if matching_page:
                    # Create filename from Excel data
                    p1 = clean_filename_part(row["Año / Nº de justificante"])
                    p2 = clean_filename_part(row["Asociado a Año / Nº"])
                    p3 = clean_filename_part(row["Importe a pagar"])
                    
                    new_filename = f"{p1}_{p2}_{p3}.pdf"
                    new_filename = re.sub(r'[<>:"/\\|?*]', '', new_filename)
                    
                    # Save PDF
                    writer = PdfWriter()
                    writer.add_page(matching_page['page'])
                    output_path = os.path.join(output_folder, new_filename)
                    
                    with open(output_path, "wb") as f:
                        writer.write(f)
                    count += 1
                else:
                    # No match found - save with row order for manual review
                    mismatches.append({
                        'row': row_idx + 1,
                        'excel_amount': excel_amount,
                        'excel_account': excel_account,
                        'status': 'No matching PDF page found'
                    })
            
            # Check for unmatched PDF pages
            for pdf_item in pdf_data:
                if pdf_item['page_num'] not in matched_pages:
                    unmatched_pages.append({
                        'page': pdf_item['page_num'] + 1,
                        'pdf_amount': pdf_item['importe'],
                        'pdf_account': pdf_item['cuenta']
                    })
        else:
            # SIMPLE MODE: Assume page order matches row order
            limit = min(num_pages, num_rows)
            
            for i in range(limit):
                row = df.iloc[i]
                page = reader.pages[i]
                
                p1 = clean_filename_part(row["Año / Nº de justificante"])
                p2 = clean_filename_part(row["Asociado a Año / Nº"])
                p3 = clean_filename_part(row["Importe a pagar"])
                
                new_filename = f"{p1}_{p2}_{p3}.pdf"
                new_filename = re.sub(r'[<>:"/\\|?*]', '', new_filename)
                
                writer = PdfWriter()
                writer.add_page(page)
                output_path = os.path.join(output_folder, new_filename)
                
                with open(output_path, "wb") as f:
                    writer.write(f)
                count += 1
        
        # Generate report
        if mismatches or unmatched_pages:
            report_path = os.path.join(output_folder, "MISMATCH_REPORT.txt")
            with open(report_path, "w", encoding="utf-8") as report:
                report.write("=" * 80 + "\n")
                report.write("DATA MATCHING REPORT\n")
                report.write("=" * 80 + "\n\n")
                report.write(f"Total files created: {count}\n")
                report.write(f"Excel rows without matching PDF: {len(mismatches)}\n")
                report.write(f"PDF pages without matching Excel row: {len(unmatched_pages)}\n\n")
                report.write("=" * 80 + "\n\n")
                
                if mismatches:
                    report.write("EXCEL ROWS WITHOUT MATCHING PDF PAGE:\n")
                    report.write("-" * 80 + "\n\n")
                    for m in mismatches:
                        report.write(f"Row Number: {m['row']}\n")
                        report.write(f"  Excel Amount: {m['excel_amount']}\n")
                        report.write(f"  Excel Account: {m['excel_account']}\n")
                        report.write(f"  Status: {m['status']}\n\n")
                    report.write("=" * 80 + "\n\n")
                
                if unmatched_pages:
                    report.write("PDF PAGES WITHOUT MATCHING EXCEL ROW:\n")
                    report.write("-" * 80 + "\n\n")
                    for u in unmatched_pages:
                        report.write(f"Page Number: {u['page']}\n")
                        report.write(f"  PDF Amount: {u['pdf_amount']}\n")
                        report.write(f"  PDF Account: {u['pdf_account']}\n\n")
            
            msg = f"⚠️ MATCHING ISSUES FOUND!\n\n"
            msg += f"Files created: {count}\n"
            msg += f"Excel rows unmatched: {len(mismatches)}\n"
            msg += f"PDF pages unmatched: {len(unmatched_pages)}\n\n"
            msg += f"📄 See MISMATCH_REPORT.txt for details\n"
            msg += f"Folder: {output_folder}"
            
            messagebox.showwarning("Validation Warning", msg)
        else:
            messagebox.showinfo("Success", 
                f"✓ Process complete!\n\n"
                f"Files created: {count}\n"
                f"All pages matched successfully!\n"
                f"Folder: {output_folder}")
    
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

if __name__ == "__main__":
    split_and_rename_pdf()
