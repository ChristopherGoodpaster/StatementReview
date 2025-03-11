import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

# Define output directory
OUTPUT_DIR = r"C:\Users\Chris\Desktop\EatWell\Project\StatementReview"

def extract_pdf_to_csv(pdf_path, output_csv_path):
    """Extracts invoice data from PDF and saves as CSV."""
    data = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            lines = text.split("\n")  # Split into lines
            
            for i in range(len(lines)):
                if "INVOICE NO." in lines[i]:  # Look for Invoice Number
                    try:
                        invoice_no = lines[i + 1].strip()  # Next line is Invoice No
                        amount = lines[i + 3].strip()  # 3 lines below is Amount
                        date = lines[i - 1].strip()  # Previous line is Date
                        
                        # Clean up values
                        invoice_no = invoice_no.replace(",", "").strip()
                        amount = amount.replace(",", "").strip()
                        date = date.replace(" ", "/").strip()  # Convert "2 10 2025" -> "2/10/2025"
                        
                        data.append([invoice_no, date, amount])
                    
                    except IndexError:
                        continue  # Skip if something is missing

    # Save extracted data to CSV
    df = pd.DataFrame(data, columns=["Invoice No", "Date", "Amount"])
    df.to_csv(output_csv_path, index=False)
    
    print(f"✅ PDF converted to CSV: {output_csv_path}")

def compare_csv_files(pdf_csv_path, quickbooks_csv_path, output_csv_path):
    """Compares extracted PDF data with QuickBooks data to find missing invoices."""
    df_pdf = pd.read_csv(pdf_csv_path)
    df_qb = pd.read_csv(quickbooks_csv_path)

    # Convert columns to string (avoid float issues)
    df_pdf["Invoice No"] = df_pdf["Invoice No"].astype(str)
    df_qb["No."] = df_qb["No."].astype(str)

    # Merge DataFrames to find missing invoices
    merged_df = df_pdf.merge(df_qb, left_on=["Invoice No"], right_on=["No."], how="left", suffixes=("_PDF", "_CSV"))

    # Find missing invoices (if "Total_CSV" is NaN, it's missing from QuickBooks)
    missing_invoices = merged_df[merged_df["Total"].isna()]

    # Save missing invoices to CSV
    missing_invoices.to_csv(output_csv_path, index=False)

    print(f"✅ Missing invoices saved to: {output_csv_path}")

def select_files():
    """GUI to select files and run the comparison."""
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    pdf_path = filedialog.askopenfilename(title="Select PDF Statement", filetypes=[("PDF Files", "*.pdf")])
    if not pdf_path:
        return
    
    quickbooks_csv_path = filedialog.askopenfilename(title="Select QuickBooks CSV", filetypes=[("CSV Files", "*.csv")])
    if not quickbooks_csv_path:
        return
    
    output_csv_filename = filedialog.asksaveasfilename(title="Save Missing Invoices CSV", defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
    if not output_csv_filename:
        return
    
    pdf_csv_path = os.path.join(OUTPUT_DIR, "extracted_pdf_data.csv")

    # Step 1: Convert PDF to CSV
    extract_pdf_to_csv(pdf_path, pdf_csv_path)

    # Step 2: Compare CSVs
    compare_csv_files(pdf_csv_path, quickbooks_csv_path, output_csv_filename)

    messagebox.showinfo("Success", f"Comparison complete! Missing invoices saved to:\n{output_csv_filename}")

if __name__ == "__main__":
    select_files()
