import pandas as pd
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, simpledialog
import re

def extract_invoice_details_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text("text")
    doc.close()
    invoice_details = re.findall(r'\b(\d{8})\s+(\d{2}/\d{2}/\d{2})\s+(\d+,\d{2}|\d+\.\d{2})\b', text)
    cleaned_details = []
    for num, date, amount in invoice_details:
        clean_amount = "{:.2f}".format(float(amount.replace(',', '')))
        cleaned_detail = f"{num} {date} {clean_amount}"
        cleaned_details.append(cleaned_detail)
    return set(cleaned_details)

def read_invoice_details_from_csv(csv_path):
    df = pd.read_csv(csv_path, usecols=['Date', 'No.', 'Total'])
    df['Date'] = pd.to_datetime(df['Date']).dt.strftime('%m/%d/%y')
    df['Total'] = df['Total'].replace(',', '', regex=True).astype(float).map("{:.2f}".format)
    df['combined'] = df['No.'].astype(str) + ' ' + df['Date'] + ' ' + df['Total']
    return set(df['combined'])

def find_missing_invoices(csv_invoices, pdf_invoices):
    return pdf_invoices - csv_invoices

def save_invoices_to_excel(missing_invoices, file_path):
    df = pd.DataFrame(list(missing_invoices), columns=['Missing Invoice Details'])
    df.to_excel(file_path, index=False, engine='openpyxl')

def main():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    pdf_path = filedialog.askopenfilename(title="Select PDF file", filetypes=[("PDF files", "*.pdf")])
    csv_path = filedialog.askopenfilename(title="Select CSV file", filetypes=[("CSV files", "*.csv")])
    output_name = simpledialog.askstring("Output File Name", "Enter the name for the output Excel file:")
    
    if pdf_path and csv_path and output_name:
        pdf_invoices = extract_invoice_details_from_pdf(pdf_path)
        csv_invoices = read_invoice_details_from_csv(csv_path)
        missing_invoices = find_missing_invoices(csv_invoices, pdf_invoices)
        
        if not output_name.endswith('.xlsx'):
            output_name += '.xlsx'
        
        save_invoices_to_excel(missing_invoices, output_name)
        print("File has been saved as:", output_name)
    else:
        print("Operation canceled.")
    
    root.destroy()

if __name__ == "__main__":
    main()
