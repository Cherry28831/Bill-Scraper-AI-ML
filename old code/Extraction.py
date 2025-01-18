import pdfplumber
import pytesseract
from PIL import Image
import re
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilenames

# Function to extract text from images using OCR (if needed)
def extract_text_from_image(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            img = page.to_image()
            text += pytesseract.image_to_string(img.original)
    return text

# Function to extract all text from PDF
def extract_full_text(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text()  # Extracts general text from each page
    return text

def extract_text_and_debug(pdf_path):
    text = extract_full_text(pdf_path)

    if not text.strip():
        print("No text found using standard extraction. Switching to OCR.")
        text = extract_text_from_image(pdf_path)

    print("Extracted text for debugging:")  # Debugging first 1000 chars
    print(text[:1000])  # Print the first 1000 characters for inspection
    return text

# Function to extract relevant details using regex
def extract_details(text):
    details = {}

    try:
        gstin_pattern = r"\b[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[A-Z0-9]{1}[Z]{1}[A-Z0-9]{1}\b"
        gstin_matches = re.findall(gstin_pattern, text)
        details['GSTIN'] = gstin_matches[-1] if gstin_matches else ""

        details['GSTIN NO'] = re.search(r'GSTIN\s*NO\s*[:\-]?\s*([\w\d]+)', text).group(1) if re.search(r'GSTIN\s*NO\s*[:\-]?\s*([\w\d]+)', text) else ""
        
        company_name_match = re.search(r'^[A-Z][A-Z &,-\.]+(?:\s+Ltd.*)?(?=\s+At\.)', text, re.MULTILINE)
        details['Company Name'] = company_name_match.group(0).strip() if company_name_match else ""
        details['Invoice No'] = re.search(r'Invoice No\.\s*[:\-]?\s*(\d+)', text).group(1) if re.search(r'Invoice No\.\s*[:\-]?\s*(\d+)', text) else ""
        details['Date of Invoice'] = re.search(r'Date of Invoice\s*[:\-]?\s*([\d\-]+)', text).group(1) if re.search(r'Date of Invoice\s*[:\-]?\s*([\d\-]+)', text) else ""

        shipped_to_match = re.search(r'Shipped to\s*[:\-]?\s*([\s\S]+?)\s*FSSAI', text)
        details['Shipped to'] = shipped_to_match.group(1).strip() if shipped_to_match else ""

        goods_description_match = re.findall(r'\d+\.\s+(.*?)\s+(\d{8})\s+([\d\.]+)\s+(\d+\.\d+)\s+(\d+\.\d+)\s+([\d,]+)', text)

        if goods_description_match:
            details['Goods Description'] = []
            details['HSN/SAC'] = []
            details['Bags'] = []
            details['Pack'] = []
            details['Quintal'] = []
            details['Rate'] = []
            details['Amount'] = []
            
            for match in goods_description_match:
                if len(match) == 6:  
                    goods, hsn, bags, pack, quintal, rate = match
                    quintal = quintal.replace(",", "")
                    rate = rate.replace(",", "")
                    amount = round(float(quintal) * float(rate), 2)
                    details['Goods Description'].append(goods)
                    details['HSN/SAC'].append(hsn)
                    details['Bags'].append(bags)
                    details['Pack'].append(pack)
                    details['Quintal'].append(quintal)
                    details['Rate'].append(rate)
                    details['Amount'].append(amount)
                else:
                    details['Goods Description'].append(match[0] if len(match) > 0 else "")
                    details['HSN/SAC'].append(match[1] if len(match) > 1 else "")
                    details['Bags'].append(match[2] if len(match) > 2 else "")
                    details['Pack'].append(match[3] if len(match) > 3 else "")
                    details['Quintal'].append(match[4] if len(match) > 4 else "")
                    details['Rate'].append(match[5] if len(match) > 5 else "")
                    details['Amount'].append(0.0)

        else:
            details['Goods Description'], details['HSN/SAC'], details['Bags'], details['Pack'], details['Quintal'], details['Rate'], details['Amount'] = [], [], [], [], [], [], []

        details['Transport'] = re.search(r'Transport\s*[:\-]?\s*([\s\S]+?)\s+Despatch Date', text).group(1).strip() if re.search(r'Transport\s*[:\-]?\s*([\s\S]+?)\s+Despatch Date', text) else ""
        details['Vehicle No'] = re.search(r'Vehicle No\.\s*[:\-]?\s*([\w\d]+)', text).group(1) if re.search(r'Vehicle No\.\s*[:\-]?\s*([\w\d]+)', text) else ""
        details['Licence No'] = re.search(r'Licence No\s*[:\-]?\s*([\w\d]+)', text).group(1) if re.search(r'Licence No\s*[:\-]?\s*([\w\d]+)', text) else ""
        details['Mobile No'] = re.search(r'Mobile No\s*[:\-]?\s*([\d]+)', text).group(1) if re.search(r'Mobile No\s*[:\-]?\s*([\d]+)', text) else ""

    except Exception as e:
        print(f"Error extracting data: {e}")

    return details

# Function to write extracted details to an Excel file
def write_to_excel(all_details, output_path):
    all_data_frames = []

    for details in all_details:
        goods_description = details.get('Goods Description', [])
        hsn_sac = details.get('HSN/SAC', [])
        bags = details.get('Bags', [])
        pack = details.get('Pack', [])
        quintal = details.get('Quintal', [])
        rate = details.get('Rate', [])
        amount = details.get('Amount', [])

        goods_df = pd.DataFrame({
            'Goods Description': goods_description,
            'HSN/SAC': hsn_sac,
            'Bags': bags,
            'Pack': pack,
            'Quintal': quintal,
            'Rate': rate,
            'Amount': amount,
        })

        for column in ['Company Name', 'Invoice No', 'Date of Invoice', 'GSTIN NO', 'GSTIN',
                       'Shipped to', 'Transport', 'Vehicle No', 'Licence No', 'Mobile No']:
            value = details.get(column, "")
            goods_df[column] = value  # Broadcast single value to all rows

        all_data_frames.append(goods_df)

    # Concatenate all DataFrames into one
    final_df = pd.concat(all_data_frames, ignore_index=True)

    # Save to Excel
    final_df.to_excel(output_path, index=False)
    print(f"Data written to {output_path}")

# Main execution
def main():
    print("Please select the PDF files:")
    Tk().withdraw()  # Hides the root window
    pdf_paths = askopenfilenames(filetypes=[("PDF files", "*.pdf")])  # Allow multiple file selection

    all_details = []

    for pdf_path in pdf_paths:
        print(f"Processing {pdf_path}...")
        
        text = extract_full_text(pdf_path)

        if not text.strip():
            print("No text found in the PDF. Using OCR for text extraction.")
            text = extract_text_from_image(pdf_path)

        details = extract_details(text)
        
        all_details.append(details)

    excel_path = "consolidated_invoice_details.xlsx"
    write_to_excel(all_details, excel_path)
    print(f"Extracted and consolidated data saved to {excel_path}")

if __name__ == "__main__":
    main()
