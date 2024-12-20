import pdfplumber
import pytesseract
from PIL import Image
import re
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# Function to extract text from images using OCR (if needed)
def extract_text_from_image(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # Convert the page to an image
            img = page.to_image()
            # Use pytesseract to extract text from the image
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

    # Use OCR if the text extraction is incomplete or empty
    if not text.strip():
        print("No text found using standard extraction. Switching to OCR.")
        text = extract_text_from_image(pdf_path)

    print("Extracted text for debugging:")
    print(text[:1000])  # Print the first 1000 characters for inspection
    return text

# Function to extract relevant details using regex
def extract_details(text):
    details = {
        'Goods Description': [],
        'Bags': [],
        'Quintal': [],
        'Rate': [],
        'Amount': [],
        'GSTIN': "",
        'GSTIN NO': "",
        'HSN/SAC': [],
        'Company Name': "",
        'Invoice No': "",
        'Date of Invoice': "",
        'Shipped to': "",
        'Transport': "",
        'Vehicle No': "",
        'Licence No': "",
        'Mobile No': ""
    }

    try:
        # Extract GSTIN
        gstin_pattern = r"GSTIN NO\s*:\s*([0-9A-Z]+)"
        gstin_match = re.search(gstin_pattern, text)
        details['GSTIN NO'] = gstin_match.group(1) if gstin_match else ""

        # Extract company name without using look-behind
        company_name_pattern = r"TAX INVOICE\s+(.*?)\s+At"
        company_name_match = re.search(company_name_pattern, text)
        details['Company Name'] = company_name_match.group(1).strip() if company_name_match else ""

        # Extract Invoice No
        invoice_no_pattern = r"Invoice No\.\s*:\s*(\d+)"
        invoice_no_match = re.search(invoice_no_pattern, text)
        details['Invoice No'] = invoice_no_match.group(1) if invoice_no_match else ""

        # Extract Date of Invoice
        date_of_invoice_pattern = r"Date of Invoice\s*:\s*([\d\- ]+)"
        date_of_invoice_match = re.search(date_of_invoice_pattern, text)
        details['Date of Invoice'] = date_of_invoice_match.group(1).strip() if date_of_invoice_match else ""

        # Extract Shipped to address
        shipped_to_pattern = r"Shipped to\s*:\s*(.*?)\s+GSTIN"
        shipped_to_match = re.search(shipped_to_pattern, text)
        details['Shipped to'] = shipped_to_match.group(1).strip() if shipped_to_match else ""

        # Extract goods details
        goods_description_match = re.findall(
            r'\d+\.\s+(.*?)\s+(\d{8})\s+([\d\.]+)\s+([\d\.]+)\s+([\d,]+)', 
            text
        )

        if goods_description_match:
            for item in goods_description_match:
                details['Goods Description'].append(item[0].strip())
                details['HSN/SAC'].append(item[1].strip())
                details['Bags'].append(item[2].strip())
                details['Quintal'].append(item[3].strip())
                details['Rate'].append(item[4].strip())
                # For Amount, you might need a specific extraction logic based on your invoice format.
                amount_pattern = r'(?<=\s)\d{1,3}(?:,\d{3})*(?:\.\d{2})?'
                amount_matches = re.findall(amount_pattern, text)
                details['Amount'].append(amount_matches.pop(0) if amount_matches else "0")

            print(f"Extracted Goods: {details['Goods Description']}")  # Debugging output
            
    except Exception as e:
        print(f"Error extracting data: {e}")

    return details

# Function to write extracted details to an Excel file
def write_to_excel(details, output_path):
    goods_df = pd.DataFrame({
        'Goods Description': details['Goods Description'],
        'Bags': details['Bags'],
        'Quintal': details['Quintal'],
        'Rate': details['Rate'],
        'Amount': details['Amount']
    })

    # Ensure HSN/SAC matches length of goods
    hsn_codes = details.get('HSN/SAC', [])[:len(goods_df)]
    goods_df['HSN/SAC'] = hsn_codes

    # Ensure all other fields are repeated for each row
    for column in ['Company Name', 'Invoice No', 'Date of Invoice', 'GSTIN NO', 
                   'GSTIN', 'Shipped to', 'Transport', 'Vehicle No', 
                   'Licence No', 'Mobile No']:
        value = details.get(column, "")
        if isinstance(value, list):  # If it's a list, ensure its length matches the DataFrame
            if len(value) == len(goods_df):
                goods_df[column] = value
            else:
                print(f"Warning: Length mismatch for column '{column}'. Skipping.")
        else:
            goods_df[column] = [value] * len(goods_df)  # Broadcast single value to all rows

    # Save to Excel
    goods_df.to_excel(output_path, index=False)


# Main execution
def main():
    print("Please select the PDF file:")
    Tk().withdraw()  # Hides the root window
    pdf_path = askopenfilename(filetypes=[("PDF files", "*.pdf")])

    if pdf_path:
        text = extract_full_text(pdf_path)

        # If the text is not extracted properly, use OCR
        if not text.strip():
            print("No text found in the PDF. Using OCR for text extraction.")
            text = extract_text_from_image(pdf_path)

        print("Extracted Text for Debugging:")
        print(text[:1000])  # Print first 1000 characters of extracted text

        details = extract_details(text)

        # Check if any important fields are empty before writing to Excel
        if not details['Goods Description']:
            print("No goods descriptions found; nothing will be written to Excel.")
            return
        
        excel_path = pdf_path.replace('.pdf', '_details.xlsx')
        
        write_to_excel(details, excel_path)
        
        print(f"Extracted details saved to {excel_path}")
    else:
        print("No file selected.")

if __name__ == "__main__":
    main()

