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
    details = {}

    try:
        # Extract the last GSTIN
        gstin_pattern = r"\b[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[A-Z0-9]{1}[Z]{1}[A-Z0-9]{1}\b"
        gstin_matches = re.findall(gstin_pattern, text)
        details['GSTIN'] = gstin_matches[-1] if gstin_matches else ""

        # Extract other fields
        details['GSTIN NO'] = re.search(r'GSTIN\s*NO\s*[:\-]?\s*([\w\d]+)', text).group(1) if re.search(r'GSTIN\s*NO\s*[:\-]?\s*([\w\d]+)', text) else ""
        details['HSN/SAC'] = re.findall(r"\b[0-9]{8}\b", text)  # List of all HSN/SAC codes

        # Extract company name and other fields
        company_name_match = re.search(r'^[A-Z][A-Z &,-\.]+(?:\s+Ltd.*)?(?=\s+At\.)', text, re.MULTILINE)
        details['Company Name'] = company_name_match.group(0).strip() if company_name_match else ""
        details['Invoice No'] = re.search(r'Invoice No\.\s*[:\-]?\s*(\d+)', text).group(1) if re.search(r'Invoice No\.\s*[:\-]?\s*(\d+)', text) else ""
        details['Date of Invoice'] = re.search(r'Date of Invoice\s*[:\-]?\s*([\d\-]+)', text).group(1) if re.search(r'Date of Invoice\s*[:\-]?\s*([\d\-]+)', text) else ""

        # Extract "Shipped to" address
        shipped_to_match = re.search(r'Shipped to\s*[:\-]?\s*([\s\S]+?)\s*FSSAI', text)
        details['Shipped to'] = shipped_to_match.group(1).strip() if shipped_to_match else ""

        # Extract goods details
        goods_description_match = re.findall(r'([A-Za-z\s]+)\s+\d+\s+(\d+)\s+(\d+\.\d+)\s+(\d+\.\d+)\s+([\d,]+)', text)
        details['Goods Description'], details['Bags'], details['Quintal'], details['Rate'], details['Amount'] = zip(*goods_description_match) if goods_description_match else ([], [], [], [], [])

        # Transport and other fields
        details['Transport'] = re.search(r'Transport\s*[:\-]?\s*([\s\S]+?)\s+Despatch Date', text).group(1).strip() if re.search(r'Transport\s*[:\-]?\s*([\s\S]+?)\s+Despatch Date', text) else ""
        details['Vehicle No'] = re.search(r'Vehicle No\.\s*[:\-]?\s*([\w\d]+)', text).group(1) if re.search(r'Vehicle No\.\s*[:\-]?\s*([\w\d]+)', text) else ""
        details['Licence No'] = re.search(r'Licence No\s*[:\-]?\s*([\w\d]+)', text).group(1) if re.search(r'Licence No\s*[:\-]?\s*([\w\d]+)', text) else ""
        details['Mobile No'] = re.search(r'Mobile No\s*[:\-]?\s*([\d]+)', text).group(1) if re.search(r'Mobile No\s*[:\-]?\s*([\d]+)', text) else ""

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

    # Ensure all other fields are repeated for each row
    for column in ['Company Name', 'Invoice No', 'Date of Invoice', 'GSTIN NO', 'GSTIN', 'HSN/SAC', 
                   'Shipped to', 'Transport', 'Vehicle No', 'Licence No', 'Mobile No']:
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

        details = extract_details(text)
        excel_path = pdf_path.replace('.pdf', '_details.xlsx')
        write_to_excel(details, excel_path)
        print(f"Extracted details saved to {excel_path}")
    else:
        print("No file selected.")

if __name__ == "__main__":
    main()