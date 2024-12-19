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

# Function to extract relevant details using regex
def extract_details(text):
    details = {}

    try:
        # Extract company name (assumes the first uppercase block near the top is the company name)
        company_name_match = re.search(r'^[A-Z][A-Z &,-\.]+(?:\s+Ltd.*)?(?=\s+At\.)', text, re.MULTILINE)
        details['Company Name'] = company_name_match.group(0).strip() if company_name_match else ""
        # Extracting other fields
        details['Invoice No'] = re.search(r'Invoice No\.\s*[:\-]?\s*(\d+)', text).group(1) if re.search(r'Invoice No\.\s*[:\-]?\s*(\d+)', text) else ""
        details['Date of Invoice'] = re.search(r'Date of Invoice\s*[:\-]?\s*([\d\-]+)', text).group(1) if re.search(r'Date of Invoice\s*[:\-]?\s*([\d\-]+)', text) else ""
        details['GSTIN NO'] = re.search(r'GSTIN\s*NO\s*[:\-]?\s*([\w\d]+)', text).group(1) if re.search(r'GSTIN\s*NO\s*[:\-]?\s*([\w\d]+)', text) else ""
        details['GSTIN'] = re.search(r'GSTIN\s*[:\-]?\s*([\w\d]+)', text).group(1) if re.search(r'GSTIN\s*[:\-]?\s*([\w\d]+)', text) else ""
        details['HSN/SAC code'] = re.search(r'HSN/SAC\s*[:\-]?\s*(\d+)', text).group(1) if re.search(r'HSN/SAC\s*[:\-]?\s*(\d+)', text) else ""
        
        # Extract Shipped to address
        shipped_to_match = re.search(r'Shipped to\s*[:\-]?\s*([\s\S]+?)\s*FSSAI', text)
        details['Shipped to'] = shipped_to_match.group(1).strip() if shipped_to_match else ""

        # Goods Description, Bags, Quintal, Rate, Amount
        goods_description_match = re.findall(r'([A-Za-z\s]+)\s+10063090\s+(\d+)\s+(\d+\.\d+)\s+(\d+\.\d+)\s+(\d+\,\d+)', text)

        # Initialize separate lists for each field
        goods_descriptions = []
        bags = []
        quintals = []
        rates = []
        amounts = []

        for match in goods_description_match:
            goods_descriptions.append(match[0].strip())
            bags.append(match[1])
            quintals.append(match[2])
            rates.append(match[3])
            amounts.append(match[4])

        # Add them as separate columns to the details dictionary
        details['Goods Description'] = goods_descriptions
        details['Bags'] = bags
        details['Quintal'] = quintals
        details['Rate'] = rates
        details['Amount'] = amounts

        # Transport, Vehicle No, Licence No, Mobile No
        details['Transport'] = re.search(r'Transport\s*[:\-]?\s*([\s\S]+?)\s+Despatch Date', text).group(1).strip() if re.search(r'Transport\s*[:\-]?\s*([\s\S]+?)\s+Despatch Date', text) else ""
        details['Vehicle No'] = re.search(r'Vehicle No\.\s*[:\-]?\s*([\w\d]+)', text).group(1) if re.search(r'Vehicle No\.\s*[:\-]?\s*([\w\d]+)', text) else ""
        details['Licence No'] = re.search(r'Licence No\s*[:\-]?\s*([\w\d]+)', text).group(1) if re.search(r'Licence No\s*[:\-]?\s*([\w\d]+)', text) else ""
        details['Mobile No'] = re.search(r'Mobile No\s*[:\-]?\s*([\d]+)', text).group(1) if re.search(r'Mobile No\s*[:\-]?\s*([\d]+)', text) else ""

    except Exception as e:
        print(f"Error extracting data: {e}")
    
    return details

# Function to write extracted details to an Excel file
def write_to_excel(details, output_path):
    df = pd.DataFrame([details])
    df.to_excel(output_path, index=False)

    goods_df = pd.DataFrame({
        'Goods Description': details['Goods Description'],
        'Bags': details['Bags'],
        'Quintal': details['Quintal'],
        'Rate': details['Rate'],
        'Amount': details['Amount']
    })
    
    # Merge the rest of the details into the goods dataframe
    for column in ['Company Name', 'Invoice No', 'Date of Invoice', 'GSTIN NO', 'GSTIN', 'HSN/SAC code', 'Shipped to', 'Transport', 'Vehicle No', 'Licence No', 'Mobile No']:
        goods_df[column] = details.get(column, "")
    
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