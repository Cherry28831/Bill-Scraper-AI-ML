import spacy
import pytesseract
import fitz  # PyMuPDF
from spacy.training.example import Example
import random

# Set the path to the Tesseract executable (change it to your local installation path)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# PDF Extraction Function: Extract text from a PDF file
def extract_text_from_pdf(pdf_path):
    # Open the provided PDF file
    doc = fitz.open(pdf_path)
    text = ""
    
    # Loop through all pages and extract text
    for page in doc:
        text += page.get_text()

    return text

# Train the NER model with invoice data
def train_ner_model(train_data):
    nlp = spacy.blank("en")  # Create a blank English model
    ner = nlp.add_pipe("ner", last=True)  # Add the NER component to the pipeline

    # Add the labels to the NER model
    for _, annotations in train_data:
        for ent in annotations.get("entities"):
            ner.add_label(ent[2])

    optimizer = nlp.begin_training()  # Start training
    for i in range(100):  # Set the number of iterations
        random.shuffle(train_data)  # Shuffle the training data
        losses = {}
        for text, annotations in train_data:
            example = Example.from_dict(nlp.make_doc(text), annotations)
            nlp.update([example], drop=0.5, losses=losses)  # Update the model
        print(f"Iteration {i+1} - Losses: {losses}")

    return nlp

# Function to extract the entities from the text using the trained NER model
def extract_entities_with_ner(nlp_model, text):
    doc = nlp_model(text)
    extracted_entities = []
    for ent in doc.ents:
        extracted_entities.append((ent.text, ent.label_))
    return extracted_entities

# Example Training Data
TRAINING_DATA = [
    (
        "Goods Description: RICE WADA KOLAM HSN/SAC: 10063090 Bags: 577 Pack: 0.260 Quintal: 150.02 Rate: 5500 Amount: 825110 Company Name: AVENUE SUPERMARTS LTD. (TURBHE) Invoice No: 210 FSSAI: 11518016000410 Date of Invoice: 14-07-2023 GSTIN NO: 27AAEPW8766H1ZY GSTIN: 27AACCA8432H1ZQ PAN NO: AAEPW8766 TAN NO: NGPVO-0770A STD: 07174-220370 Shipped to: AVENUE SUPERMARTS LTD. (TURBHE) Transport: SAIKRUPA LORRY ARRENGERS Place of Supply: Maharashtra Vehicle No: MH04GR1659 Licence No: 8614 Mobile No: 8355908709",
        {
            "entities": [
                (117, 120, "INVOICE_NO"),
                (147, 159, "DATE_OF_INVOICE"),
                (122, 145, "FSSAI"),
                (85, 115, "COMPANY_NAME"),
                (17, 34, "GOODS_DESCRIPTION"),
                (36, 46, "HSN_SAC"),
                (48, 51, "BAGS"),
                (53, 58, "PACK"),
                (60, 68, "QUINTAL"),
                (70, 75, "RATE"),
                (77, 83, "AMOUNT"),
                (161, 179, "GSTIN_NO"),
                (181, 197, "GSTIN"),
                (199, 207, "PAN_NO"),
                (209, 215, "TAN_NO"),
                (217, 230, "STD"),
                (232, 282, "SHIPPED_TO"),
                (284, 310, "TRANSPORT"),
                (312, 325, "PLACE_OF_SUPPLY"),
                (327, 338, "VEHICLE_NO"),
                (340, 344, "LICENCE_NO"),
                (346, 358, "MOBILE_NO")
            ]
        }
    )
]

# Main code to run PDF Extraction + NER pipeline
if __name__ == "__main__":
    # Path to your invoice PDF
    pdf_path = "tani 21.pdf"  # Replace with your PDF file path
    
    # Step 1: Extract text from the PDF
    extracted_text = extract_text_from_pdf(pdf_path)
    print("Extracted Text from PDF:")
    print(extracted_text)

    # Step 2: Train the NER model using the example training data
    print("Training the NER model...")
    trained_ner_model = train_ner_model(TRAINING_DATA)

    # Step 3: Extract entities using the trained NER model on the extracted PDF text
    print("\nExtracted Entities using NER:")
    entities = extract_entities_with_ner(trained_ner_model, extracted_text)
    
    # Print the entities detected by the model
    for entity in entities:
        print(f"{entity[0]} | {entity[1]}")
