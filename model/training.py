import spacy
from spacy.training.example import Example
import random
import fitz  # PyMuPDF

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

def save_model(model, output_dir="invoice_ner_model"):
    model.to_disk(output_dir)  # Save the trained model to a directory
    print(f"Model saved to {output_dir}")

def extract_text_from_pdf(pdf_path):
    # Open the provided PDF file
    doc = fitz.open(pdf_path)
    text = ""
    
    # Loop through all pages and extract text
    for page in doc:
        text += page.get_text()

    return text

# Example of how to use it
if __name__ == "__main__":
    # Define the training data with the correct sequence
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

    # Train the model with the data
    trained_model = train_ner_model(TRAINING_DATA)

    # Save the trained model
    save_model(trained_model)

    # Extract text from a PDF (example.pdf is the path to your PDF file)
    pdf_text = extract_text_from_pdf("tani 21.pdf")

    # Ensure the entire text is concatenated into a single string before processing
    pdf_text = pdf_text.replace("\n", " ")  # Remove newlines and concatenate

    # Test the trained model with extracted text
    doc = trained_model(pdf_text)

    # Print the entities detected by the model
# Print the extracted entities in a table format
    print("\nExtracted Entities:")
    print("| Entity Text | Entity Label |")
    print("|-------------|--------------|")
    for ent in doc.ents:
        print(f"{ent.text} | {ent.label_}")
