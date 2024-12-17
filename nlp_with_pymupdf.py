import os
import pandas as pd
import fitz  # PyMuPDF
import pytesseract
from pdf2image import convert_from_path
from io import StringIO
from pdfminer.high_level import extract_text

# Define keywords
keywords = ["Spare Parts","Floating", "Trunnion mounted","Sriharsha"] 

# Function to extract text using PyMuPDF
def extract_text_pymupdf(file_path):
    text_chunks = []
    with fitz.open(file_path) as pdf:
        for page in pdf:
            text_chunks.append(page.get_text("blocks"))  # Extract text by blocks
    # print(text_chunks)
    return [" ".join(block[4] for block in page_blocks) for page_blocks in text_chunks]

# Function to handle scanned PDFs (OCR)
def extract_text_with_ocr(file_path):
    images = convert_from_path(file_path, dpi=300)
    text_chunks = []
    for img in images:
        text = pytesseract.image_to_string(img, lang="eng")
        text_chunks.extend(text.split("\n\n"))  # Split into chunks
    return text_chunks

# Function to determine whether a PDF is scanned
def is_scanned_pdf(file_path):
    with fitz.open(file_path) as pdf:
        return any(page.get_pixmap() for page in pdf if len(page.get_text("blocks")) == 0)

# Function to preprocess extracted text
def preprocess_text(chunks):
    clean_chunks = []
    for chunk in chunks:
        chunk = chunk.strip()  # Remove leading/trailing spaces
        chunk = " ".join(chunk.split())  # Normalize whitespace
        clean_chunks.append(chunk)
    return clean_chunks

# Function to search for keywords in chunks
def search_keywords_in_chunks(chunks, keywords):
    results = []
    for keyword in keywords:
        for chunk in chunks:
            if keyword.lower() in chunk.lower():
                results.append({"Keyword": keyword, "Chunk": chunk})
    return results

# Main function
def process_pdfs(directory, keywords):
    all_results = []
    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):
            file_path = os.path.join(directory, filename)
            print(f"Processing: {filename}")
            if is_scanned_pdf(file_path):
                chunks = extract_text_with_ocr(file_path)
            else:
                chunks = extract_text_pymupdf(file_path)
            
            chunks = preprocess_text(chunks)
            results = search_keywords_in_chunks(chunks, keywords)
            all_results.extend(results)
    return all_results

# Directory containing PDF files
pdf_directory = "data"

# Process PDFs and store results
results = process_pdfs(pdf_directory, keywords)

# Save results to Excel
df = pd.DataFrame(results)
df.to_excel("extracted_keywords.xlsx", index=False)

print("Extraction complete. Results saved to 'extracted_keywords.xlsx'.")
