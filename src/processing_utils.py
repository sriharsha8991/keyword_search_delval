# processing_utils.py

import fitz  # PyMuPDF
import time
import pandas as pd
import os
import re
import textwrap
from pdf2image import convert_from_path
import pytesseract
from src.config import KEYWORDS, MODEL

# -------------------------- TEXT EXTRACTION -------------------------- #

def extract_text_pymupdf(file_path: str) -> list:
    """Extract text from a regular PDF using PyMuPDF."""
    text_chunks = []
    with fitz.open(file_path) as pdf:
        for page in pdf:
            text_chunks.append(page.get_text("blocks"))  
    return [" ".join(block[4] for block in page_blocks) for page_blocks in text_chunks]

def extract_text_with_ocr(file_path: str) -> list:
    """Extract text from scanned PDFs using OCR."""
    images = convert_from_path(file_path, dpi=300)
    text_chunks = []
    for img in images:
        text = pytesseract.image_to_string(img, lang="eng")
        text_chunks.extend(text.split("\n\n"))  # Split into chunks
    return text_chunks

def is_scanned_pdf(file_path: str) -> bool:
    """Determine if a PDF is scanned by checking if it lacks text."""
    with fitz.open(file_path) as pdf:
        return any(page.get_pixmap() for page in pdf if len(page.get_text("blocks")) == 0)

# -------------------------- TEXT PREPROCESSING -------------------------- #

def preprocess_text(chunks: list) -> list:
    """Clean up and preprocess text chunks."""
    clean_chunks = []
    for chunk in chunks:
        chunk = chunk.strip()  
        chunk = " ".join(chunk.split())  
        clean_chunks.append(chunk)
    return clean_chunks

def format_chunk_general(chunk: str) -> str:
    """Format text chunks for better readability."""
    chunk = re.sub(r'(\.|\;|\:)\s+', r'\1\n', chunk)  # Add line breaks after punctuation
    chunk = re.sub(r'\b(IEC \d{4,}|NACE MR\d{4,}|ISO \d{4,})', r'\n  - \1', chunk)  # Bullet points
    formatted_text = "\n".join(textwrap.fill(line, width=80) for line in chunk.split("\n"))
    return formatted_text.strip()

# -------------------------- KEYWORD SEARCH -------------------------- #

def search_full_keywords_in_chunks(chunks: list, keywords: list, client) -> list:
    """
    Search for keywords in text chunks and highlight them.
    """
    results = []
    keyword_patterns = [re.compile(rf"\b{re.escape(keyword)}\b", re.IGNORECASE) for keyword in keywords]

    for chunk in chunks:
        highlighted_chunk = chunk  # Start with the original chunk
        for keyword, pattern in zip(keywords, keyword_patterns):
            if pattern.search(chunk):
                highlighted_chunk = pattern.sub(f"**{keyword}**", highlighted_chunk)
                generated_chunk = generate_report(client, highlighted_chunk)
                time.sleep(0.3)
                results.append({"Keyword": keyword, "generated_Chunk": generated_chunk, "Chunk": highlighted_chunk})
    return results

# -------------------------- GROQ INTEGRATION -------------------------- #

def generate_report(client, context: str) -> str:
    """Use Groq API to structure text content into readable format."""
    messages = [
        {
            "role": "user",
            "content": (
                f"""# Data Structuring Task

## Input
{context}

## Processing Requirements
- Restructure the provided text for maximum readability
- Maintain complete original content integrity
- Use clear, logical formatting techniques
- Prioritize visual clarity and comprehension

## Desired Output
- Structured, easily scannable document
- Preserved original meaning"""
            ),
        }
    ]
    chat_completion = client.chat.completions.create(
        messages=messages, model=MODEL, temperature=0.1
    )
    return chat_completion.choices[0].message.content
