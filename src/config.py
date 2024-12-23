# config.py

import os
import re
import textwrap
from groq import Groq


# -------------------------- CONFIGURATION -------------------------- #

# Define keywords to search for
KEYWORDS = [
    "SIL", "NACE", "IGC", "UT", "EN10204-3.2", "HARDNESS TESTING", 
    "ISO 5208", "FET", "PMI", "TAT"
]

# Groq API Configuration
GROQ_API_KEY = "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"  
MODEL = "llama3-8b-8192"

# Input and Output Paths
PDF_DIRECTORY = "data"  # Replace with the directory containing PDFs
EXCEL_OUTPUT_PATH = "output/extracted_keywords_with_groq.xlsx"

# -------------------------- INITIALIZATION -------------------------- #

def initialize_groq_client(api_key: str) -> Groq:
    """Initialize the Groq client with the provided API key."""
    return Groq(api_key=api_key)
