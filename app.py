import streamlit as st
import pandas as pd
from io import BytesIO
from src.processing_utils import preprocess_text, extract_text_pymupdf, extract_text_with_ocr, is_scanned_pdf, search_full_keywords_in_chunks
from src.config import initialize_groq_client, KEYWORDS
from src.excel_and_main import save_results_to_excel

# Initialize Groq client
GROQ_API_KEY = "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" 
client = initialize_groq_client(GROQ_API_KEY)

# Streamlit app title and description
st.title("PDF Keyword Extraction and Report Generation")
st.write("Upload your PDF files and select keywords to generate a structured report in Excel format.")

# Step 1: File Upload
uploaded_files = st.file_uploader("Upload PDFs", type=["pdf"], accept_multiple_files=True)

# Step 2: Keyword Selection
st.write("Select or search for keywords to process:")
selected_keywords = st.multiselect(
    "Search and select keywords:",
    options=KEYWORDS,
    default=KEYWORDS[:3],  # Select the first 3 by default
    help="You can search for keywords and select multiple."
)

# Process button
if st.button("Process PDFs"):
    if not uploaded_files:
        st.error("Please upload at least one PDF file.")
    elif not selected_keywords:
        st.error("Please select at least one keyword.")
    else:
        st.info("Processing PDFs. This may take some time...")

        # Step 3: Process PDFs
        all_results = []
        for uploaded_file in uploaded_files:
            file_path = uploaded_file.name
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            # Check if scanned or regular PDF
            if is_scanned_pdf(file_path):
                chunks = extract_text_with_ocr(file_path)
            else:
                chunks = extract_text_pymupdf(file_path)

            chunks = preprocess_text(chunks)

            results = search_full_keywords_in_chunks(chunks, selected_keywords, client)
            all_results.extend(results)

        # Step 4: Save results to Excel
        st.success("Extraction Successfull converting to excel")
        output_excel_path = "output/output_results.xlsx"
        save_results_to_excel(all_results, output_excel_path)

        # Convert Excel to BytesIO for download
        output = BytesIO()
        with open(output_excel_path, "rb") as excel_file:
            output.write(excel_file.read())
        output.seek(0)

        # Step 5: Provide Downloadable Excel
        st.success("Processing complete! Download the results below.")
        st.download_button(
            label="Download Excel File",
            data=output,
            file_name="extracted_keywords_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
