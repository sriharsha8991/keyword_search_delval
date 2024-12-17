import os
import pandas as pd
from langchain.document_loaders import PyPDFLoader
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain.vectorstores import Chroma
from langchain.embeddings import SentenceTransformerEmbeddings

# Step 1: Load and Split Documents from Multiple Files
def load_and_split_documents_from_files(file_paths):
    """
    Loads and splits documents into manageable chunks from the given file paths.
    """
    all_chunks = []
    for file_path in file_paths:
        loader = PyPDFLoader(file_path)
        documents = loader.load()
        text_splitter = RecursiveCharacterTextSplitter(chunk_size=1000, chunk_overlap=0)
        chunks = text_splitter.split_documents(documents)
        all_chunks.extend(chunks)  # Aggregate chunks from all files
    return all_chunks

# Step 2: Create Vector Store
def create_vector_store(chunks):
    """
    Creates a vector store from the given document chunks using embeddings.
    """
    embedding_model = SentenceTransformerEmbeddings(model_name="all-MiniLM-L6-v2")
    vectorstore = Chroma.from_documents(documents=chunks, embedding=embedding_model)
    return vectorstore

# Step 3: Retrieve Chunks with Keywords
def retrieve_chunks_with_keywords(vectorstore, keywords):
    """
    Retrieves document chunks that match any of the given keywords.
    """
    results = []
    for keyword in keywords:
        keyword_results = vectorstore.similarity_search(keyword, k=10)  # Adjust 'k' as needed
        results.extend([(keyword, chunk.page_content) for chunk in keyword_results])
    return results

# Step 4: Sanitize DataFrame for Excel
def sanitize_value(value):
    """
    Removes illegal characters from strings to make them Excel-compatible.
    """
    if isinstance(value, str):
        return ''.join(c if ord(c) < 65536 else ' ' for c in value)
    return value

def sanitize_dataframe(df):
    """
    Applies sanitization to the entire DataFrame.
    """
    return df.applymap(sanitize_value)

# Step 5: Save DataFrame to Excel
def save_dataframe_to_excel(df, file_name):
    """
    Saves the DataFrame to an Excel file, handling potential issues with illegal characters.
    """
    try:
        sanitized_df = sanitize_dataframe(df)
        sanitized_df.to_excel(file_name, index=False)  # Ensure index isn't saved unless needed
        print(f"Data successfully saved to {file_name}")
    except Exception as e:
        print(f"An error occurred while saving the DataFrame: {e}")

# Main Execution
if __name__ == "__main__":
    # Define the directory containing PDFs and the list of keywords
    directory_path = "data"  # Replace with your directory path
    keywords = ["Reinforced", "Trunnion mounted", "Pumps"]

    # Collect all PDF file paths from the directory
    file_paths = [os.path.join(directory_path, file) for file in os.listdir(directory_path) if file.endswith(".pdf")]

    # Process: Load, split, and create vector store
    print("Loading and splitting documents...")
    chunks = load_and_split_documents_from_files(file_paths)

    print("Creating vector store...")
    vectorstore = create_vector_store(chunks)

    # Retrieve chunks for the given keywords
    print("Retrieving chunks for keywords...")
    results = retrieve_chunks_with_keywords(vectorstore, keywords)

    # Convert results to DataFrame
    print("Converting results to DataFrame...")
    df = pd.DataFrame(results, columns=["Keyword", "Chunk"])

    # Save DataFrame to Excel
    file_name = "key_Chunks.xlsx"
    print("Saving DataFrame to Excel...")
    save_dataframe_to_excel(df, file_name)
