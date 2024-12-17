import pandas as pd
from langchain.document_loaders import PyPDFLoader
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain.vectorstores import Chroma
from langchain.embeddings import SentenceTransformerEmbeddings
import os
# from langchain_google_genai import GoogleGenerativeAIEmbeddings

# def load_and_split_documents(file_path):
#     loader = PyPDFLoader(file_path)
#     documents = loader.load()
#     text_splitter = RecursiveCharacterTextSplitter(chunk_size=300, chunk_overlap=0)
#     chunks = text_splitter.split_documents(documents)
#     return chunks

def load_and_split_documents_from_files(file_paths):
    all_chunks = []
    for file_path in file_paths:
        loader = PyPDFLoader(file_path)
        documents = loader.load()
        text_splitter = RecursiveCharacterTextSplitter(chunk_size=1000, chunk_overlap=200)
        chunks = text_splitter.split_documents(documents)
        all_chunks.extend(chunks)  # Aggregate chunks from all files
    return all_chunks


def create_vector_store(chunks):
    embedding_model = SentenceTransformerEmbeddings(model_name="all-MiniLM-L6-v2")
    # embedding_model = GoogleGenerativeAIEmbeddings(model="models/embedding-001",api_key="AIzaSyBHTV8_2Ul2nrKdLEht5BKWbQEkgIZvqIA")
    vectorstore = Chroma.from_documents(documents=chunks, embedding=embedding_model)
    return vectorstore


# def retrieve_chunks_with_keyword(vectorstore, keyword):
#     results = vectorstore.similarity_search(keyword,k=50)  # Adjust k based on expected matches
#     return [(keyword, chunk.page_content) for chunk in results]

def sanitize_value(value):
    if isinstance(value, str):
        # Replace illegal characters with a space or remove them
        return ''.join(c if ord(c) < 65536 else ' ' for c in value)
    return value

# Function to sanitize the entire DataFrame
def sanitize_dataframe(df):
    return df.applymap(sanitize_value)

def results_to_dataframe(results):
    df = pd.DataFrame(results, columns=["Keyword", "Chunk"])
    return df

def retrieve_chunks_with_keywords(vectorstore, keywords):
    results = []
    for keyword in keywords:
        keyword_results = vectorstore.similarity_search(keyword, k=10)  # Adjust k based on expected matches
        results.extend([(keyword, chunk.page_content) for chunk in keyword_results])
    return results

# file_path = r"C:\Users\sriharsha\Downloads\Attachment-3 ADD-50092-HC-838-63-11-003-A-MATERIALS FOR USE IN H2S CONTAINING ENVIRONMENTS.pdf"  # Replace with your file path
keyword = ["Floating", "Trunnion mounted", "centric","Triple offset","API 609","Monogram"]            
# Collect all PDF files from the directory
directory_path = "data"
file_paths = [os.path.join(directory_path, file) for file in os.listdir(directory_path) if file.endswith(".pdf")]


chunks = load_and_split_documents_from_files(file_paths)
vectorstore = create_vector_store(chunks)
results = retrieve_chunks_with_keywords(vectorstore, keyword)
df = results_to_dataframe(results)
file_name = "key_values.xlsx"


def save_dataframe_to_excel(df, file_name):
    try:
        sanitized_df = sanitize_dataframe(df)
        sanitized_df.to_excel(file_name, index=False)  # Ensure index isn't saved unless needed
        print(f"Data successfully saved to {file_name}")
    except Exception as e:
        print(f"An error occurred while saving the DataFrame: {e}")

df = pd.DataFrame(results, columns=["Keyword", "Chunk"])

# Save DataFrame to Excel
save_dataframe_to_excel(df, "key_values.xlsx")