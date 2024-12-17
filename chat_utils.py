import streamlit as st
import os
from groq import Groq
import pandas as pd

# Function to initialize the Groq client
def initialize_groq_client(api_key: str) -> Groq:
    """
    Initialize the Groq client with the provided API key.
    """
    return Groq(api_key=api_key)

def generate_summary_report(client: Groq, context: str) -> str:
    messages = [
        {
            "role": "user",
            "content": f"""You are an expert in structurising the given data into much readable formats for humans by properly arranging them into readable sections.

            High priority: You are not allowed to add any kind of new content or knowledge to the given context. you just need to make it readable
            
            Here is the context for you in a text: {context}
            **Your answer should be in a strucutred  and readable format**
            """,
        }
    ]
    model= "llama3-8b-8192"
    
    # Make the API call
    chat_completion = client.chat.completions.create(
        messages=messages,
        model=model,
        temperature = 0.1
    )
    
    # Return the response content
    return chat_completion.choices[0].message.content
def get_text(excel_path):
    text_list = []
    df = pd.read_excel(excel_path)
    for i in df["Chunk"]:
        text_list.append(i)
    return text_list

def generate_report(): 
    api_key = "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"  # Replace with your actual API key
    client = initialize_groq_client(api_key)

    try:  
        ls = get_text(r"C:\Users\sriharsha\Desktop\Data Smith AI\keyword_search\extracted_keywords.xlsx")
        new_ls = []
        for context in ls:
            report = generate_summary_report(client, context)
            new_ls.append(report)
    except Exception as e:
        print(e)
    
    
