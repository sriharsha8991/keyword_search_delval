# keyword_search_delval

This repo is only for single feature of the entire delval project which is a keyword search.

The .py files below shows different stages of development:
- nlp_use.py -- basic version cannot handle all the use cases 
- nlp_with_pymupdf.py --  has bettter keyword extraction compared to the previous version but also accounts for all the sub words which is not the clients requirement
- nlp_with_pymupdf2_new.py -- can be able retrieve information properly, but storing of the information in the excel is very raw and not a better readable format.
-  groq_version.py -- latest version so far can be able to pass the chunks to the llm and stores the structured outputs back into excel again in a much readable format
- chat_utils.py -- groq api chat completion set up using llama 3.2 8B LLM