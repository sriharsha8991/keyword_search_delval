import openai
from abc import ABC, abstractmethod
from loguru import logger
import os


from groq import Groq

client = Groq(
    api_key="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",
)

class AgentBase(ABC):
    def __init__(self,name, max_retries=2,verbose=True):
        self.name = name
        self.max_retries = max_retries
        self.verbose = verbose
    
    @abstractmethod
    def execute(self,*args,**kwargs):
        pass

    def call_groq(self,messages,temperature=0.7,max_tokens=150):
        retries = 0
        while retries<self.max_retries:
            try:
                if self.verbose:
                    logger.info(f"{self.name} Sending messages to openai")

                    for msg in messages:
                        logger.debug(f"{msg['role']}: {msg['content']}")

                response = client.chat.completions.create(

    messages=messages,
    model="llama3-8b-8192",
    temperature=temperature,
    max_tokens=max_tokens,
    top_p=1,
    stop=None,
    stream=False
)
                reply = response.choices[0].message.content
                if self.verbose:
                    logger.info(f"{self.name} recived response : {reply}")
                return reply
            
            except Exception as e:
                retries += 1
                logger.error(f"[{self.name}] Error during open ai call:{e} Retry {retries}/{self.max_retries}")
        raise Exception(f"[{self.name}] failed to get response from Open A after {self.max_retries} retries")