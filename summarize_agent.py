from .agent_base import AgentBase

class SummarizeTool(AgentBase):
    def __init__(self, max_retries=2, verbose=True):
        super().__init__(name='SummarizeTool', max_retries=max_retries, verbose=verbose)

    def execute(self,text):
       messages = [
           {'role':"system", 'content':"you are an AI assistant that summarizes medical data"},
           {
               'role':'user',
               'content':(
                   "Please provide a concise summary of the following medical text:\n\n"
                   f"{text}\n\nSummary:"
               )
           }
       ]
       summary = self.call_groq(messages,max_tokens=300)
       return summary
    
