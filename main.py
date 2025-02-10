from pptx import Presentation
from langchain_ollama import OllamaLLM
from langchain_core.prompts import ChatPromptTemplate

filePath = input("Enter file path (single = file, multiple = directory): ")

template = """"
    You are a highly intelligent and helpful study assistant. You help students by answering questions based on the provided course material.

    User course content (reference when answering questions): {content}

    Conversation history:
    {context}

    Question:
    {question}
    
    Answer:
    """
model = OllamaLLM(model="llama3")
prompt = ChatPromptTemplate.from_template(template)
chain = prompt | model

def extractText(filePath):
    file = Presentation(filePath)
    text = []
    for slide in file.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text.append(shape.text)
    return "\n".join(text)

def modelConversation():
    context = ""
    print("Type 'exit' to exit conversation.")
    # result = chain.invoke({"context": content, "question": question})
    # context += f"\nCourse Content: {content}\nAI: {result}"
    while True:
        print("Bot: What would you like help with?")
        userInput = input("You: ")
        if userInput.lower() == "exit":
            break

        result = chain.invoke({"content": content, "context": context, "question": userInput})
        print("Bot:", result)
        context += f"\nUser: {userInput}\nAI: {result}"


try:
    content = extractText(filePath)
    print("\nExtracted Text:\n", content[:50])
except:
    print("Error")

if __name__ == "__main__":
    modelConversation()