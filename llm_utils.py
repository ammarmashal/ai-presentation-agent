import os
from dotenv import load_dotenv
from groq import Groq


load_dotenv()
GROQ_API_KEY = os.getenv("GROQ_API_KEY")

if not GROQ_API_KEY:
    raise ValueError("❌ GROQ_API_KEY not found. Please add it to your .env file.")

client = Groq(api_key=GROQ_API_KEY)

def get_presentation_outline(topic: str) -> str:
    """
    Generate a simple presentation outline using Groq LLM.
    Args:
        topic (str): The topic of the presentation.
    Returns:
        str: The generated outline text.
    """
    prompt = f"Generate a presentation outline with 3 slides about {topic}."

    response = client.chat.completions.create(
        model="llama3-8b-8192",   # ❌ You can choose other models as needed
        messages=[{"role": "user", "content": prompt}],
        max_tokens=500,
    )

    return response.choices[0].message.content


if __name__ == "__main__":
    topic = "Artificial Intelligence"
    outline = get_presentation_outline(topic)
    print("Generated Outline:\n")
    print(outline)
