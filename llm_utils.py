import os
from dotenv import load_dotenv
from groq import Groq
import pprint

# Load environment variables
load_dotenv()
GROQ_API_KEY = os.getenv("GROQ_API_KEY")

def initialize_groq_client():
    """Initialize and return the Groq client"""
    if not GROQ_API_KEY:
        raise ValueError("❌ GROQ_API_KEY not found. Please add it to your .env file.")
    
    client = Groq(api_key=GROQ_API_KEY)
    print("✅ Groq client initialized.")
    return client


def get_user_preferences():
    """Get user input for presentation preferences"""
    topic = input("Enter presentation topic: ").strip()
    if not topic:
        topic = "Artificial Intelligence"  # Default topic
    
    while True:
        detail_level = input("Choose detail level (simple/detailed) [default: simple]: ").strip().lower()
        if not detail_level:
            detail_level = "simple"
            break
        if detail_level in ["simple", "detailed"]:
            break
        print("Please enter 'simple' or 'detailed'")
    
    return topic, detail_level

def get_presentation_outline(client, topic: str, detail_level: str = "simple") -> str:
    """
    Generate a presentation outline using Groq LLM with option for simple or detailed output.
    
    Args:
        client: Initialized Groq client
        topic (str): The topic of the presentation.
        detail_level (str): "simple" for 3-5 slides, "detailed" for more comprehensive outline.
    
    Returns:
        str: The generated outline text.
    """
    if detail_level == "simple":
        prompt = f"Generate a concise presentation outline with 3-5 slides about {topic}. Use clear section headers marked with ** around them and bullet points with - for each slide."
    else:  # detailed
        prompt = f"Generate a comprehensive presentation outline with 7-10 slides about {topic}, including detailed points for each slide. Use clear section headers marked with ** around them and bullet points with - for each slide."
    
    try:
        response = client.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[
                {"role": "system", "content": "You are an AI presentation assistant. Create well-structured presentation outlines with clear sections marked with ** and bullet points starting with -."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            timeout=30  # Add timeout to prevent hanging
        )
        return response.choices[0].message.content
    except Exception as e:
        raise Exception(f"❌ Failed to generate outline: {str(e)}")

import re

# In your llm_utils.py file, replace the parse_llm_output_to_outline function with:

def parse_llm_output_to_outline(llm_output):
    """
    Convert LLM output to a structured dictionary for PowerPoint generation.
    This version handles nested bullet points.
    
    Args:
        llm_output (str): The raw output from the LLM
    
    Returns:
        dict: Structured outline with sections as keys and nested content as values
    """
    outline = {}
    current_section = None
    
    for line in llm_output.splitlines():
        line = line.strip()
        if not line:
            continue
        
        # Detect section headers (formatted with ** or Slide X:)
        if (line.startswith("**") and line.endswith("**")) or line.startswith("Slide"):
            # Clean up the section title
            clean_line = re.sub(r'\*\*', '', line)
            if clean_line.startswith("Slide"):
                # Extract just the title part after "Slide X:"
                clean_line = re.sub(r'^Slide\s+\d+:\s*', '', clean_line)
            current_section = clean_line.strip()
            outline[current_section] = []
        
        # Detect bullet points
        elif current_section and line.startswith("-"):
            clean_point = line.lstrip("- ").strip()
            # Remove any remaining markdown formatting
            clean_point = re.sub(r'\*\*', '', clean_point)
            outline[current_section].append(clean_point)
    
    return outline


def get_mock_outline(topic, detail_level):
    """Return a mock outline for testing when API is unavailable"""
    if detail_level == "simple":
        return f"""
**Introduction to {topic}**
- What is {topic}?
- Why is {topic} important?
- Key concepts overview

**Applications of {topic}**
- Real-world examples
- Industry impact
- Future potential

**Conclusion**
- Summary of key points
- Final thoughts
- Q&A
"""
    else:
        return f"""
Introduction to {topic}
- Definition and background
- Historical context
- Importance in modern world
- Key terminology

**Core Concepts of {topic}**
- Fundamental principles
- Theoretical foundations
- Key components and elements
- Relationships between concepts

**Applications of {topic}**
- Industry use cases
- Real-world implementations
- Success stories
- Case studies

**Benefits and Advantages**
- Economic impact
- Efficiency improvements
- Quality enhancements
- Competitive advantages

**Challenges and Limitations**
- Implementation barriers
- Technical limitations
- Cost considerations
- Adoption challenges

**Future Trends**
- Emerging developments
- Research directions
- Market predictions
- Innovation opportunities

**Conclusion and Recommendations**
- Summary of key insights
- Strategic recommendations
- Call to action
- References
"""

def generate_outline(topic=None, detail_level=None):
    """Main function to generate an outline with user preferences or provided arguments."""
    try:
        # Use provided arguments or get user preferences
        if not topic or not detail_level:
            topic, detail_level = get_user_preferences()
        
        # Initialize client and generate outline
        client = initialize_groq_client()
        outline_text = get_presentation_outline(client, topic, detail_level)
        
        print(f"\nGenerated {detail_level} outline for '{topic}':\n")
        print(outline_text)
        
        # Parse to dictionary
        outline_dict = parse_llm_output_to_outline(outline_text)
        
        print("\nParsed Outline Structure:")
        for section, points in outline_dict.items():
            print(f"{section}: {len(points)} bullet points")
        
        return outline_dict, topic
        
    except Exception as e:
        print(f"❌ Error: {e}")
        print("Using mock data instead...")
        
        # Use mock data if API fails
        if not topic or not detail_level:
            topic, detail_level = get_user_preferences()
        outline_text = get_mock_outline(topic, detail_level)
        outline_dict = parse_llm_output_to_outline(outline_text)
        
        print(f"\nMock {detail_level} outline for '{topic}':\n")
        print(outline_text)
        
        print("\nParsed Outline Structure:")
        for section, points in outline_dict.items():
            print(f"{section}: {len(points)} bullet points")
        
        return outline_dict, topic

if __name__ == "__main__":
    # Test the LLM functionality
    outline, topic = generate_outline()
    
    # Show the parsed outline
    print("\nFinal Parsed Outline:")
    pprint.pprint(outline)
