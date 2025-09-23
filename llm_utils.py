import os
from dotenv import load_dotenv
from groq import Groq
import pprint
import re
import random
import argparse


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

def extract_topic_from_input(user_input):
    """
    Extract the main topic from user input more intelligently
    """
    if not user_input:
        return "Artificial Intelligence"
    
    # Convert to lowercase for processing
    user_input = user_input.lower().strip()
    
    # Common patterns that indicate the actual topic follows
    topic_patterns = [
        r'(?:about|on|regarding|concerning|related to|for)\s+([^\.\?\!]+)',
        r'(?:presentation|ppt|slides|talk|speech)\s+(?:about|on|regarding)\s+([^\.\?\!]+)',
        r'(?:give me|create|generate|make|build)\s+(?:a\s+)?(?:presentation|ppt|slides)\s+(?:about|on|regarding)?\s*([^\.\?\!]+)',
        r'(?:i want|i need)\s+(?:a\s+)?(?:presentation|ppt|slides)\s+(?:about|on|regarding)?\s*([^\.\?\!]+)'
    ]
    
    # Try to extract topic using patterns
    for pattern in topic_patterns:
        match = re.search(pattern, user_input)
        if match:
            topic = match.group(1).strip()
            # Remove any remaining command words
            topic = re.sub(r'^(?:about|on|regarding|for|a|the)\s+', '', topic)
            # Capitalize properly (title case)
            topic = ' '.join(word.capitalize() for word in topic.split())
            return topic if topic else "Artificial Intelligence"
    
    # Fallback: remove common command phrases and keep the rest
    command_phrases = [
        'give me', 'can you', 'could you', 'please', 'create', 'generate',
        'make', 'build', 'i want', 'i need', 'a presentation', 'a ppt',
        'powerpoint', 'slides', 'presentation', 'about', 'on', 'regarding',
        'for', 'the'
    ]
    
    cleaned = user_input
    for phrase in command_phrases:
        cleaned = re.sub(r'\b' + phrase + r'\b', '', cleaned)
    
    # Clean up and capitalize
    cleaned = re.sub(r'[^\w\s]', '', cleaned)
    cleaned = cleaned.strip()
    if cleaned:
        cleaned = ' '.join(word.capitalize() for word in cleaned.split())
    
    return cleaned if cleaned else "Artificial Intelligence"

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
    Generate a presentation outline using Groq LLM with a generated title and specific constraints.
    """
    system_prompt = (
        "You are an expert presentation assistant. Your task is to create a professional presentation outline. "
        "Strictly adhere to the user's formatting and content constraints, including bullet point and sentence length. "
        "Each slide should be concise and focused on a single idea. Use markdown headers for slide titles."
    )

    if detail_level == "simple":
        prompt = f"""
        Generate a concise, professional presentation outline about **{topic}**. The presentation should have 30 slides. Adhere to the following strict formatting rules:
        
        1. The first line of your response must be a concise, one-line of 1-2 words title for the presentation, formatted with a single asterisk on each side, like this: *Presentation Title*.
        2. The second line must be blank.
        3. The first slide is "**Introduction**". It must have 2-3 full sentences of introductory text, not bullet points.
        4. Each subsequent slide title must be on its own line surrounded by double asterisks: **Slide Title**
        5. All other slides must have exactly 8-10 very short bullet points
        6. Each bullet point should start with • and be a concise keyword/phrase
        7. The final slide must be "**Conclusion**"

        EXAMPLE:
        *Digital Twins*

        **Introduction**
        This is the introductory text about the topic.

        **Core Concepts**
        • Concept 1
        • Concept 2  
        • Concept 3
        • Concept 4
        • Concept 5
        • Concept 6
        • Concept 7


        **Applications**
        • App 1
        • App 2
        • App 3
        • App 4
        • App 5
        • App 6
        • App 7


        **Conclusion**
        • Summary points 

        Now generate the outline for: {topic}
        """
    
    else:  # detailed
        prompt = f"""
        Generate a comprehensive, professional presentation outline about **{topic}**. 

        STRICT FORMATTING RULES:
        1. The first line of your response must be a concise, one-line of 1-2 words title for the presentation, formatted with a single asterisk on each side, like this: *Presentation Title*.
        2. The second line must be blank.
        3. The first slide must be "**Introduction**" with 3-5 full sentences (no bullet points)
        4. Each subsequent slide title must be on its own line surrounded by double asterisks: **Slide Title**
        5. Under each slide title, provide exactly 2 detailed bullet points
        6. Each bullet point must start with • and be a meaningful sentence (15-30 words)
        7. The final slide must be "**Conclusion**"
        8. The presentation should have 12-15 slides total

        EXAMPLE:
        *Digital Twins*

        **Introduction**
        This is the introductory text with full sentences about the topic. It provides context and sets the stage for the presentation. The introduction should engage the audience.

        **Core Concepts**
        • First detailed bullet point explaining a core concept
        • Second detailed bullet point with additional information
        • Third bullet point completing the explanation

        **Implementation**
        • First implementation step or consideration
        • Second important implementation aspect

        **Conclusion**
        • Summary point about the presentation content about 20-30 words 

        Now generate the outline for: {topic}"""

    try:
        response = client.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            timeout=45
        )
        return response.choices[0].message.content
    except Exception as e:
        raise Exception(f"❌ Failed to generate outline: {str(e)}")

def process_bullet_points(content_lines):
    """
    Process content lines into structured bullet points with proper level detection
    """
    bullets = []
    
    for line in content_lines:
        if not line.strip():
            continue
            
        # Detect indentation level (2 spaces = 1 level)
        leading_spaces = len(line) - len(line.lstrip())
        level = leading_spaces // 2  # 2 spaces per level
        
        # Clean the line from markdown and bullet indicators
        cleaned = clean_markdown(line.strip())
        
        # Remove bullet indicators (•, -, *, numbered)
        cleaned = re.sub(r'^[•\-*]\s+', '', cleaned)  # Remove bullet symbols
        cleaned = re.sub(r'^\d+[\.\)]\s+', '', cleaned)  # Remove numbered bullets
        
        if cleaned:  # Only add non-empty content
            bullets.append({
                "text": cleaned,
                "level": max(0, min(level, 3))  # Limit levels to 0-3
            })
    
    return bullets

def clean_markdown(text):
    """
    Remove markdown formatting from text (preserve structure for parsing)
    """
    if not text:
        return text
    
    # Remove bold/italic but preserve for title detection
    text = re.sub(r'\*\*(.*?)\*\*', r'\1', text)  # Remove **bold**
    text = re.sub(r'\*(.*?)\*', r'\1', text)      # Remove *italic*
    
    # Remove other markdown elements
    text = re.sub(r'#+\s*', '', text)  # Remove headers
    text = re.sub(r'\[([^\]]+)\]\([^)]+\)', r'\1', text)  # Remove links
    text = re.sub(r'!\[([^\]]+)\]\([^)]+\)', '', text)  # Remove images
    
    # Clean up extra spaces
    text = re.sub(r'\s+', ' ', text).strip()
    
    return text
def parse_llm_output_to_outline(llm_output: str):
    """
    Parses LLM output to extract a presentation title and the slide outline.
    """
    lines = llm_output.splitlines()
    presentation_title = ""
    outline = {}

    # Extract the presentation title from the first line
    if lines:
        first_line = lines[0].strip()
        title_match = re.match(r'^\*(.+?)\*$', first_line)
        if title_match:
            presentation_title = title_match.group(1).strip()
            # Remove the title and blank line from the content
            content_lines = lines[2:] 
        else:
            # Fallback if title format is not matched
            presentation_title = "Presentation"
            content_lines = lines
    else:
        presentation_title = "Presentation"
        content_lines = []

    current_slide = None
    slide_content = []
    
    # Pattern to detect slide titles: **Title** or # Title formats
    slide_pattern = r'^\s*(?:\*\*([^*]+)\*\*|#\s+([^#]+))$'
    
    for line in content_lines:
        line = line.strip()
        if not line:
            continue
        
        # Check if this is a slide title
        title_match = re.match(slide_pattern, line)
        
        if title_match:
            slide_title = title_match.group(1) or title_match.group(2)
            slide_title = slide_title.strip()
            
            # Save previous slide content if it exists
            if current_slide:
                outline[current_slide] = process_bullet_points(slide_content)
            
            # Start a new slide
            current_slide = slide_title
            slide_content = []
            
        else:
            # Add content to the current slide
            if current_slide:
                slide_content.append(line)
    
    # Add the final slide
    if current_slide:
        outline[current_slide] = process_bullet_points(slide_content)
    
    # Ensure Introduction slide is the first if it exists
    if "Introduction" in outline:
        ordered_outline = {"Introduction": outline.pop("Introduction")}
        ordered_outline.update(outline)
        outline = ordered_outline
    
    return presentation_title, outline


# used when Groq API is unavailable
def get_mock_outline(topic, detail_level):
    """Return a mock outline for testing when API is unavailable"""
    if detail_level == "simple":
        # Simple presentation with 7-12 slides worth of content
        title = f"**{topic}**"
        sections = [
            "Introduction",
            "Core Concepts",
            "Key Features", 
            "Applications",
            "Benefits",
            "Implementation",
            "Case Studies",
            "Best Practices",
            "Future Trends",
            "Conclusion"
        ]
        
        content = f"{title}\n\n"
        content += f"**{sections[0]}**\n"
        content += f"- Brief overview of {topic} and its significance\n- Main objectives and goals\n- Target audience and use cases\n\n"
        
        for i, section in enumerate(sections[1:-1], 1):
            content += f"**{section}**\n"
            content += f"- Key point 1 about {section.lower()}\n- Key point 2 about {section.lower()}\n- Key point 3 about {section.lower()}\n\n"
        
        content += f"**{sections[-1]}**\n"
        content += "- Summary of main points\n- Key takeaways\n- Next steps and recommendations"
        
        return content
        
    else:
        # Detailed presentation with 10-15 slides worth of content
        title = f"**Comprehensive Analysis of {topic}**"
        sections = [
            "Introduction",
            "Historical Background",
            "Fundamental Principles",
            "Technical Architecture", 
            "Key Components",
            "Implementation Methods",
            "Industry Applications",
            "Success Stories",
            "Benefits and Advantages",
            "Challenges and Limitations",
            "Future Developments",
            "Best Practices",
            "Case Study Analysis",
            "Conclusion and Recommendations"
        ]
        
        content = f"{title}\n\n"
        content += f"**{sections[0]}**\n"
        content += f"- Comprehensive overview of {topic} and its significance in modern context\n"
        content += f"- Detailed explanation of core concepts and their interrelationships\n"
        content += f"- Discussion of the evolution and current state of {topic}\n\n"
        
        for i, section in enumerate(sections[1:-1], 1):
            content += f"**{section}**\n"
            content += f"- In-depth analysis of first aspect of {section.lower()}\n"
            content += f"- Detailed examination of second aspect with specific examples\n"
            content += f"- Comprehensive review of third aspect including practical implications\n\n"
        
        content += f"**{sections[-1]}**\n"
        content += "- Detailed summary of all key insights and findings\n"
        content += "- Specific recommendations for implementation and adoption\n"
        content += "- Future outlook and potential developments in the field"
        
        return content



def generate_outline(topic=None, detail_level=None):
    """Main function to generate an outline with user preferences or provided arguments."""
    try:
        if not topic or not detail_level:
            user_topic_input, detail_level = get_user_preferences()
            topic = extract_topic_from_input(user_topic_input)
        
        client = initialize_groq_client()
        outline_text = get_presentation_outline(client, topic, detail_level)
        
        print(f"\nGenerated {detail_level} outline:\n")
        print(outline_text)
        
        # This line is correct, it expects two return values
        presentation_title, outline_dict = parse_llm_output_to_outline(outline_text)
        
        print(f"\nPresentation Title: {presentation_title}")
        print(f"Parsed Outline Structure ({len(outline_dict)} slides):")
        for section, points in outline_dict.items():
            print(f"• {section}: {len(points)} bullet points")
        
        slide_count = len(outline_dict)
        if detail_level == "simple" and not (7 <= slide_count <= 12):
            print(f"⚠️ Warning: Simple presentation has {slide_count} slides (expected 7-12)")
        elif detail_level == "detailed" and not (10 <= slide_count <= 15):
            print(f"⚠️ Warning: Detailed presentation has {slide_count} slides (expected 10-15)")
        
        return outline_dict, presentation_title
        
    except Exception as e:
        print(f"❌ Error: {e}")
        print("Using mock data instead...")
        
        if not topic or not detail_level:
            user_topic_input, detail_level = get_user_preferences()
            topic = extract_topic_from_input(user_topic_input)
        
        outline_text = get_mock_outline(topic, detail_level)
        
        # Ensure this call also returns two values
        presentation_title, outline_dict = parse_llm_output_to_outline(outline_text)
        
        print(f"\nMock {detail_level} outline:\n")
        print(outline_text)
        
        print(f"\nPresentation Title: {presentation_title}")
        print(f"Parsed Outline Structure ({len(outline_dict)} slides):")
        for section, points in outline_dict.items():
            print(f"• {section}: {len(points)} bullet points")
        
        return outline_dict, presentation_title


if __name__ == "__main__":
    # Test the LLM functionality
    outline, presentation_title = generate_outline(args.topic, args.detail)
    
    # Show the parsed outline
    print("\nFinal Parsed Outline:")
    pprint.pprint(outline)