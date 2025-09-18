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
        Generate a concise, professional presentation outline about **{topic}**. The presentation should have 10-12 slides. Adhere to the following strict formatting rules:
        
        1. The first slide is "**Introduction**". It must have 2-3 full sentences of introductory text, not bullet points.
        2. Each subsequent slide title must be on its own line surrounded by double asterisks: **Slide Title**
        3. All other slides must have exactly 5-7 very short bullet points
        4. Each bullet point should start with • and be a concise keyword/phrase
        5. The final slide must be "**Conclusion**"

        EXAMPLE:
        **Introduction**
        This is the introductory text about the topic.

        **Core Concepts**
        • Concept 1
        • Concept 2  
        • Concept 3
        • Concept 4
        • Concept 5

        **Applications**
        • App 1
        • App 2
        • App 3
        • App 4

        **Conclusion**
        • Summary point 

        Now generate the outline for: {topic}
        """
    
    else:  # detailed
        # In get_presentation_outline function, update the prompt:
        prompt = f"""
        Generate a comprehensive, professional presentation outline about **{topic}**. 

        STRICT FORMATTING RULES:
        1. The first slide must be "**Introduction**" with 3-5 full sentences (no bullet points)
        2. Each subsequent slide title must be on its own line surrounded by double asterisks: **Slide Title**
        3. Under each slide title, provide exactly 2-3 detailed bullet points
        4. Each bullet point must start with • and be a meaningful sentence (15-30 words)
        5. The final slide must be "**Conclusion**"

        EXAMPLE:
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
        • Summary of key takeaways

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
    Parse LLM output with strict slide boundary detection based on **Title** markers
    """
    outline = {}
    current_slide = None
    slide_content = []
    main_title = None
    
    # Pattern to detect slide titles: both **Title** and # Title formats
    slide_pattern = r'^\s*(?:\*\*([^*]+)\*\*|#\s+([^#]+))$'
    
    lines = llm_output.splitlines()
    
    # First, try to find the actual presentation title by skipping LLM's intro phrases
    intro_phrases = [
        'here is', 'this is', 'below is', 'following is',
        'presentation outline', 'outline for', 'generated outline',
        'i have', 'i\'ve created', 'the following'
    ]
    
    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue
        
        line_lower = line.lower()
        
        # Skip LLM's introductory phrases
        if any(phrase in line_lower for phrase in intro_phrases):
            continue
            
        # Look for the first meaningful line that could be a title
        if (len(line) > 3 and len(line) < 80 and  # Reasonable title length
            not line.startswith(('#', '*', '-', '•', '1.', '2.', '3.')) and  # Not a bullet, header, or number
            not line.endswith((':', '-')) and  # Not a label or separator
            not re.match(r'^\d+\.', line) and  # Not a numbered list
            not re.match(r'^[•\-*]', line)):  # Not a bullet point
            
            # Clean the line and use as main title
            main_title = clean_markdown(line)
            break
    
    # Now parse the slides
    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue
        
        # Check if this is a slide title (either **Title** or # Title)
        title_match = re.match(slide_pattern, line)
        
        if title_match:
            # Extract title from either capture group
            slide_title = title_match.group(1) or title_match.group(2)
            slide_title = slide_title.strip()
            
            # Save previous slide content if exists
            if current_slide and slide_content:
                outline[current_slide] = process_bullet_points(slide_content)
            
            # Start new slide
            current_slide = slide_title
            slide_content = []
            
        else:
            # Add content to current slide
            if current_slide is not None:
                slide_content.append(line)
    
    # Add the final slide
    if current_slide and slide_content:
        outline[current_slide] = process_bullet_points(slide_content)
    
    # If no main title found, use the first slide title or user topic
    if not main_title:
        if outline:
            # Use the first non-introduction slide title
            for slide_title in outline.keys():
                if "introduction" not in slide_title.lower():
                    main_title = slide_title
                    break
            # If all slides are introduction-like, use the first one
            if not main_title:
                main_title = next(iter(outline.keys()))
        else:
            main_title = "Presentation"
    
    # Ensure Introduction slide exists and is first
    if "Introduction" not in outline and any("introduction" in key.lower() for key in outline.keys()):
        # Rename existing introduction-like slide
        for key in list(outline.keys()):
            if "introduction" in key.lower():
                outline["Introduction"] = outline.pop(key)
                break
    
    # Reorder to ensure Introduction is first if it exists
    if "Introduction" in outline:
        ordered_outline = {"Introduction": outline["Introduction"]}
        for key, value in outline.items():
            if key != "Introduction":
                ordered_outline[key] = value
        outline = ordered_outline
    
    return main_title, outline



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
        # Use provided arguments or get user preferences
        if not topic or not detail_level:
            user_topic_input, detail_level = get_user_preferences()
            # Extract clean topic from user input for the prompt
            topic = extract_topic_from_input(user_topic_input)
        
        # Initialize client and generate outline
        client = initialize_groq_client()
        outline_text = get_presentation_outline(client, topic, detail_level)
        
        print(f"\nGenerated {detail_level} outline:\n")
        print(outline_text)
        
        # Parse to dictionary and extract title
        presentation_title, outline_dict = parse_llm_output_to_outline(outline_text)
        
        # If we couldn't extract a title, use the cleaned topic
        if not presentation_title:
            presentation_title = topic
        
        print(f"\nPresentation Title: {presentation_title}")
        print(f"Parsed Outline Structure ({len(outline_dict)} slides):")
        for section, points in outline_dict.items():
            print(f"• {section}: {len(points)} bullet points")
        
        # Validate slide count
        slide_count = len(outline_dict)
        if detail_level == "simple" and not (7 <= slide_count <= 12):
            print(f"⚠️  Warning: Simple presentation has {slide_count} slides (expected 7-12)")
        elif detail_level == "detailed" and not (10 <= slide_count <= 15):
            print(f"⚠️  Warning: Detailed presentation has {slide_count} slides (expected 10-15)")
        
        return outline_dict, presentation_title
        
    except Exception as e:
        print(f"❌ Error: {e}")
        print("Using mock data instead...")
        
        # Use mock data if API fails
        if not topic or not detail_level:
            user_topic_input, detail_level = get_user_preferences()
            topic = extract_topic_from_input(user_topic_input)
        
        outline_text = get_mock_outline(topic, detail_level)
        presentation_title, outline_dict = parse_llm_output_to_outline(outline_text)
        
        # If we couldn't extract a title, use the cleaned topic
        if not presentation_title:
            presentation_title = topic
        
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