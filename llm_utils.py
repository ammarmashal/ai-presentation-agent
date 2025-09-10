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
    Clean up user input to extract the main topic
    """
    if not user_input:
        return "Artificial Intelligence"
    
    # Remove polite phrases and questions
    patterns_to_remove = [
        r"please(\s+(can you|could you|i want))?",
        r"can you",
        r"could you",
        r"i want",
        r"i need",
        r"generate",
        r"create",
        r"make",
        r"about",
        r"on",
        r"a presentation",
        r"a ppt",
        r"powerpoint",
        r"slides",
        r"presentation"
    ]
    
    cleaned = user_input.lower()
    for pattern in patterns_to_remove:
        cleaned = re.sub(pattern, "", cleaned)
    
    # Remove extra spaces and punctuation
    cleaned = re.sub(r'[^\w\s]', '', cleaned)
    cleaned = cleaned.strip()
    
    # Capitalize first letter of each word
    cleaned = cleaned.title()
    
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
    Generate a presentation outline using Groq LLM with specific structure and page counts
    """
    # Enhanced system prompt for consistent formatting and structure
    system_prompt = (
        "You are an AI presentation assistant. Create professional presentation outlines with strict formatting:\n"
        "- Each slide header must be on its own line with **Header** format\n"
        "- Use bullet points starting with '- ' for content\n"
        "- For detailed presentations: include 2-4 detailed bullet points per slide with explanations\n"
        "- For simple presentations: include 3-5 concise bullet points per slide\n"
        "- Use exactly two spaces per indent level for sub-bullets\n"
        "- Do not use numbering, tables, or code blocks\n"
        "- Ensure the presentation has logical flow: Introduction -> Main Content -> Conclusion\n"
        "- Make the content professional and suitable for business/educational presentations"
    )
    
    if detail_level == "simple":
        prompt = f"""Generate a concise professional presentation about {topic} with 10-12 slides. Structure it as follows:

1. The presentation should have:
    - Slide 1: Title slide (just the main topic as title).
    - Slide 2: Introduction (title = "Introduction") with 3 to 5 full sentences introducing the topic.  
    - Then: For each slide, create **a main slide title** and **3-5 very short bullet points only** (keywords or short phrases, no explanations).  
2. Do not add numbering or "Slide X:" text.  
3. Return the output in plain text format with:
    - Title
    - Subtitle / Introduction
    - Bullets

Ensure the presentation is professional, well-structured, and suitable for business audiences."""
    
    else:  # detailed
        prompt = f"""Generate a comprehensive professional presentation about {topic} with 12-15 slides. Structure it as follows:

1. The presentation should have:
   - Slide 1: Title slide (just the main topic as title).
   - Slide 2: Introduction (title = "Introduction") with 3 to 5 full sentences introducing the topic.  
   - Then: For each slide, create a main slide title, and under it:
       - 2 to 3 bullet points
       - Each bullet point should have 2–3 full sentences explaining it with muximam kines = 3.  
2. Do not add numbering or "Slide X:" text.  
3. Return the output in plain text format with:
   - Title
   - Subtitle / Introduction
   - Bullets with explanation

Ensure each bullet point provides substantial detail and explanation. Make it professional and suitable for expert audiences."""
    
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

def parse_llm_output_to_outline(llm_output: str):
    """
    Parse LLM output to a structured outline and extract the main title
    Returns: (title, outline_dict)
    """
    outline = {}
    current_section = None
    default_section = "Content"
    main_title = None
    
    lines = llm_output.splitlines()
    
    for i, raw_line in enumerate(lines):
        if not raw_line.strip():
            continue

        stripped = raw_line.strip()

        # Extract main title from the first header found
        if main_title is None:
            header_match = re.match(r'^\*\*(.+?)\*\*$', stripped)
            if header_match:
                main_title = header_match.group(1).strip()
                # Skip adding this to outline as it's the main title
                continue

        # Standalone section headers (after the main title)
        header_match = re.match(r'^\*\*(.+?)\*\*$', stripped)
        if header_match:
            current_section = header_match.group(1).strip()
            outline[current_section] = []
            continue

        slide_match = re.match(r'^Slide\s+\d+:\s*(.+)$', stripped, flags=re.IGNORECASE)
        if slide_match:
            current_section = slide_match.group(1).strip()
            outline[current_section] = []
            continue

        # Bullets (keep leading whitespace to compute level)
        bullet_match = re.match(r'^[ \t]*[-*•]\s+(.*)$', raw_line)
        if bullet_match:
            # Compute level from leading whitespace
            leading_ws_len = len(raw_line) - len(raw_line.lstrip(' \t'))
            ws_prefix = raw_line[:leading_ws_len]
            tabs = ws_prefix.count('\t')
            spaces = ws_prefix.count(' ')
            level = tabs + (spaces // 2)  # 2 spaces = one level

            text = bullet_match.group(1).strip()
            # Remove surrounding bold if present on the bullet text itself
            text = re.sub(r'^\*\*(.+?)\*\*$', r'\1', text)

            if not current_section:
                current_section = default_section
                outline.setdefault(current_section, [])

            outline[current_section].append({
                "text": text,
                "level": max(0, min(level, 8))  # PowerPoint levels 0..8
            })
            continue

        # Non-bullet content: treat as section if none exists yet, else as level-0 point
        if not current_section:
            current_section = stripped
            outline[current_section] = []
        else:
            outline[current_section].append({"text": stripped, "level": 0})

    # If no main title was found in headers, use the first line or a default
    if main_title is None and lines:
        first_line = lines[0].strip()
        if first_line:
            main_title = first_line
            # Remove any markdown formatting from the title
            main_title = re.sub(r'^\*\*(.+?)\*\*$', r'\1', main_title)
    
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