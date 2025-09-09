import os
import argparse
from pptx import Presentation
from llm_utils import generate_outline

def list_available_themes(theme_folder):
    """List all available themes in the specified folder"""
    themes = [f for f in os.listdir(theme_folder) if f.endswith(".pptx")]
    if not themes:
        raise FileNotFoundError(f"❌ No .pptx themes found in folder: {theme_folder}")
    return themes

def create_presentation(outline, topic, theme_path, filename="presentation.pptx"):
    """Create a PowerPoint presentation using a .pptx theme file"""
    # Validate theme file
    if not theme_path.endswith(".pptx"):
        raise ValueError(f"❌ The theme file must be a .pptx presentation: {theme_path}")
    
    # Load the theme file
    if os.path.exists(theme_path):
        prs = Presentation(theme_path)
        print(f"🎨 Theme loaded: {theme_path}")
    else:
        raise FileNotFoundError(f"❌ Theme file not found: {theme_path}")
    
    # Title slide
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title = title_slide.shapes.title
    subtitle = title_slide.placeholders[1] if len(title_slide.placeholders) > 1 else None
    title.text = topic
    if subtitle:
        subtitle.text = "Created with AI Presentation Generator"
    
    # Content slides
    for section, points in outline.items():
        if not section.strip():
            continue
            
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        title_shape = slide.shapes.title
        content_shape = slide.placeholders[1]
        
        title_shape.text = section
        
        if points and len(points) > 0:
            content_text = "\n".join(f"• {point}" for point in points)
            content_shape.text = content_text
    
    # Delete the first slide (if needed)
    xml_slides = prs.slides._sldIdLst  # Access the slide ID list
    first_slide_id = xml_slides[0]  # Get the first slide ID
    xml_slides.remove(first_slide_id)  # Remove the first slide
    
    # Save the presentation
    prs.save(filename)
    print(f"✅ Presentation saved as: {filename}")
    return filename


def main():
    parser = argparse.ArgumentParser(description="Generate AI-powered presentations")
    parser.add_argument("--topic", type=str, default="Artificial Intelligence", help="Presentation topic")
    parser.add_argument("--detail", type=str, default="simple", choices=["simple", "detailed"], help="Detail level")
    parser.add_argument("--theme", type=str, help="Name of the theme file (without extension)")
    parser.add_argument("--output", type=str, default="presentation.pptx", help="Output filename")
    
    args = parser.parse_args()
    
    theme_folder = "themes"
    available_themes = list_available_themes(theme_folder)
    
    print("🎯 AI Presentation Generator")
    print("=" * 40)
    print(f"Topic: {args.topic}")
    print(f"Detail Level: {args.detail}")
    print("Available Themes:")
    for theme in available_themes:
        print(f" - {os.path.splitext(theme)[0]}")
    print("=" * 40)
    
    if not args.theme:
        raise ValueError("❌ Please specify a theme using the --theme argument.")
    
    theme_path = os.path.join(theme_folder, f"{args.theme}.pptx")
    if not os.path.exists(theme_path):
        raise FileNotFoundError(f"❌ Theme '{args.theme}' not found in folder: {theme_folder}")
    
    # Generate outline
    outline, topic = generate_outline(args.topic, args.detail)
    
    # Create presentation
    create_presentation(outline, topic, theme_path, args.output)
    
    print("🎉 Presentation created successfully!")

if __name__ == "__main__":
    main()