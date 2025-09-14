import os
import argparse
from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE
from llm_utils import generate_outline
from image_search import add_images_to_presentation  




def list_available_themes(theme_folder):
    """List all available themes in the specified folder"""
    themes = [f for f in os.listdir(theme_folder) if f.endswith(".pptx")]
    if not themes:
        raise FileNotFoundError(f"❌ No .pptx themes found in folder: {theme_folder}")
    return themes


def create_presentation(outline, presentation_title, theme_path, detail_level, filename="presentation.pptx"):
    """Create a PowerPoint presentation using a .pptx theme file with proper nested bullets"""
    try:
        # Validate theme file
        if not theme_path.endswith(".pptx"):
            raise ValueError(f"❌ The theme file must be a .pptx presentation: {theme_path}")
        
        # Load the theme file
        if os.path.exists(theme_path):
            prs = Presentation(theme_path)
            print(f"🎨 Theme loaded: {theme_path}")
        else:
            raise FileNotFoundError(f"❌ Theme file not found: {theme_path}")
        
        # Title slide - USE THE LLM-GENERATED TITLE
        title_slide = prs.slides.add_slide(prs.slide_layouts[0])
        title = title_slide.shapes.title
        subtitle = title_slide.placeholders[1] if len(title_slide.placeholders) > 1 else None
        title.text = presentation_title  # Use the LLM-generated title here
        if subtitle:
            subtitle.text = "Created with AI Presentation Generator"
        
        # Content slides
        slide_index_map = {}  # Track the actual slide index for each outline section
        current_slide_idx = 1  # Start after title slide
        
        for section, points in outline.items():
            if not section.strip() or not points:
                continue
                
            # Create a new slide
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide_index_map[section] = current_slide_idx
            current_slide_idx += 1
            
            title_shape = slide.shapes.title
            content_shape = slide.placeholders[1]
            
            # Set the section title
            title_shape.text = section
            
            # Clear any default text and set up text frame
            text_frame = content_shape.text_frame
            text_frame.clear()
            text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            text_frame.word_wrap = True

            # Add bullet points with proper nesting
            first = True
            for point in points:
                level = max(0, min(point.get("level", 0), 8))

                if first:
                    p = text_frame.paragraphs[0]
                    first = False
                else:
                    p = text_frame.add_paragraph()

                p.text = point["text"]
                p.level = level
        
        # Add relevant images to the presentation (passing detail_level)
        prs = add_images_to_presentation(prs, outline, presentation_title, detail_level, num_images=3)
        
        # Delete the first slide (if needed)
        try:
            xml_slides = prs.slides._sldIdLst
            if len(xml_slides) > 0:
                first_slide_id = xml_slides[0]
                xml_slides.remove(first_slide_id)
        except:
            # If slide deletion fails, continue
            pass
        
        # Save the presentation
        prs.save(filename)
        print(f"✅ Presentation saved as: {filename}")
        return filename
        
    except Exception as e:
        print(f"❌ Error creating presentation with theme: {e}")
        print("🔄 Creating a basic presentation instead...")
        
        # Fallback: create a simple presentation without theme
        prs = Presentation()
        
        # Title slide
        title_slide = prs.slides.add_slide(prs.slide_layouts[0])
        title = title_slide.shapes.title
        subtitle = title_slide.placeholders[1] if len(title_slide.placeholders) > 1 else None
        title.text = presentation_title  # Use the LLM-generated title here
        if subtitle:
            subtitle.text = "Created with AI Presentation Generator"
        
        # Content slides
        for section, points in outline.items():
            if not section.strip() or not points:
                continue
                
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            title_shape = slide.shapes.title
            content_shape = slide.placeholders[1]
            title_shape.text = section
            
            text_frame = content_shape.text_frame
            text_frame.clear()
            text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            text_frame.word_wrap = True
            
            first = True
            for point in points:
                level = point.get("level", 0)
                
                if first:
                    p = text_frame.paragraphs[0]
                    first = False
                else:
                    p = text_frame.add_paragraph()
                
                p.text = point["text"]
                p.level = level
        
        # Add relevant images to the presentation (passing detail_level)
        prs = add_images_to_presentation(prs, outline, presentation_title, detail_level, num_images=3)
        
        # Save the fallback presentation
        fallback_filename = f"basic_{filename}"
        prs.save(fallback_filename)
        print(f"✅ Created basic presentation: {fallback_filename}")
        return fallback_filename

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
    outline, presentation_title = generate_outline(args.topic, args.detail)
    
    # Create presentation - pass the LLM-generated title AND detail level
    create_presentation(outline, presentation_title, theme_path, args.detail, args.output)

    print("🎉 Presentation created successfully!")

if __name__ == "__main__":
    main()