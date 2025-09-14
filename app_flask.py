from flask import Flask, render_template, request, send_file, session, redirect, url_for
import os
import tempfile
from llm_utils import generate_outline
from ppt_generator import create_presentation, list_available_themes
from flask import send_from_directory


app = Flask(__name__)
app.secret_key = 'your-secret-key-here'  


def find_matching_image(theme_name):
    """
    Try to find a matching image for the theme
    """
    image_folder = "static/images/themes"
    if not os.path.exists(image_folder):
        return None
    
    # Try exact match first
    possible_extensions = ['.png', '.jpg', '.jpeg', '.webp']
    for ext in possible_extensions:
        image_path = f"{image_folder}/{theme_name}{ext}"
        if os.path.exists(image_path):
            return f"/{image_path}"
    
    # Try partial matches
    images = os.listdir(image_folder)
    theme_lower = theme_name.lower()
    
    for img in images:
        if img.lower().endswith(('.png', '.jpg', '.jpeg', '.webp')):
            img_name = os.path.splitext(img)[0].lower()
            # Check if theme contains image name or vice versa
            if img_name in theme_lower or theme_lower in img_name:
                return f"/static/images/themes/{img}"
    
    return None

def get_theme_image_url(theme_name, index=None):
    """
    Get image URL for theme with optional index
    """
    # Try to find matching image
    image_url = find_matching_image(theme_name)
    if image_url:
        return image_url
    
    # If index is provided, try that specific numbered image
    if index is not None:
        image_files = [f"theme{index+1}.png", f"theme{index+1}.jpg"]
        for image_file in image_files:
            image_path = f"static/images/themes/{image_file}"
            if os.path.exists(image_path):
                return f"/{image_path}"
    
    # Try all numbered images as fallback
    for i in range(1, 20):
        image_files = [f"theme{i}.png", f"theme{i}.jpg"]
        for image_file in image_files:
            image_path = f"static/images/themes/{image_file}"
            if os.path.exists(image_path):
                return f"/{image_path}"
    
    # Final fallback to placeholder
    return f"https://placehold.co/400x250/2563eb/white?text={theme_name.replace(' ', '+')}&font=montserrat"




@app.route('/')
def index():
    return redirect(url_for('step1'))

@app.route('/step1', methods=['GET', 'POST'])
def step1():
    if request.method == 'POST':
        topic = request.form.get('topic', '').strip()
        detail_level = request.form.get('detail_level', 'simple')
        
        if not topic:
            return render_template('step1.html', error="Please enter a presentation topic")
        
        session['topic'] = topic
        session['detail_level'] = detail_level
        return redirect(url_for('step2'))
    
    return render_template('step1.html')


@app.route('/step2', methods=['GET', 'POST'])
def step2():
    print("📄 Step 2 accessed")
    if 'topic' not in session:
        return redirect(url_for('step1'))
    
    theme_folder = "themes"
    try:
        themes = list_available_themes(theme_folder)
        theme_names = [os.path.splitext(theme)[0] for theme in themes]
        
        print(f"🎨 Available themes: {theme_names}")
        
        if not themes:
            return render_template('step2.html', error="No themes found in the themes folder")
        
        # Create theme data with image URLs
        theme_data = []
        for theme_name in theme_names:
            theme_data.append({
                'name': theme_name,
                'image_url': get_theme_image_url(theme_name)
            })
        
        if request.method == 'POST':
            theme = request.form.get('theme')
            if theme:
                session['theme'] = theme
                print(f"🎨 Theme selected: {theme}")
                return redirect(url_for('step3'))
        
        return render_template('step2.html', themes=theme_data)
    
    except Exception as e:
        print(f"❌ Error in step2: {e}")
        return render_template('step2.html', error=f"Error loading themes: {str(e)}")
    
    except Exception as e:
        return render_template('step2.html', error=f"Error loading themes: {str(e)}")




@app.route('/static/<path:filename>')
def serve_static(filename):
    return send_from_directory('static', filename)




@app.route('/step3', methods=['GET', 'POST'])
def step3():
    if 'topic' not in session or 'theme' not in session:
        return redirect(url_for('step1'))
    
    if request.method == 'POST':
        try:
            # Generate outline
            outline, presentation_title = generate_outline(
                session['topic'], 
                session.get('detail_level', 'simple')
            )
            
            # Create presentation
            theme_path = os.path.join("themes", f"{session['theme']}.pptx")
            
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp_file:
                presentation_path = create_presentation(
                    outline, 
                    presentation_title, 
                    theme_path, 
                    session.get('detail_level', 'simple'),  # Pass detail level
                    tmp_file.name
                )
            
            session['presentation_path'] = presentation_path
            session['presentation_title'] = presentation_title
            
            return render_template('step3.html', 
                                    topic=session['topic'],
                                    detail_level=session.get('detail_level', 'simple'),
                                    theme=session['theme'],
                                    outline=outline,
                                    presentation_title=presentation_title,
                                    presentation_ready=True)
            
        except Exception as e:
            return render_template('step3.html', 
                                    topic=session['topic'],
                                    detail_level=session.get('detail_level', 'simple'),
                                    theme=session['theme'],
                                    error=f"Error generating presentation: {str(e)}")
    
    return render_template('step3.html', 
                            topic=session['topic'],
                            detail_level=session.get('detail_level', 'simple'),
                            theme=session['theme'])



@app.route('/download')
def download():
    if 'presentation_path' not in session or not os.path.exists(session['presentation_path']):
        return "Presentation not found", 404
    
    presentation_title = session.get('presentation_title', 'presentation')
    return send_file(
        session['presentation_path'],
        as_attachment=True,
        download_name=f"{presentation_title.replace(' ', '_')}.pptx",
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )




@app.route('/reset')
def reset():
    session.clear()
    return redirect(url_for('step1'))

if __name__ == '__main__':
    # Create templates directory if it doesn't exist
    os.makedirs('templates', exist_ok=True)
    app.run(debug=True, port=5000)