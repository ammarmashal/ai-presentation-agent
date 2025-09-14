import streamlit as st
import os
import tempfile
import sys


# Add the current directory to Python path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

try:
    from llm_utils import generate_outline
    from ppt_generator import create_presentation, list_available_themes
except ImportError as e:
    st.error(f"Import error: {e}. Please make sure all required modules are installed.")
    st.stop()

# Set page configuration
st.set_page_config(
    page_title="AI Presentation Generator",
    page_icon="📊",
    layout="wide"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.5rem;
        color: #2c3e50;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'step' not in st.session_state:
    st.session_state.step = 1
if 'topic' not in st.session_state:
    st.session_state.topic = ''
if 'detail_level' not in st.session_state:
    st.session_state.detail_level = 'simple'
if 'theme' not in st.session_state:
    st.session_state.theme = ''
if 'outline' not in st.session_state:
    st.session_state.outline = None
if 'presentation_title' not in st.session_state:
    st.session_state.presentation_title = ''
if 'presentation_path' not in st.session_state:
    st.session_state.presentation_path = ''

# Step 1: Topic Input
def step1():
    st.markdown('<h1 class="main-header">AI Presentation Generator</h1>', unsafe_allow_html=True)
    st.markdown('<h2 class="sub-header">Step 1: Enter Presentation Details</h2>', unsafe_allow_html=True)
    
    topic = st.text_input(
        "Presentation Topic",
        value=st.session_state.topic,
        placeholder="Enter your presentation topic (e.g., Artificial Intelligence)"
    )
    
    detail_level = st.radio(
        "Detail Level",
        options=["simple", "detailed"],
        index=0 if st.session_state.detail_level == "simple" else 1
    )
    
    if st.button("Next →"):
        if topic.strip():
            st.session_state.topic = topic.strip()
            st.session_state.detail_level = detail_level
            st.session_state.step = 2
            st.experimental_rerun()
        else:
            st.error("Please enter a presentation topic.")

# Step 2: Theme Selection
def step2():
    st.markdown('<h1 class="main-header">AI Presentation Generator</h1>', unsafe_allow_html=True)
    st.markdown('<h2 class="sub-header">Step 2: Choose a Theme</h2>', unsafe_allow_html=True)
    
    # Get available themes
    theme_folder = "themes"
    try:
        themes = list_available_themes(theme_folder)
        theme_names = [os.path.splitext(theme)[0] for theme in themes]
        
        if not themes:
            st.error("No themes found in the themes folder. Please add some .pptx theme files.")
            if st.button("← Back"):
                st.session_state.step = 1
                st.experimental_rerun()
            return
        
        # Display theme selection
        selected_theme = st.selectbox(
            "Select a theme:",
            options=theme_names,
            index=theme_names.index(st.session_state.theme) if st.session_state.theme in theme_names else 0
        )
        
        st.session_state.theme = selected_theme
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("← Back"):
                st.session_state.step = 1
                st.experimental_rerun()
        with col2:
            if st.button("Next →"):
                st.session_state.step = 3
                st.experimental_rerun()
                
    except Exception as e:
        st.error(f"Error loading themes: {str(e)}")
        if st.button("← Back"):
            st.session_state.step = 1
            st.experimental_rerun()

# Step 3: Presentation Generation
def step3():
    st.markdown('<h1 class="main-header">AI Presentation Generator</h1>', unsafe_allow_html=True)
    st.markdown('<h2 class="sub-header">Step 3: Generate and Download Presentation</h2>', unsafe_allow_html=True)
    
    # Show selected options
    st.info(f"**Topic:** {st.session_state.topic}")
    st.info(f"**Detail Level:** {st.session_state.detail_level}")
    st.info(f"**Theme:** {st.session_state.theme}")
    
    if st.button("Generate Presentation"):
        with st.spinner("Generating presentation outline..."):
            try:
                # Generate outline
                outline, presentation_title = generate_outline(
                    st.session_state.topic, 
                    st.session_state.detail_level
                )
                
                st.session_state.outline = outline
                st.session_state.presentation_title = presentation_title
                
                # Create presentation
                theme_path = os.path.join("themes", f"{st.session_state.theme}.pptx")
                
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp_file:
                    presentation_path = create_presentation(
                        outline, 
                        presentation_title, 
                        theme_path, 
                        tmp_file.name
                    )
                    st.session_state.presentation_path = presentation_path
                
                st.success("Presentation generated successfully!")
                
            except Exception as e:
                st.error(f"Error generating presentation: {str(e)}")
    
    # Display outline if available
    if st.session_state.outline:
        st.subheader("Presentation Outline")
        for section, points in st.session_state.outline.items():
            with st.beta_expander(section):
                for point in points:
                    st.write(f"- {point['text']}")
    
    # Download button if presentation is generated
    if st.session_state.presentation_path and os.path.exists(st.session_state.presentation_path):
        st.subheader("Download Your Presentation")
        
        with open(st.session_state.presentation_path, "rb") as file:
            btn = st.download_button(
                label="Download PowerPoint",
                data=file,
                file_name=f"{st.session_state.presentation_title.replace(' ', '_')}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
    
    # Navigation buttons
    col1, col2 = st.columns(2)
    with col1:
        if st.button("← Back to Theme Selection"):
            st.session_state.step = 2
            st.experimental_rerun()
    with col2:
        if st.button("Create New Presentation"):
            # Reset session state
            for key in list(st.session_state.keys()):
                if key != 'step':
                    st.session_state[key] = ''
            st.session_state.step = 1
            st.session_state.detail_level = 'simple'
            st.experimental_rerun()

# Main app
def main():
    if st.session_state.step == 1:
        step1()
    elif st.session_state.step == 2:
        step2()
    elif st.session_state.step == 3:
        step3()

if __name__ == "__main__":
    main()