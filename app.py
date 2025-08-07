import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches
from PIL import Image
import json
import re
import io
import zipfile
from datetime import datetime
import base64
from clipboard_component import copy_component, paste_component

# Configure the page
st.set_page_config(
    page_title="PowerPoint AI Field Filler",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for better styling
# Custom CSS for better styling
st.markdown("""
<style>
    /* Center the main content and set a maximum width */
    .main .block-container {
        max-width: 1080px;
        margin: 0 auto;
    }

    .main-header {
        text-align: center;
        color: #2c3e50;
        font-size: 2.5rem;
        margin-bottom: 2rem;
        padding: 1rem;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border-radius: 10px;
    }
    .step-container {
        background: #f8f9fa;
        padding: 20px;
        border-radius: 10px;
        margin: 20px 0;
        border-left: 5px solid #007bff;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 15px;
        margin: 10px 0;
    }
    .warning-box {
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        border-radius: 5px;
        padding: 15px;
        margin: 10px 0;
    }
    .ai-button {
        margin: 5px;
        padding: 10px 20px;
        border-radius: 5px;
        border: none;
        color: white;
        font-weight: bold;
        text-decoration: none;
        display: inline-block;
        text-align: center;
    }
    .nipr-btn { background-color: #004d40; } /* Dark Teal for NiprGPT */
</style>
""", unsafe_allow_html=True)

def analyze_powerpoint_fields(uploaded_file):
    """Analyze PowerPoint file for field placeholders"""
    try:
        prs = Presentation(uploaded_file)
        found_fields = set()
        field_locations = []
        
        for slide_num, slide in enumerate(prs.slides, 1):
            for shape in slide.shapes:
                if hasattr(shape, "text_frame") and shape.text_frame:
                    text_content = shape.text_frame.text
                    field_pattern = r'\{\{([^}]+)\}\}'
                    matches = re.findall(field_pattern, text_content)
                    
                    for field in matches:
                        found_fields.add(field)
                        field_locations.append({
                            'field': field,
                            'slide': slide_num,
                            'context': text_content[:100] + '...' if len(text_content) > 100 else text_content
                        })
                
                elif shape.has_table:
                    table = shape.table
                    for row_num, row in enumerate(table.rows):
                        for cell_num, cell in enumerate(row.cells):
                            text_content = cell.text
                            field_pattern = r'\{\{([^}]+)\}\}'
                            matches = re.findall(field_pattern, text_content)
                            
                            for field in matches:
                                found_fields.add(field)
                                field_locations.append({
                                    'field': field,
                                    'slide': slide_num,
                                    'location': f'Table R{row_num+1}C{cell_num+1}',
                                    'context': text_content[:50] + '...' if len(text_content) > 50 else text_content
                                })
        
        return list(found_fields), field_locations
    except Exception as e:
        st.error(f"Error analyzing PowerPoint: {str(e)}")
        return [], []

def generate_ai_prompt(fields, project_data):
    """Generate AI prompt"""
    field_descriptions = [f"  - {field}" for field in sorted(fields)]
    
    prompt = f"""I need you to analyze project data and extract information for specific PowerPoint fields. Return ONLY a valid JSON object with the field names as keys and extracted values as values.

**PowerPoint Fields to Fill:**
{chr(10).join(field_descriptions)}

**Instructions:**
1. Extract relevant information from the project data for each field
2. If a field name suggests specific content (e.g., "commander_name" should be a person's name), extract accordingly
3. Keep values concise but informative - suitable for presentation slides
4. Conduct market research with a focus on Department of Defense, Department of the Air Force, and with the goals of the 100th ARW and 352nd SOW mission goals in mind
5. For fields with money, phone numbers, or other implied formatting, format the extracted values accordingly
6. For fields you can't determine from the data, use "TBD" or leave reasonable placeholder text
7. Return ONLY the JSON object - no explanations or additional text

**Project Data to Analyze:**
{project_data}

Please analyze the above data and return the JSON object with field values:"""
    
    return prompt

def copy_run_formatting(source_run, target_run):
    """Copy formatting between runs with enhanced error handling"""
    try:
        # Basic font properties with safe access
        try:
            if hasattr(source_run.font, 'name') and source_run.font.name:
                target_run.font.name = source_run.font.name
        except: pass
        
        try:
            if hasattr(source_run.font, 'size') and source_run.font.size:
                target_run.font.size = source_run.font.size
        except: pass
        
        try:
            if hasattr(source_run.font, 'bold') and source_run.font.bold is not None:
                target_run.font.bold = source_run.font.bold
        except: pass
        
        try:
            if hasattr(source_run.font, 'italic') and source_run.font.italic is not None:
                target_run.font.italic = source_run.font.italic
        except: pass
        
        try:
            if hasattr(source_run.font, 'underline') and source_run.font.underline is not None:
                target_run.font.underline = source_run.font.underline
        except: pass
        
        # Enhanced color handling
        try:
            source_color = source_run.font.color
            target_color = target_run.font.color
            
            # Try RGB color first
            if hasattr(source_color, 'rgb') and source_color.rgb is not None:
                target_color.rgb = source_color.rgb
            # Try theme color
            elif hasattr(source_color, 'theme_color') and source_color.theme_color is not None:
                target_color.theme_color = source_color.theme_color
                if hasattr(source_color, 'brightness') and source_color.brightness is not None:
                    target_color.brightness = source_color.brightness
            # Try element copy as fallback
            else:
                try:
                    if hasattr(source_color, '_element'):
                        target_color._element = source_color._element
                except: pass
        except: pass
        
    except: pass

def replace_in_existing_runs(paragraph, placeholder, replacement_text):
    """Replace text in existing runs"""
    if placeholder not in paragraph.text:
        return False
    
    full_text = paragraph.text
    placeholder_start = full_text.find(placeholder)
    placeholder_end = placeholder_start + len(placeholder)
    
    current_pos = 0
    for run in paragraph.runs:
        run_start = current_pos
        run_end = current_pos + len(run.text)
        
        if run_start <= placeholder_start and placeholder_end <= run_end:
            placeholder_start_in_run = placeholder_start - run_start
            placeholder_end_in_run = placeholder_end - run_start
            
            original_run_text = run.text
            new_run_text = (
                original_run_text[:placeholder_start_in_run] +
                replacement_text +
                original_run_text[placeholder_end_in_run:]
            )
            
            run.text = new_run_text
            return True
        
        current_pos = run_end
    
    return False

def replace_text_preserve_formatting(paragraph, placeholder, replacement_text):
    """Replace text while preserving formatting"""
    if placeholder not in paragraph.text:
        return False

    # Find reference run for formatting
    placeholder_start = paragraph.text.find(placeholder)
    placeholder_end = placeholder_start + len(placeholder)
    
    current_pos = 0
    reference_run = None
    
    for run in paragraph.runs:
        run_start = current_pos
        run_end = current_pos + len(run.text)
        
        if run_start < placeholder_end and run_end > placeholder_start:
            if reference_run is None:
                reference_run = run
        
        current_pos = run_end

    if not reference_run:
        return False

    # Store paragraph formatting
    original_alignment = paragraph.alignment
    original_level = paragraph.level

    # Replace text
    new_text = paragraph.text.replace(placeholder, replacement_text)
    
    # Clear and rebuild paragraph
    paragraph.clear()
    
    try:
        paragraph.alignment = original_alignment
        paragraph.level = original_level
    except: pass
    
    # Add new run with formatting
    new_run = paragraph.add_run()
    new_run.text = new_text
    copy_run_formatting(reference_run, new_run)

    return True

def replace_image_by_alt_text(prs, image_data, progress_container):
    """Replace images based on alt text placeholders"""
    if not image_data:
        return 0
    
    replacements_made = 0
    
    # Image patterns for alt text
    image_patterns = [
        r'\{\{.*image.*\}\}',
        r'\{\{.*photo.*\}\}',
        r'\{\{.*picture.*\}\}',
        r'\{\{.*graphic.*\}\}',
        r'\{\{.*logo.*\}\}'
    ]
    
    for slide_num, slide in enumerate(prs.slides, 1):
        shapes_to_replace = []
        
        for shape in slide.shapes:
            if hasattr(shape, '_element') and shape._element.tag.endswith('}pic'):
                try:
                    # Get alt text
                    alt_text = ""
                    if hasattr(shape, 'element'):
                        nvPicPr = shape.element.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}nvPicPr')
                        if nvPicPr is not None:
                            cNvPr = nvPicPr.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}cNvPr')
                            if cNvPr is not None:
                                alt_text = cNvPr.get('descr', '') or cNvPr.get('title', '')
                    
                    if alt_text:
                        for pattern in image_patterns:
                            if re.search(pattern, alt_text, re.IGNORECASE):
                                shapes_to_replace.append({
                                    'shape': shape,
                                    'alt_text': alt_text,
                                    'slide_num': slide_num
                                })
                                break
                except: continue
        
        # Replace identified shapes
        for shape_info in shapes_to_replace:
            try:
                shape = shape_info['shape']
                left = shape.left
                top = shape.top
                width = shape.width
                height = shape.height
                
                # Remove old image
                shape_element = shape._element
                shape_element.getparent().remove(shape_element)
                
                # Add new image with uploaded data
                image_stream = io.BytesIO(image_data)
                new_picture = slide.shapes.add_picture(
                    image_stream, left, top, width, height
                )
                
                replacements_made += 1
                progress_container.success(f"🖼️ Replaced image on Slide {slide_num}")
                
            except Exception as e:
                progress_container.warning(f"⚠️ Image replacement failed on Slide {slide_num}: {str(e)}")
    
    return replacements_made

def fill_powerpoint_with_data(prs, json_data, uploaded_image, progress_container):
    """Fill PowerPoint with data preserving formatting"""
    try:
        if isinstance(json_data, str):
            json_str_match = re.search(r'\{.*\}', json_data, re.DOTALL)
            if json_str_match:
                json_str = json_str_match.group(0)
                data = json.loads(json_str)
            else:
                raise json.JSONDecodeError("No JSON object found", "", 0)
        else:
            data = json_data

        replacements_made = 0

        # Handle image replacement if image was uploaded
        if uploaded_image:
            progress_container.info("🖼️ Processing image replacement...")
            image_replacements = replace_image_by_alt_text(prs, uploaded_image.getvalue(), progress_container)
            if image_replacements == 0:
                progress_container.warning("⚠️ No images with placeholder alt text found")

        progress_container.info(f"📊 Processing {len(data)} text fields...")

        for slide_num, slide in enumerate(prs.slides, 1):
            for shape in slide.shapes:
                if hasattr(shape, "text_frame") and shape.text_frame:
                    for para in shape.text_frame.paragraphs:
                        for field, value in data.items():
                            placeholder = f"{{{{{field}}}}}"
                            if placeholder in para.text:
                                # Try in-place replacement first
                                success = replace_in_existing_runs(para, placeholder, str(value))
                                if not success:
                                    # Try formatting-preserving replacement
                                    success = replace_text_preserve_formatting(para, placeholder, str(value))
                                
                                if success:
                                    replacements_made += 1

                elif shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            for para in cell.text_frame.paragraphs:
                                for field, value in data.items():
                                    placeholder = f"{{{{{field}}}}}"
                                    if placeholder in para.text:
                                        success = replace_in_existing_runs(para, placeholder, str(value))
                                        if not success:
                                            success = replace_text_preserve_formatting(para, placeholder, str(value))
                                        if success:
                                            replacements_made += 1

        progress_container.success(f"✅ Made {replacements_made} text replacements")
        return prs, replacements_made

    except Exception as e:
        progress_container.error(f"Error filling PowerPoint: {e}")
        return None, 0

def create_download_link(data, filename):
    """Create a download link for the filled PowerPoint"""
    b64 = base64.b64encode(data).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{b64}" download="{filename}">📥 Download Filled PowerPoint</a>'
    return href

def main():
    # --- NEW: Banners at the very top ---
    st.warning('**DO NOT ENTER CONTROLLED UNCLASSIFIED INFORMATION INTO THIS SYSTEM**')
    
    # CORRECTED CODE
    try:
        st.image("banner.png", use_container_width=True)
    except Exception as e:
    # This will prevent the app from crashing if the banner.png is not found
        st.info("Info: `banner.png` not found. Skipping image banner.")

    # Header
    st.markdown('<div class="main-header">📊 PowerPoint AI Field Filler</div>', unsafe_allow_html=True)
    
    st.markdown("**Transform your PowerPoint templates with AI-powered data filling!**")

    if 'fields' not in st.session_state:
        st.session_state.fields = []
    if 'field_locations' not in st.session_state:
        st.session_state.field_locations = []
    if 'ai_prompt' not in st.session_state:
        st.session_state.ai_prompt = ""

    # --- MODIFIED: Step 1 ---
    st.markdown('<div class="step-container">', unsafe_allow_html=True)
    st.markdown("### 📁 Step 1: Choose Your PowerPoint Template")
    
    template_choice = st.radio(
        "Select a template source:",
        ('Use the One-Pager template', 'Upload your own template'),
        horizontal=True,
        label_visibility="collapsed"
    )

    uploaded_file = None
    
    if template_choice == 'Upload your own template':
        uploaded_file = st.file_uploader(
            "Choose your PowerPoint template with {{field_name}} placeholders",
            type=['pptx'],
            help="Upload a PowerPoint file containing placeholders like {{project_title}}, {{commander_name}}, etc."
        )
    else:
        try:
            with open("onepager_template.pptx", "rb") as f:
                uploaded_file = io.BytesIO(f.read())
            st.success("✅ Loaded the built-in 'One-Pager' template.")
        except FileNotFoundError:
            st.error("Error: `onepager_template.pptx` not found in the repository. Please ask the app administrator to upload it.")
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    if uploaded_file is not None:
        with st.spinner('🔍 Analyzing PowerPoint fields...'):
            # When using the local file, we give it a name attribute for consistency
            if not hasattr(uploaded_file, 'name'):
                 uploaded_file.name = "onepager_template.pptx"
            st.session_state.fields, st.session_state.field_locations = analyze_powerpoint_fields(uploaded_file)
        
        if st.session_state.fields:
            st.markdown('<div class="success-box">', unsafe_allow_html=True)
            st.success(f"Found {len(st.session_state.fields)} placeholders in '{uploaded_file.name}'!")
            
            with st.expander("Click to see found fields and their locations"):
                col1, col2 = st.columns(2)
                with col1:
                    st.write("**Found Fields:**")
                    for field in sorted(st.session_state.fields):
                        st.write(f"• `{{{{{field}}}}}`")
                
                with col2:
                    st.write("**Field Locations (first 5):**")
                    for loc in st.session_state.field_locations[:5]:
                        st.write(f"• `{{{{{loc['field']}}}}}` on Slide {loc['slide']}")
                    if len(st.session_state.field_locations) > 5:
                        st.write(f"... and {len(st.session_state.field_locations) - 5} more")
            
            st.markdown('</div>', unsafe_allow_html=True)
            
            # Step 2: Project Data Input
            st.markdown('<div class="step-container">', unsafe_allow_html=True)
            st.markdown("### 📝 Step 2: Enter Your Project Data")
            
            project_data = st.text_area(
                "Paste your raw project data here:",
                height=200,
                placeholder="Enter your project details, requirements, team information, etc. The AI will extract relevant information for each field."
            )
            
            # Image upload section
            st.markdown("**Product Image (Optional):**")
            uploaded_image = st.file_uploader(
                "Choose an image file",
                type=['png', 'jpg', 'jpeg', 'gif', 'bmp'],
                help="This will replace images in your PowerPoint that have alt text like {{project_image}}"
            )
            
            if uploaded_image:
                st.image(uploaded_image, caption="Uploaded Image Preview", width=200)
                st.info("💡 **Tip:** In PowerPoint, right-click any image → Edit Alt Text → Add description like `{{project_image}}` to replace it with this uploaded image")
            
            st.markdown('</div>', unsafe_allow_html=True)
            
            if project_data.strip():
                if st.button("🤖 Generate AI Prompt", type="primary"):
                    st.session_state.ai_prompt = generate_ai_prompt(st.session_state.fields, project_data)
                
                if st.session_state.ai_prompt:
                    # Step 3: AI Prompt
                    st.markdown('<div class="step-container">', unsafe_allow_html=True)
                    st.markdown("### 📋 Step 3: Copy Prompt to AI")
                    
                    st.info("Copy this prompt and paste it into your preferred AI assistant.")
                    
                    # Display the prompt in an expandable section
                    with st.expander("📄 Click to view the generated AI Prompt", expanded=True):
                        st.code(st.session_state.ai_prompt, language="text")
                    
                    # CORRECT LINE
                    copy_component("📋 Copy Prompt to Clipboard", st.session_state.ai_prompt)

                    # --- MODIFIED: AI Service Button ---
                    st.markdown("**Quick Link to AI Service:**")
                    st.markdown(
                        f'<a href="https://niprgpt.mil/" target="_blank" class="ai-button nipr-btn">🚀 Open NiprGPT</a>',
                        unsafe_allow_html=True
                    )
                    
                    st.markdown('</div>', unsafe_allow_html=True)
                    
                    # Step 4: AI Response
                    st.markdown('<div class="step-container">', unsafe_allow_html=True)
                    st.markdown("### 🔄 Step 4: Paste AI Response")
                    
                    ai_response = st.text_area(
                        "Paste the AI's JSON response here:",
                        height=150,
                        placeholder='{"project_title": "Your Project", "commander_name": "John Doe", ...}'
                    )
                    
                    if ai_response.strip():
                        try:
                            # Validate JSON
                            json_str_match = re.search(r'\{.*\}', ai_response, re.DOTALL)
                            if not json_str_match:
                                raise json.JSONDecodeError("No JSON object found in the response.", ai_response, 0)
                            
                            json_data = json.loads(json_str_match.group(0))
                            
                            st.success("✅ Valid JSON detected!")
                            
                            # Show preview of data
                            with st.expander("👀 Preview Data"):
                                st.json(json_data)
                            
                            if st.button("🚀 Generate Filled PowerPoint", type="primary"):
                                progress_container = st.container()
                                
                                with st.spinner('🔄 Filling PowerPoint template...'):
                                    # Reset file pointer
                                    uploaded_file.seek(0)
                                    prs = Presentation(uploaded_file)
                                    
                                    # Fill the PowerPoint
                                    filled_prs, replacements = fill_powerpoint_with_data(
                                        prs, json_data, uploaded_image, progress_container
                                    )
                                    
                                    if filled_prs:
                                        # Save to bytes
                                        output_buffer = io.BytesIO()
                                        filled_prs.save(output_buffer)
                                        output_buffer.seek(0)
                                        
                                        # Generate filename
                                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                        filename = f"filled_presentation_{timestamp}.pptx"
                                        
                                        st.success(f"✅ PowerPoint generated successfully! ({replacements} replacements made)")
                                        
                                        # Download button
                                        st.download_button(
                                            label="📥 Download Filled PowerPoint",
                                            data=output_buffer.getvalue(),
                                            file_name=filename,
                                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                                        )
                                        
                                        st.balloons()  # Celebration animation!
                        
                        except json.JSONDecodeError as e:
                            st.error(f"❌ Invalid JSON format in the AI response: {str(e)}")
                            st.info("💡 Please ensure you paste the entire, unmodified JSON object from the AI.")
                    
                    st.markdown('</div>', unsafe_allow_html=True)
        
        elif uploaded_file is not None: # This check prevents the message from showing before analysis
            st.markdown('<div class="warning-box">', unsafe_allow_html=True)
            st.warning("⚠️ No {{field_name}} placeholders found in your PowerPoint!")
            st.write("Make sure your PowerPoint contains placeholders like:")
            st.code("{{project_title}}, {{commander_name}}, {{problem_description}}")
            st.markdown('</div>', unsafe_allow_html=True)

    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; padding: 20px;">
        <p>🚀 Built for NIPR environments • No local installation required • Works in any browser</p>
        <p>💡 <strong>How to use:</strong> Upload PowerPoint → Add project data → Get AI response → Download filled presentation</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()






