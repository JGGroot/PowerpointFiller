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
import docx
import glob
import os

# Configure the page
st.set_page_config(
    page_title="Document AI Field Filler",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Custom CSS for better styling
st.markdown("""
<style>
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
    .nipr-btn { background-color: #004d40; }
    .copy-button {
        background-color: #007bff;
        color: white;
        border: none;
        padding: 10px 20px;
        border-radius: 5px;
        cursor: pointer;
        font-weight: bold;
    }
    .copy-button:hover {
        background-color: #0056b3;
    }
</style>
""", unsafe_allow_html=True)

# --- Clipboard Functions (Replacement) ---
def copy_component(button_text, text_to_copy):
    """Simple clipboard functionality replacement"""
    st.markdown(f"""
    <button class="copy-button" onclick="navigator.clipboard.writeText(`{text_to_copy.replace('`', '\\`')}`).then(function() {{
        alert('Copied to clipboard!');
    }}).catch(function(err) {{
        console.error('Could not copy text: ', err);
        alert('Copy failed. Please copy manually.');
    }})">
    {button_text}
    </button>
    """, unsafe_allow_html=True)
    
    # Fallback - show the text in a code block for manual copying
    with st.expander("📋 Manual Copy (if button doesn't work)"):
        st.code(text_to_copy, language="text")

# --- Analysis Functions ---
def analyze_powerpoint_fields(uploaded_file):
    """Analyze PowerPoint file for field placeholders with improved error handling"""
    try:
        # Handle both file objects and file paths
        if isinstance(uploaded_file, str):
            prs = Presentation(uploaded_file)
        else:
            # Reset file pointer if it's a file object
            if hasattr(uploaded_file, 'seek'):
                uploaded_file.seek(0)
            prs = Presentation(uploaded_file)
        
        found_fields = set()
        field_locations = []
        field_pattern = r'\{\{([^}]+)\}\}'
        
        for slide_num, slide in enumerate(prs.slides, 1):
            for shape in slide.shapes:
                try:
                    if hasattr(shape, "text_frame") and shape.text_frame:
                        text_content = shape.text_frame.text
                        if text_content:
                            matches = re.findall(field_pattern, text_content)
                            for field in matches:
                                found_fields.add(field.strip())
                                field_locations.append({
                                    'field': field.strip(),
                                    'slide': slide_num,
                                    'context': text_content[:100] + '...' if len(text_content) > 100 else text_content
                                })
                    
                    if hasattr(shape, 'has_table') and shape.has_table:
                        table = shape.table
                        for row_num, row in enumerate(table.rows):
                            for cell_num, cell in enumerate(row.cells):
                                if cell.text:
                                    text_content = cell.text
                                    matches = re.findall(field_pattern, text_content)
                                    for field in matches:
                                        found_fields.add(field.strip())
                                        field_locations.append({
                                            'field': field.strip(),
                                            'slide': slide_num,
                                            'location': f'Table R{row_num+1}C{cell_num+1}',
                                            'context': text_content[:50] + '...' if len(text_content) > 50 else text_content
                                        })
                except Exception as shape_error:
                    # Skip problematic shapes but continue processing
                    continue
        
        return list(found_fields), field_locations
    except Exception as e:
        st.error(f"Error analyzing PowerPoint: {str(e)}")
        return [], []

def analyze_word_fields(uploaded_file):
    """Analyze a Word document for field placeholders with improved error handling"""
    try:
        # Handle both file objects and file paths
        if isinstance(uploaded_file, str):
            doc = docx.Document(uploaded_file)
        else:
            # Reset file pointer if it's a file object
            if hasattr(uploaded_file, 'seek'):
                uploaded_file.seek(0)
            doc = docx.Document(uploaded_file)
        
        found_fields = set()
        field_pattern = r'\{\{([^}]+)\}\}'
        
        # Check paragraphs
        for para in doc.paragraphs:
            if para.text:
                matches = re.findall(field_pattern, para.text)
                for field in matches:
                    found_fields.add(field.strip())
        
        # Check tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if para.text:
                            matches = re.findall(field_pattern, para.text)
                            for field in matches:
                                found_fields.add(field.strip())
        
        return list(found_fields), []
    except Exception as e:
        st.error(f"Error analyzing Word document: {e}")
        return [], []

# --- Helper and Filling Functions ---
def generate_ai_prompt(fields, project_data):
    """Generate AI prompt with improved formatting"""
    field_descriptions = [f"  - {field}" for field in sorted(fields)]
    prompt = f"""I need you to analyze project data and extract information for specific document fields. Return ONLY a valid JSON object with the field names as keys and extracted values as values.

**Document Fields to Fill:**
{chr(10).join(field_descriptions)}

**Instructions:**
1. Extract relevant information from the data for each field
2. If a field name suggests specific content (e.g., "commander_name" should be a person's name), extract accordingly
3. Be clear, professional, and concise. You are drafting documents for official government use so no slang etc.
4. Conduct market research with a focus on Department of Defense, Department of the Air Force, and with the goals of the 100th ARW and 352nd SOW mission goals in mind
5. For fields with money, phone numbers, or other implied formatting, format the extracted values accordingly
6. For fields you can't determine from the data, use "TBD" or leave reasonable placeholder text based on context
7. Return ONLY the JSON object - no explanations or additional text

**Project Data to Analyze:**
{project_data}

Please analyze the above data and return the JSON object with field values"""
    return prompt

def replace_text_in_paragraph(paragraph, key, value):
    """Replaces text in a paragraph, preserving formatting with improved logic"""
    if key not in paragraph.text:
        return False
    
    try:
        # Simple case - key is in a single run
        for run in paragraph.runs:
            if key in run.text:
                run.text = run.text.replace(key, str(value))
                return True
        
        # Complex case - key spans multiple runs
        runs = paragraph.runs
        full_text = "".join(run.text for run in runs)
        
        if key in full_text:
            # Find which runs contain the key
            new_full_text = full_text.replace(key, str(value), 1)
            
            # Clear all runs and put new text in first run
            for i, run in enumerate(runs):
                if i == 0:
                    run.text = new_full_text
                else:
                    run.text = ""
            return True
    except Exception as e:
        st.warning(f"Could not replace '{key}' in paragraph: {e}")
        return False
    
    return False

def fill_powerpoint_with_data(prs, json_data, uploaded_image=None):
    """Fill PowerPoint with data preserving formatting with improved error handling"""
    replacements_made = 0
    
    try:
        for slide_num, slide in enumerate(prs.slides, 1):
            for shape in slide.shapes:
                try:
                    # Handle text frames
                    if hasattr(shape, "text_frame") and shape.text_frame:
                        for paragraph in shape.text_frame.paragraphs:
                            for field, value in json_data.items():
                                placeholder = f"{{{{{field}}}}}"
                                if replace_text_in_paragraph(paragraph, placeholder, str(value)):
                                    replacements_made += 1
                    
                    # Handle tables
                    elif hasattr(shape, 'has_table') and shape.has_table:
                        for row in shape.table.rows:
                            for cell in row.cells:
                                for paragraph in cell.text_frame.paragraphs:
                                    for field, value in json_data.items():
                                        placeholder = f"{{{{{field}}}}}"
                                        if replace_text_in_paragraph(paragraph, placeholder, str(value)):
                                            replacements_made += 1
                except Exception as shape_error:
                    # Skip problematic shapes but continue
                    continue
        
        return prs, replacements_made
    except Exception as e:
        st.error(f"Error filling PowerPoint: {e}")
        return prs, replacements_made

def fill_word_with_data(doc_file, data):
    """Fill a Word document with data, preserving formatting with improved error handling"""
    try:
        # Handle both file objects and file paths
        if isinstance(doc_file, str):
            doc = docx.Document(doc_file)
        else:
            if hasattr(doc_file, 'seek'):
                doc_file.seek(0)
            doc = docx.Document(doc_file)
        
        replacements_made = 0
        
        # Fill paragraphs
        for paragraph in doc.paragraphs:
            for field, value in data.items():
                placeholder = f"{{{{{field}}}}}"
                if replace_text_in_paragraph(paragraph, placeholder, str(value)):
                    replacements_made += 1
        
        # Fill tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for field, value in data.items():
                            placeholder = f"{{{{{field}}}}}"
                            if replace_text_in_paragraph(paragraph, placeholder, str(value)):
                                replacements_made += 1
        
        return doc, replacements_made
    except Exception as e:
        st.error(f"Error filling Word document: {e}")
        return None, 0

# --- Main Application Logic ---
def main():
    st.warning('**DO NOT ENTER CONTROLLED UNCLASSIFIED INFORMATION INTO THIS SYSTEM**')
    
    # Try to load banner image
    try:
        if os.path.exists("banner.png"):
            st.image("banner.png", use_container_width=True)
    except Exception as e:
        st.info("Info: `banner.png` not found. Skipping image banner.")
    
    st.markdown('<div class="main-header">📊 Document AI Field Filler</div>', unsafe_allow_html=True)
    st.markdown("**Transform your templates with AI-powered data filling! This tool will take unformatted data and conduct research, formatting, organization, data extraction, and place it in a pre-made template or bring your own!**")
    
    # Initialize session state
    if 'fields' not in st.session_state:
        st.session_state.fields = []
    if 'field_locations' not in st.session_state:
        st.session_state.field_locations = []
    if 'ai_prompt' not in st.session_state:
        st.session_state.ai_prompt = ""
    
    st.markdown('<div class="step-container">', unsafe_allow_html=True)
    st.markdown("### 📁 Step 1: Choose Your Template")
    
    # Safely get template files
    try:
        if os.path.exists("templates"):
            template_files = glob.glob("templates/*.*")
            template_options = ["Upload my own template"] + [os.path.basename(f) for f in template_files]
        else:
            template_files = []
            template_options = ["Upload my own template"]
            st.info("Templates directory not found. You can upload your own template.")
    except Exception as e:
        st.warning(f"Could not scan templates directory: {e}")
        template_options = ["Upload my own template"]
        template_files = []
    
    selected_template = st.selectbox("Select a template or upload your own:", options=template_options)
    
    source_file = None
    if selected_template == "Upload my own template":
        source_file = st.file_uploader("Choose your template file", type=['pptx', 'docx'])
    else:
        # Find the full path for the selected template
        template_path = os.path.join("templates", selected_template)
        if os.path.exists(template_path):
            source_file = template_path
        else:
            st.error(f"Template file not found: {template_path}")
    
    st.markdown('</div>', unsafe_allow_html=True)
    
   if source_file is not None:
    filename = source_file.name if hasattr(source_file, 'name') else os.path.basename(source_file)
    file_extension = filename.split('.')[-1].lower()
    
    if file_extension not in ['pptx', 'docx']:
        st.error("Unsupported file type. Please upload a .pptx or .docx file.")
        return
    
    with st.spinner('🔍 Analyzing template fields...'):
        try:
            if file_extension == 'pptx':
                st.session_state.fields, st.session_state.field_locations = analyze_powerpoint_fields(source_file)
            elif file_extension == 'docx':
                st.session_state.fields, st.session_state.field_locations = analyze_word_fields(source_file)
        except Exception as e:
            st.error(f"Error analyzing template: {e}")
            return
    
    if st.session_state.fields:
        st.markdown('<div class="success-box">', unsafe_allow_html=True)
        st.success(f"Found {len(st.session_state.fields)} placeholders in '{filename}'!")
        
        with st.expander("Click to see found fields and their locations"):
            col1, col2 = st.columns(2)
            with col1:
                st.write("**Found Fields:**")
                for field in sorted(st.session_state.fields):
                    st.write(f"• {field}")
            
            with col2:
                st.write("**Field Locations:**")
                if file_extension == 'pptx' and st.session_state.field_locations:
                    for location in st.session_state.field_locations:
                        location_text = f"Slide {location['slide']}"
                        if 'location' in location:
                            location_text += f" - {location['location']}"
                        st.write(f"• **{location['field']}**: {location_text}")
                else:
                    st.write("Location data is not available for Word documents.")
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="step-container">', unsafe_allow_html=True)
        st.markdown("### 📝 Step 2: Enter Your Project Data")
        
        project_data = st.text_area(
            "Paste your raw project data here:", 
            height=200,
            placeholder="Enter your project data, requirements, specifications, or any relevant information here..."
        )
        
        # Image upload for PowerPoint only
        uploaded_image = None
        if file_extension == 'pptx':
            uploaded_image = st.file_uploader(
                "Choose an image file (for PowerPoint only)", 
                type=['png', 'jpg', 'jpeg'],
                help="This image can be used to replace placeholder images in your PowerPoint template"
            )
            if uploaded_image:
                st.image(uploaded_image, caption="Uploaded Image Preview", width=200)
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        if project_data.strip():
            col1, col2 = st.columns([1, 1])
            
            with col1:
                if st.button("🤖 Generate AI Prompt", type="primary", use_container_width=True):
                    with st.spinner("Generating AI prompt..."):
                        st.session_state.ai_prompt = generate_ai_prompt(st.session_state.fields, project_data)
                    st.success("AI prompt generated successfully!")
            
            if st.session_state.ai_prompt:
                st.markdown('<div class="step-container">', unsafe_allow_html=True)
                st.markdown("### 📋 Step 3: Copy Prompt to AI")
                st.info("Copy this prompt and paste it into your preferred AI assistant.")
                
                with st.expander("📄 Click to view the generated AI Prompt", expanded=True):
                    st.code(st.session_state.ai_prompt, language="text")
                
                # Use the fixed copy component
                copy_component("📋 Copy Prompt to Clipboard", st.session_state.ai_prompt)
                
                st.markdown("**Quick Link to AI Service:**")
                st.markdown(
                    f'<a href="https://niprgpt.mil/" target="_blank" class="ai-button nipr-btn">🚀 Open NiprGPT</a>', 
                    unsafe_allow_html=True
                )
                
                st.markdown('</div>', unsafe_allow_html=True)
                
                st.markdown('<div class="step-container">', unsafe_allow_html=True)
                st.markdown("### 🔄 Step 4: Paste AI Response & Generate")
                
                ai_response = st.text_area(
                    "Paste the AI's JSON response here:", 
                    height=150,
                    placeholder='Paste the JSON response from your AI assistant here...\nExample: {"field1": "value1", "field2": "value2"}'
                )
                
                if ai_response.strip():
                    try:
                        # More robust JSON extraction
                        json_str_match = re.search(r'\{.*\}', ai_response, re.DOTALL)
                        if not json_str_match:
                            # Try to find JSON in code blocks
                            code_block_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', ai_response, re.DOTALL)
                            if code_block_match:
                                json_str_match = code_block_match
                                json_data = json.loads(code_block_match.group(1))
                            else:
                                raise json.JSONDecodeError("No JSON object found", "", 0)
                        else:
                            json_data = json.loads(json_str_match.group(0))
                        
                        st.success("✅ Valid JSON detected!")
                        
                        # Show preview of the data
                        with st.expander("Preview extracted data"):
                            st.json(json_data)
                        
                        if st.button("🚀 Generate Filled Document", type="primary", use_container_width=True):
                            progress_container = st.container()
                            
                            with st.spinner('🔄 Filling template...'):
                                try:
                                    output_buffer = io.BytesIO()
                                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                    
                                    if file_extension == 'pptx':
                                        # Reset file pointer for file objects
                                        if hasattr(source_file, 'seek'):
                                            source_file.seek(0)
                                        
                                        prs = Presentation(source_file)
                                        filled_doc, replacements = fill_powerpoint_with_data(
                                            prs, json_data, uploaded_image
                                        )
                                        filled_doc.save(output_buffer)
                                        download_filename = f"filled_presentation_{timestamp}.pptx"
                                        mime_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                                    
                                    elif file_extension == 'docx':
                                        filled_doc, replacements = fill_word_with_data(source_file, json_data)
                                        if filled_doc:
                                            filled_doc.save(output_buffer)
                                            download_filename = f"filled_document_{timestamp}.docx"
                                            mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                        else:
                                            st.error("Failed to process Word document.")
                                            return
                                    
                                    progress_container.success(f"✅ Document generated successfully! Made {replacements} field replacements.")
                                    
                                    # Offer download
                                    if output_buffer.getvalue():
                                        st.download_button(
                                            label=f"📥 Download Filled {file_extension.upper()}",
                                            data=output_buffer.getvalue(),
                                            file_name=download_filename,
                                            mime=mime_type,
                                            use_container_width=True
                                        )
                                        st.balloons()
                                    else:
                                        st.error("Failed to generate document - output is empty.")
                                
                                except Exception as generation_error:
                                    st.error(f"Error generating document: {generation_error}")
                    
                    except json.JSONDecodeError as e:
                        st.error(f"❌ Invalid JSON format: {e}")
                        st.info("Make sure you paste only the JSON object from the AI response. It should start with { and end with }")
                    
                    except Exception as e:
                        st.error(f"❌ Error processing AI response: {e}")
                
                st.markdown('</div>', unsafe_allow_html=True)
    
    else:
        st.markdown('<div class="warning-box">', unsafe_allow_html=True)
        st.warning("⚠️ No {{field_name}} placeholders found in your template!")
        st.info("""
        To use this tool, your template should contain placeholders in the format {{field_name}}.
        For example: {{project_name}}, {{date}}, {{commander_name}}, etc.
        """)
        st.markdown('</div>', unsafe_allow_html=True)rue)
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; padding: 20px;">
        <p>🚀 Built for NIPR environments • No local installation required • Works in any browser</p>
        <p><small>Version 2.0 - Enhanced error handling and improved user experience</small></p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        st.error(f"Application error: {e}")
        st.info("Please refresh the page and try again. If the problem persists, check your template file format.")

