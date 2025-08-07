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
import docx
import glob

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
    .nipr-btn { background-color: #004d40; } /* Dark Teal for NiprGPT */
</style>
""", unsafe_allow_html=True)

# --- Analysis Functions ---

def analyze_powerpoint_fields(uploaded_file):
    """(Corrected) Analyze PowerPoint file for field placeholders"""
    try:
        prs = Presentation(uploaded_file)
        found_fields = set()
        field_locations = []
        
        for slide_num, slide in enumerate(prs.slides, 1):
            for shape in slide.shapes:
                if hasattr(shape, "text_frame") and shape.text_frame and shape.text_frame.text:
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
                            if cell.text:
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

def analyze_word_fields(uploaded_file):
    """(Corrected) Analyze a Word document for field placeholders."""
    try:
        doc = docx.Document(uploaded_file)
        found_fields = set()
        field_pattern = r'\{\{([^}]+)\}\}'

        for para in doc.paragraphs:
            if para.text:
                matches = re.findall(field_pattern, para.text)
                for field in matches:
                    found_fields.add(field)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if para.text:
                            matches = re.findall(field_pattern, para.text)
                            for field in matches:
                                found_fields.add(field)
        
        return list(found_fields), []
        
    except Exception as e:
        st.error(f"Error analyzing Word document: {e}")
        return [], []

# --- Helper and Filling Functions ---

def generate_ai_prompt(fields, project_data):
    """Generate AI prompt"""
    field_descriptions = [f"  - {field}" for field in sorted(fields)]
    
    prompt = f"""I need you to analyze project data and extract information for specific PowerPoint fields. Return ONLY a valid JSON object with the field names as keys and extracted values as values.
**PowerPoint Fields to Fill:**

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
    """(NEW) Replaces text in a paragraph, preserving formatting.
    This is a robust function for both pptx and docx.
    """
    if key not in paragraph.text:
        return
    
    # Replace the simple case
    for run in paragraph.runs:
        if key in run.text:
            run.text = run.text.replace(key, value)
            return

    # Handle cases where the key is split across multiple runs
    runs = paragraph.runs
    full_text = "".join(run.text for run in runs)
    if key in full_text:
        start_index = full_text.find(key)
        end_index = start_index + len(key)
        
        current_pos = 0
        runs_to_modify = []
        for run in runs:
            run_len = len(run.text)
            if current_pos < end_index and current_pos + run_len > start_index:
                runs_to_modify.append(run)
            current_pos += run_len
        
        if runs_to_modify:
            # Replace text in the first run and clear others
            original_text = "".join(run.text for run in runs_to_modify)
            new_text = original_text.replace(key, value, 1)
            runs_to_modify[0].text = new_text
            for i in range(1, len(runs_to_modify)):
                runs_to_modify[i].text = ""

def fill_powerpoint_with_data(prs, json_data, uploaded_image, progress_container):
    """(CORRECTED) Fill PowerPoint with data preserving formatting."""
    replacements_made = 0
    if uploaded_image:
        # Placeholder for image replacement logic
        pass

    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text_frame") and shape.text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for field, value in json_data.items():
                        placeholder = f"{{{{{field}}}}}"
                        replace_text_in_paragraph(paragraph, placeholder, str(value))
                        replacements_made += 1 # Count attempt
            elif shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for paragraph in cell.text_frame.paragraphs:
                            for field, value in json_data.items():
                                placeholder = f"{{{{{field}}}}}"
                                replace_text_in_paragraph(paragraph, placeholder, str(value))
                                replacements_made += 1 # Count attempt
    return prs, replacements_made

def fill_word_with_data(doc_file, data):
    """Fill a Word document with data, preserving formatting."""
    doc = docx.Document(doc_file)
    for paragraph in doc.paragraphs:
        for field, value in data.items():
            placeholder = f"{{{{{field}}}}}"
            replace_text_in_paragraph(paragraph, placeholder, str(value))
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for field, value in data.items():
                        placeholder = f"{{{{{field}}}}}"
                        replace_text_in_paragraph(paragraph, placeholder, str(value))
    return doc

# --- Main Application Logic ---

def main():
    st.warning('**DO NOT ENTER CONTROLLED UNCLASSIFIED INFORMATION INTO THIS SYSTEM**')
    try:
        st.image("banner.png", use_container_width=True)
    except Exception as e:
        st.info("Info: `banner.png` not found. Skipping image banner.")

    st.markdown('<div class="main-header">📊 Document AI Field Filler</div>', unsafe_allow_html=True)
    st.markdown("**Transform your templates with AI-powered data filling! This tool will take unformatted data and conduct research, formatting, organization, data extraction, and place it in a pre-made template or bring your own!**")

    if 'fields' not in st.session_state:
        st.session_state.fields = []
    if 'field_locations' not in st.session_state:
        st.session_state.field_locations = []
    if 'ai_prompt' not in st.session_state:
        st.session_state.ai_prompt = ""

    st.markdown('<div class="step-container">', unsafe_allow_html=True)
    st.markdown("### 📁 Step 1: Choose Your Template")

    try:
        template_files = glob.glob("templates/*.*")
        template_options = ["Upload my own template"] + template_files
    except Exception as e:
        st.error(f"Could not scan templates directory: {e}")
        template_options = ["Upload my own template"]

    selected_template = st.selectbox("Select a template or upload your own:", options=template_options)
    source_file = None 

    if selected_template == "Upload my own template":
        source_file = st.file_uploader("Choose your template file", type=['pptx', 'docx'])
    else:
        source_file = selected_template
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    if source_file is not None:
        filename = source_file.name if hasattr(source_file, 'name') else source_file
        file_extension = filename.split('.')[-1].lower()

        with st.spinner('🔍 Analyzing template fields...'):
            if file_extension == 'pptx':
                st.session_state.fields, st.session_state.field_locations = analyze_powerpoint_fields(source_file)
            elif file_extension == 'docx':
                st.session_state.fields, st.session_state.field_locations = analyze_word_fields(source_file)
            else:
                st.error("Unsupported file type.")
                return

        if st.session_state.fields:
            st.markdown('<div class="success-box">', unsafe_allow_html=True)
            st.success(f"Found {len(st.session_state.fields)} placeholders in '{filename}'!")
            
            with st.expander("Click to see found fields and their locations"):
                col1, col2 = st.columns(2)
                with col1:
                    st.write("**Found Fields:**")
                    st.write(st.session_state.fields)
                with col2:
                    st.write("**Field Locations (PowerPoint only):**")
                    if file_extension == 'pptx' and st.session_state.field_locations:
                        st.write(st.session_state.field_locations)
                    else:
                        st.write("Location data is not available for Word documents.")

            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown('<div class="step-container">', unsafe_allow_html=True)
            st.markdown("### 📝 Step 2: Enter Your Project Data")
            project_data = st.text_area("Paste your raw project data here:", height=200)
            
            uploaded_image = None
            if file_extension == 'pptx':
                uploaded_image = st.file_uploader("Choose an image file (for PowerPoint only)", type=['png', 'jpg', 'jpeg'])
                if uploaded_image:
                    st.image(uploaded_image, caption="Uploaded Image Preview", width=200)

            st.markdown('</div>', unsafe_allow_html=True)

            if project_data.strip():
                if st.button("🤖 Generate AI Prompt", type="primary"):
                    st.session_state.ai_prompt = generate_ai_prompt(st.session_state.fields, project_data)
                
                if st.session_state.ai_prompt:
                    st.markdown('<div class="step-container">', unsafe_allow_html=True)
                    st.markdown("### 📋 Step 3: Copy Prompt to AI")
                    st.info("Copy this prompt and paste it into your preferred AI assistant.")
                    
                    with st.expander("📄 Click to view the generated AI Prompt", expanded=True):
                        st.code(st.session_state.ai_prompt, language="text")
                    
                    copy_component("📋 Copy Prompt to Clipboard", st.session_state.ai_prompt)

                    st.markdown("**Quick Link to AI Service:**")
                    st.markdown(f'<a href="https://niprgpt.mil/" target="_blank" class="ai-button nipr-btn">🚀 Open NiprGPT</a>', unsafe_allow_html=True)
                    st.markdown('</div>', unsafe_allow_html=True)

                    st.markdown('<div class="step-container">', unsafe_allow_html=True)
                    st.markdown("### 🔄 Step 4: Paste AI Response & Generate")
                    ai_response = st.text_area("Paste the AI's JSON response here:", height=150)

                    if ai_response.strip():
                        try:
                            json_str_match = re.search(r'\{.*\}', ai_response, re.DOTALL)
                            if not json_str_match:
                                raise json.JSONDecodeError("No JSON object found", "", 0)
                            json_data = json.loads(json_str_match.group(0))
                            st.success("✅ Valid JSON detected!")

                            if st.button("🚀 Generate Filled Document", type="primary"):
                                progress_container = st.container()
                                with st.spinner('🔄 Filling template...'):
                                    if hasattr(source_file, 'seek'):
                                        source_file.seek(0)
                                    
                                    output_buffer = io.BytesIO()
                                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                    
                                    if file_extension == 'pptx':
                                        prs = Presentation(source_file)
                                        filled_doc, _ = fill_powerpoint_with_data(prs, json_data, uploaded_image, progress_container)
                                        filled_doc.save(output_buffer)
                                        download_filename = f"filled_presentation_{timestamp}.pptx"
                                        mime_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                                    elif file_extension == 'docx':
                                        filled_doc = fill_word_with_data(source_file, json_data)
                                        filled_doc.save(output_buffer)
                                        download_filename = f"filled_document_{timestamp}.docx"
                                        mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                    
                                    progress_container.success("✅ Document generated successfully!")
                                    st.download_button(
                                        label=f"📥 Download Filled {file_extension.upper()}",
                                        data=output_buffer.getvalue(),
                                        file_name=download_filename,
                                        mime=mime_type
                                    )
                                    st.balloons()
                        except json.JSONDecodeError as e:
                            st.error(f"❌ Invalid JSON format: {e}")
                    st.markdown('</div>', unsafe_allow_html=True)
        elif source_file is not None:
            st.markdown('<div class="warning-box">', unsafe_allow_html=True)
            st.warning("⚠️ No {{field_name}} placeholders found in your template!")
            st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; padding: 20px;">
        <p>🚀 Built for NIPR environments • No local installation required • Works in any browser</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()

