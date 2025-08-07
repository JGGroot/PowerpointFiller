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

# PDF support imports
import fitz  # PyMuPDF
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import PyPDF2

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
        background: #330066;
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
        background-color: #fff000;
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

# Google Analytics - Replace 'G-HMVVJJ6C17' with your actual Google Analytics ID
st.markdown("""
<!-- Google tag (gtag.js) -->
<script async src="https://www.googletagmanager.com/gtag/js?id=G-HMVVJJ6C17"></script>
<script>
  window.dataLayer = window.dataLayer || [];
  function gtag(){dataLayer.push(arguments);}
  gtag('js', new Date());
  gtag('config', 'G-HMVVJJ6C17');
</script>
""", unsafe_allow_html=True)

# --- Analysis Functions ---

def analyze_pdf_fields(uploaded_file):
    """Analyze PDF file for field placeholders and form fields"""
    try:
        # Reset file pointer
        if hasattr(uploaded_file, 'seek'):
            uploaded_file.seek(0)
        
        # Read PDF with PyMuPDF
        pdf_bytes = uploaded_file.read()
        pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        found_fields = set()
        field_locations = []
        field_pattern = r'\{\{([^}]+)\}\}'
        
        # Method 1: Extract text and look for {{field_name}} patterns
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            text_content = page.get_text()
            
            # Find field patterns in text
            matches = re.findall(field_pattern, text_content)
            for field in matches:
                found_fields.add(field)
                field_locations.append({
                    'field': field,
                    'page': page_num + 1,
                    'type': 'text',
                    'context': text_content[:100] + '...' if len(text_content) > 100 else text_content
                })
        
        # Method 2: Check for form fields (if it's a fillable PDF)
        try:
            # Reset file pointer for PyPDF2
            if hasattr(uploaded_file, 'seek'):
                uploaded_file.seek(0)
            
            pdf_reader = PyPDF2.PdfReader(uploaded_file)
            
            if pdf_reader.is_encrypted:
                st.warning("PDF is encrypted. Form field detection may be limited.")
            
            # Check each page for form fields
            for page_num, page in enumerate(pdf_reader.pages):
                if '/Annots' in page:
                    annotations = page['/Annots']
                    if annotations:
                        for annotation_ref in annotations:
                            annotation = annotation_ref.get_object()
                            if annotation.get('/Subtype') == '/Widget':
                                field_name = annotation.get('/T')
                                if field_name:
                                    field_name_str = field_name
                                    # Check if field name contains our pattern
                                    pattern_matches = re.findall(field_pattern, field_name_str)
                                    if pattern_matches:
                                        for field in pattern_matches:
                                            found_fields.add(field)
                                            field_locations.append({
                                                'field': field,
                                                'page': page_num + 1,
                                                'type': 'form_field',
                                                'field_name': field_name_str
                                            })
                                    else:
                                        # Add the form field name itself as a potential field
                                        found_fields.add(field_name_str)
                                        field_locations.append({
                                            'field': field_name_str,
                                            'page': page_num + 1,
                                            'type': 'form_field',
                                            'field_name': field_name_str
                                        })
        except Exception as e:
            st.warning(f"Could not analyze PDF form fields: {e}")
        
        pdf_document.close()
        return list(found_fields), field_locations
        
    except Exception as e:
        st.error(f"Error analyzing PDF: {str(e)}")
        return [], []

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
    """(FIXED) Analyze a Word document for field placeholders including text boxes."""
    try:
        doc = docx.Document(uploaded_file)
        found_fields = set()
        field_pattern = r'\{\{([^}]+)\}\}'

        # Check regular paragraphs
        for para in doc.paragraphs:
            if para.text:
                matches = re.findall(field_pattern, para.text)
                for field in matches:
                    found_fields.add(field)

        # Check tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if para.text:
                            matches = re.findall(field_pattern, para.text)
                            for field in matches:
                                found_fields.add(field)

        # Check text boxes and shapes (NEW CODE)
        try:
            from docx.oxml.ns import nsdecls, qn
            from lxml import etree
            
            # Define namespaces explicitly
            namespaces = {
                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
            }
            
            # Get the document's XML tree
            doc_xml = doc.element
            
            # Look for text boxes using various possible paths
            textbox_paths = [
                './/w:drawing//w:txbxContent//w:p//w:t',
                './/w:drawing//a:txBody//a:p//a:t', 
                './/w:object//w:drawing//w:txbxContent//w:p//w:t',
                './/w:pict//w:textbox//w:txbxContent//w:p//w:t'
            ]
            
            for path in textbox_paths:
                try:
                    text_elements = doc_xml.xpath(path, namespaces=namespaces)
                    for t_elem in text_elements:
                        if t_elem.text:
                            matches = re.findall(field_pattern, t_elem.text)
                            for field in matches:
                                found_fields.add(field)
                except Exception:
                    # Skip this path if it doesn't work
                    continue
            
            # Alternative approach: search through all text elements in the document
            # This is a broader search that should catch text boxes regardless of structure
            try:
                all_text_elements = doc_xml.xpath('.//w:t', namespaces=namespaces)
                for t_elem in all_text_elements:
                    if t_elem.text:
                        matches = re.findall(field_pattern, t_elem.text)
                        for field in matches:
                            found_fields.add(field)
            except Exception:
                pass
                
        except Exception as e:
            # If XML parsing fails, fall back to a simpler approach
            st.warning(f"Advanced text box detection failed: {e}. Using basic detection only.")

        # Alternative approach: check headers and footers which might contain text boxes
        for section in doc.sections:
            # Check headers
            if section.header:
                for para in section.header.paragraphs:
                    if para.text:
                        matches = re.findall(field_pattern, para.text)
                        for field in matches:
                            found_fields.add(field)
            
            # Check footers  
            if section.footer:
                for para in section.footer.paragraphs:
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

def fill_pdf_with_data(pdf_file, data):
    """Fill PDF with data using PyMuPDF (fitz)"""
    try:
        # Reset file pointer
        if hasattr(pdf_file, 'seek'):
            pdf_file.seek(0)
        
        pdf_bytes = pdf_file.read()
        pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        field_pattern = r'\{\{([^}]+)\}\}'
        replacements_made = 0
        
        # Method 1: Replace text content (for text-based PDFs)
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            
            # Get all text instances
            text_instances = page.get_text("dict")
            
            # Look for and replace field patterns
            for block in text_instances["blocks"]:
                if "lines" in block:
                    for line in block["lines"]:
                        for span in line["spans"]:
                            text = span.get("text", "")
                            
                            # Check if this text contains any of our fields
                            for field, value in data.items():
                                placeholder = f"{{{{{field}}}}}"
                                if placeholder in text:
                                    # Get the rectangle coordinates
                                    rect = fitz.Rect(span["bbox"])
                                    
                                    # Remove the old text by drawing a white rectangle over it
                                    page.add_redact_annot(rect, fill=(1, 1, 1))
                                    page.apply_redactions()
                                    
                                    # Add the new text
                                    new_text = text.replace(placeholder, str(value))
                                    page.insert_text(
                                        rect.top_left,
                                        new_text,
                                        fontsize=span.get("size", 12),
                                        color=(0, 0, 0)
                                    )
                                    replacements_made += 1
        
        # Method 2: Try to handle form fields (if it's a fillable PDF)
        try:
            if hasattr(pdf_file, 'seek'):
                pdf_file.seek(0)
            
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            pdf_writer = PyPDF2.PdfWriter()
            
            # If there are form fields, try to fill them
            for page in pdf_reader.pages:
                if '/Annots' in page:
                    annotations = page['/Annots']
                    if annotations:
                        for annotation_ref in annotations:
                            annotation = annotation_ref.get_object()
                            if annotation.get('/Subtype') == '/Widget':
                                field_name = annotation.get('/T')
                                if field_name:
                                    # Check if we have data for this field
                                    for field, value in data.items():
                                        placeholder = f"{{{{{field}}}}}"
                                        if placeholder in field_name or field == field_name:
                                            # Update form field value
                                            if '/V' in annotation:
                                                annotation.update({PyPDF2.generic.NameObject('/V'): 
                                                                 PyPDF2.generic.TextStringObject(str(value))})
                                                replacements_made += 1
                
                pdf_writer.add_page(page)
            
            # Create output buffer for form-filled PDF
            if replacements_made > 0:
                form_output = io.BytesIO()
                pdf_writer.write(form_output)
                form_output.seek(0)
                
                # Reopen with fitz to continue with text replacements
                pdf_document = fitz.open(stream=form_output.getvalue(), filetype="pdf")
        
        except Exception as e:
            st.warning(f"Form field filling partially failed: {e}. Continuing with text replacement.")
        
        # Save the modified PDF
        output_buffer = io.BytesIO()
        pdf_document.save(output_buffer)
        pdf_document.close()
        
        return output_buffer.getvalue(), replacements_made
        
    except Exception as e:
        st.error(f"Error filling PDF: {str(e)}")
        return None, 0

def fill_powerpoint_with_data(prs, json_data, uploaded_image, progress_container):
    """(CORRECTED) Fill PowerPoint with data preserving formatting."""
    replacements_made = 0
    # Image replacement functionality temporarily disabled
    # if uploaded_image:
    #     # Placeholder for image replacement logic
    #     pass

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
    """(FIXED) Fill a Word document with data, preserving formatting and handling text boxes."""
    doc = docx.Document(doc_file)
    
    # Fill regular paragraphs
    for paragraph in doc.paragraphs:
        for field, value in data.items():
            placeholder = f"{{{{{field}}}}}"
            replace_text_in_paragraph(paragraph, placeholder, str(value))
    
    # Fill tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for field, value in data.items():
                        placeholder = f"{{{{{field}}}}}"
                        replace_text_in_paragraph(paragraph, placeholder, str(value))
    
    # Fill text boxes (NEW CODE)
    try:
        from lxml import etree
        
        # Define namespaces explicitly
        namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
        }
        
        # Get the document's XML tree
        doc_xml = doc.element
        
        # Look for text elements in text boxes using various paths
        textbox_paths = [
            './/w:drawing//w:txbxContent//w:p//w:t',
            './/w:drawing//a:txBody//a:p//a:t',
            './/w:object//w:drawing//w:txbxContent//w:p//w:t',
            './/w:pict//w:textbox//w:txbxContent//w:p//w:t'
        ]
        
        for path in textbox_paths:
            try:
                text_elements = doc_xml.xpath(path, namespaces=namespaces)
                
                for t_elem in text_elements:
                    if t_elem.text:
                        original_text = t_elem.text
                        modified_text = original_text
                        
                        # Replace all placeholders
                        for field, value in data.items():
                            placeholder = f"{{{{{field}}}}}"
                            modified_text = modified_text.replace(placeholder, str(value))
                        
                        # Update the text if it was modified
                        if modified_text != original_text:
                            t_elem.text = modified_text
                            
            except Exception:
                # Skip this path if it doesn't work
                continue
                
        # Fallback: try to replace in all text elements
        try:
            all_text_elements = doc_xml.xpath('.//w:t', namespaces=namespaces)
            for t_elem in all_text_elements:
                if t_elem.text:
                    original_text = t_elem.text
                    modified_text = original_text
                    
                    for field, value in data.items():
                        placeholder = f"{{{{{field}}}}}"
                        if placeholder in modified_text:
                            modified_text = modified_text.replace(placeholder, str(value))
                    
                    if modified_text != original_text:
                        t_elem.text = modified_text
                        
        except Exception:
            pass
    
    except Exception as e:
        st.warning(f"Advanced text box filling failed: {e}. Basic filling completed.")
    
    # Fill headers and footers
    try:
        for section in doc.sections:
            # Fill headers
            if section.header:
                for paragraph in section.header.paragraphs:
                    for field, value in data.items():
                        placeholder = f"{{{{{field}}}}}"
                        replace_text_in_paragraph(paragraph, placeholder, str(value))
            
            # Fill footers
            if section.footer:
                for paragraph in section.footer.paragraphs:
                    for field, value in data.items():
                        placeholder = f"{{{{{field}}}}}"
                        replace_text_in_paragraph(paragraph, placeholder, str(value))
    
    except Exception as e:
        st.error(f"Error filling headers/footers: {e}")
    
    return doc

# --- Main Application Logic ---

def main():
    st.warning('**DO NOT ENTER CUI OR PII INTO THIS SYSTEM - FOR BETA TESTING AND NON-OFFICIAL USE ONLY**')
    try:
        st.image("banner.png", use_container_width=True)
    except Exception as e:
        st.info("Info: `banner.png` not found. Skipping image banner.")

    st.markdown('<div class="main-header">📊 Document AI Field Filler</div>', unsafe_allow_html=True)
    st.markdown("**Transform your templates with AI-powered data filling! This tool supports PowerPoint, Word, and PDF documents. It will take unformatted data and conduct research, formatting, organization, data extraction, and place it in a pre-made template or bring your own!**")

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
        source_file = st.file_uploader("Choose your template file", type=['pptx', 'docx', 'pdf'])
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
            elif file_extension == 'pdf':
                st.session_state.fields, st.session_state.field_locations = analyze_pdf_fields(source_file)
            else:
                st.error("Unsupported file type. Supported formats: PowerPoint (.pptx), Word (.docx), PDF (.pdf)")
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
                    st.write("**Field Locations:**")
                    if st.session_state.field_locations:
                        for location in st.session_state.field_locations:
                            if file_extension == 'pdf':
                                st.write(f"• **{location['field']}** (Page {location['page']}, Type: {location['type']})")
                            elif file_extension == 'pptx':
                                st.write(f"• **{location['field']}** (Slide {location['slide']})")
                    else:
                        st.write("Location data not available for this document type.")

            st.markdown('</div>', unsafe_allow_html=True)

            # Create tabs for AI Generation and Manual Entry
            tab1, tab2 = st.tabs(["🤖 AI Generation", "✏️ Manual Entry"])
            
            # AI Generation Tab
            with tab1:
                st.markdown('<div class="step-container">', unsafe_allow_html=True)
                st.markdown("### 📝 Step 2: Enter Your Applicable Data Or Text. This can be formatted in any way, stream of thought, lists, sentences, etc. The more you provide the better your result will be. Any field on your template that is not covered will be TBD")
                project_data = st.text_area("Enter your data here:", height=200)
                
                # Image upload temporarily disabled for all formats
                uploaded_image = None

                st.markdown('</div>', unsafe_allow_html=True)

                if project_data.strip():
                    if st.button("🤖 Generate AI Prompt", type="primary", key="ai_prompt_btn"):
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

                                if st.button("🚀 Generate Filled Document", type="primary", key="ai_generate_btn"):
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
                                        elif file_extension == 'pdf':
                                            filled_pdf_bytes, replacements = fill_pdf_with_data(source_file, json_data)
                                            if filled_pdf_bytes:
                                                output_buffer.write(filled_pdf_bytes)
                                                download_filename = f"filled_document_{timestamp}.pdf"
                                                mime_type = "application/pdf"
                                                progress_container.success(f"✅ PDF generated successfully! Made {replacements} replacements.")
                                            else:
                                                st.error("Failed to generate filled PDF")
                                                continue
                                        
                                        if file_extension != 'pdf':
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
            
            # Manual Entry Tab
            with tab2:
                st.markdown('<div class="step-container">', unsafe_allow_html=True)
                st.markdown("### ✏️ Manual Entry - Fill Fields Directly")
                st.info(f"Found {len(st.session_state.fields)} fields to fill. Leave any field blank to keep the placeholder in the document.")
                
                # Initialize session state for manual entry if not exists
                if 'manual_entry_data' not in st.session_state:
                    st.session_state.manual_entry_data = {}
                
                # Create input fields for each found field
                st.markdown("**Fill in the fields below:**")
                
                # Split fields into two columns
                col1, col2 = st.columns(2)
                fields_per_column = (len(st.session_state.fields) + 1) // 2
                
                with col1:
                    for field in st.session_state.fields[:fields_per_column]:
                        # Use session state to preserve values
                        if field not in st.session_state.manual_entry_data:
                            st.session_state.manual_entry_data[field] = ""
                        
                        st.session_state.manual_entry_data[field] = st.text_input(
                            f"**{field}**",
                            value=st.session_state.manual_entry_data[field],
                            key=f"manual_field_{field}_1",
                            help=f"Enter value for {{{{ {field} }}}}"
                        )
                
                with col2:
                    for field in st.session_state.fields[fields_per_column:]:
                        # Use session state to preserve values
                        if field not in st.session_state.manual_entry_data:
                            st.session_state.manual_entry_data[field] = ""
                        
                        st.session_state.manual_entry_data[field] = st.text_input(
                            f"**{field}**",
                            value=st.session_state.manual_entry_data[field],
                            key=f"manual_field_{field}_2",
                            help=f"Enter value for {{{{ {field} }}}}"
                        )
                
                # Add utility buttons and generation
                st.markdown("---")
                col_clear, col_preview, col_generate = st.columns(3)
                
                with col_clear:
                    if st.button("🗑️ Clear All Fields", help="Clear all entered data"):
                        for field in st.session_state.fields:
                            st.session_state.manual_entry_data[field] = ""
                        st.rerun()
                
                with col_preview:
                    # Show preview of filled vs empty fields
                    filled_count = sum(1 for field in st.session_state.fields 
                                     if st.session_state.manual_entry_data.get(field, "").strip())
                    st.metric("Fields to Fill", f"{filled_count}/{len(st.session_state.fields)}")
                
                with col_generate:
                    if st.button("🚀 Generate Document", type="primary", key="manual_generate_btn"):
                        # Prepare data dictionary, excluding empty fields
                        manual_data = {}
                        filled_count = 0
                        
                        for field in st.session_state.fields:
                            value = st.session_state.manual_entry_data.get(field, "").strip()
                            if value:  # Only include non-empty fields
                                manual_data[field] = value
                                filled_count += 1
                        
                        st.info(f"Filling {filled_count} out of {len(st.session_state.fields)} fields. Empty fields will remain as placeholders.")
                        
                        progress_container = st.container()
                        with st.spinner('🔄 Generating document with manual entry...'):
                            if hasattr(source_file, 'seek'):
                                source_file.seek(0)
                            
                            output_buffer = io.BytesIO()
                            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                            
                            if file_extension == 'pptx':
                                prs = Presentation(source_file)
                                filled_doc, _ = fill_powerpoint_with_data(prs, manual_data, None, progress_container)
                                filled_doc.save(output_buffer)
                                download_filename = f"manual_filled_presentation_{timestamp}.pptx"
                                mime_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                            elif file_extension == 'docx':
                                filled_doc = fill_word_with_data(source_file, manual_data)
                                filled_doc.save(output_buffer)
                                download_filename = f"manual_filled_document_{timestamp}.docx"
                                mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            elif file_extension == 'pdf':
                                filled_pdf_bytes, replacements = fill_pdf_with_data(source_file, manual_data)
                                if filled_pdf_bytes:
                                    output_buffer.write(filled_pdf_bytes)
                                    download_filename = f"manual_filled_document_{timestamp}.pdf"
                                    mime_type = "application/pdf"
                                    progress_container.success(f"✅ PDF generated successfully! Made {replacements} replacements.")
                                else:
                                    st.error("Failed to generate filled PDF")
                                    continue
                            
                            if file_extension != 'pdf':
                                progress_container.success("✅ Document generated successfully with manual entry!")
                            
                            st.download_button(
                                label=f"📥 Download Manual Filled {file_extension.upper()}",
                                data=output_buffer.getvalue(),
                                file_name=download_filename,
                                mime=mime_type
                            )
                            st.balloons()
                
                # Show a detailed preview of what will be filled
                with st.expander("📋 Preview of Field Mappings", expanded=False):
                    preview_data = []
                    for field in sorted(st.session_state.fields):
                        value = st.session_state.manual_entry_data.get(field, "").strip()
                        status = "✅ Will be filled" if value else "⚪ Will remain as placeholder"
                        preview_data.append({
                            "Field": f"{{{{{field}}}}}",
                            "Value": value if value else "(empty)",
                            "Status": status
                        })
                    
                    if preview_data:
                        preview_df = pd.DataFrame(preview_data)
                        st.dataframe(preview_df, use_container_width=True, hide_index=True)
                
                st.markdown('</div>', unsafe_allow_html=True)

        elif source_file is not None:
            st.markdown('<div class="warning-box">', unsafe_allow_html=True)
            st.warning("⚠️ No {{field_name}} placeholders found in your template!")
            
            if file_extension == 'pdf':
                st.info("""
                **PDF Tips:**
                - For text-based PDFs: Add placeholders like {{field_name}} in the text
                - For form-based PDFs: Use form field names that match your data fields
                - Some complex PDF structures may not be fully supported
                """)
            
            st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; padding: 20px;">
        <p>🚀 Built for NIPR environments • No local installation required • Works in any browser</p>
        <p>📄 Supports PowerPoint (.pptx), Word (.docx), and PDF (.pdf) templates</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
