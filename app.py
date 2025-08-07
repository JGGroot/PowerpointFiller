import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pyperclip
from pptx import Presentation
from pptx.util import Inches
from PIL import Image, ImageTk
import json
import re
import os
import webbrowser
from datetime import datetime

class PowerPointFillerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PowerPoint AI Field Filler")
        self.root.geometry("800x700")
        self.root.configure(bg='#f0f0f0')
        
        # Variables
        self.pptx_path = tk.StringVar()
        self.found_fields = []
        self.ai_prompt = ""
        self.project_data = ""
        self.uploaded_image = None  # Store uploaded image path
        
        self.setup_ui()
        
    def setup_ui(self):
        """Create the user interface"""
        
        # Main title
        title_label = tk.Label(
            self.root, 
            text="PowerPoint AI Field Filler", 
            font=('Arial', 18, 'bold'),
            bg='#f0f0f0',
            fg='#2c3e50'
        )
        title_label.pack(pady=20)
        
        # Create notebook for steps
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Step 1: File Selection
        self.create_step1()
        
        # Step 2: Project Data Input
        self.create_step2()
        
        # Step 3: AI Response and Output
        self.create_step3()
        
        # AI Service buttons at bottom
        self.create_ai_buttons()
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready - Select your PowerPoint file to begin")
        status_bar = tk.Label(
            self.root, 
            textvariable=self.status_var, 
            relief=tk.SUNKEN, 
            anchor=tk.W,
            bg='#ecf0f1'
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def create_step1(self):
        """Step 1: PowerPoint file selection and field analysis"""
        
        step1_frame = ttk.Frame(self.notebook)
        self.notebook.add(step1_frame, text="1. Select PowerPoint")
        
        # Instructions
        instructions = tk.Label(
            step1_frame,
            text="Select your PowerPoint template with {{field_name}} placeholders",
            font=('Arial', 12),
            wraplength=700,
            justify='left'
        )
        instructions.pack(pady=20)
        
        # File selection frame
        file_frame = tk.Frame(step1_frame)
        file_frame.pack(pady=10, fill='x', padx=20)
        
        tk.Label(file_frame, text="PowerPoint File:", font=('Arial', 10, 'bold')).pack(anchor='w')
        
        file_select_frame = tk.Frame(file_frame)
        file_select_frame.pack(fill='x', pady=5)
        
        self.file_entry = tk.Entry(file_select_frame, textvariable=self.pptx_path, font=('Arial', 10))
        self.file_entry.pack(side='left', fill='x', expand=True, padx=(0, 10))
        
        browse_btn = tk.Button(
            file_select_frame, 
            text="Browse", 
            command=self.browse_file,
            bg='#3498db',
            fg='white',
            font=('Arial', 10, 'bold')
        )
        browse_btn.pack(side='right')
        
        # Analyze button
        analyze_btn = tk.Button(
            step1_frame,
            text="Analyze PowerPoint Fields",
            command=self.analyze_fields,
            bg='#27ae60',
            fg='white',
            font=('Arial', 12, 'bold'),
            height=2
        )
        analyze_btn.pack(pady=20)
        
        # Results display
        results_label = tk.Label(step1_frame, text="Found Fields:", font=('Arial', 10, 'bold'))
        results_label.pack(anchor='w', padx=20)
        
        self.fields_text = scrolledtext.ScrolledText(
            step1_frame, 
            height=15, 
            width=80,
            font=('Courier', 9)
        )
        self.fields_text.pack(pady=10, padx=20, fill='both', expand=True)
    
    def create_step2(self):
        """Step 2: Project data input and AI prompt generation"""
        
        step2_frame = ttk.Frame(self.notebook)
        self.notebook.add(step2_frame, text="2. Input Project Data")
        
        # Instructions
        instructions = tk.Label(
            step2_frame,
            text="Paste your raw project data below, then generate the AI prompt:",
            font=('Arial', 12),
            wraplength=700,
            justify='left'
        )
        instructions.pack(pady=20)
        
        # Project data input
        data_label = tk.Label(step2_frame, text="Project Data:", font=('Arial', 10, 'bold'))
        data_label.pack(anchor='w', padx=20)
        
        self.project_text = scrolledtext.ScrolledText(
            step2_frame, 
            height=8, 
            width=80,
            font=('Arial', 10)
        )
        self.project_text.pack(pady=10, padx=20, fill='x')
        
        # Image upload section
        image_frame = tk.Frame(step2_frame)
        image_frame.pack(pady=10, fill='x', padx=20)
        
        image_label = tk.Label(image_frame, text="Product Image (Optional):", font=('Arial', 10, 'bold'))
        image_label.pack(anchor='w')
        
        image_upload_frame = tk.Frame(image_frame)
        image_upload_frame.pack(fill='x', pady=5)
        
        self.image_path_var = tk.StringVar()
        self.image_entry = tk.Entry(image_upload_frame, textvariable=self.image_path_var, font=('Arial', 10), state='readonly')
        self.image_entry.pack(side='left', fill='x', expand=True, padx=(0, 10))
        
        browse_image_btn = tk.Button(
            image_upload_frame,
            text="Browse Image",
            command=self.browse_image,
            bg='#e67e22',
            fg='white',
            font=('Arial', 10, 'bold')
        )
        browse_image_btn.pack(side='right')
        
        # Image preview
        self.image_preview_frame = tk.Frame(image_frame, bg='#ecf0f1', height=120)
        self.image_preview_frame.pack(fill='x', pady=5)
        self.image_preview_frame.pack_propagate(False)
        
        self.image_preview_label = tk.Label(
            self.image_preview_frame,
            text="No image selected\n(Will replace images with alt text like {{project_image}})",
            font=('Arial', 9),
            fg='#7f8c8d',
            bg='#ecf0f1'
        )
        self.image_preview_label.pack(expand=True)
        
        # Generate prompt button
        generate_btn = tk.Button(
            step2_frame,
            text="Generate AI Prompt",
            command=self.generate_prompt,
            bg='#e74c3c',
            fg='white',
            font=('Arial', 12, 'bold'),
            height=2
        )
        generate_btn.pack(pady=20)
        
        # AI Prompt display with copy button
        prompt_frame = tk.Frame(step2_frame)
        prompt_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        prompt_header = tk.Frame(prompt_frame)
        prompt_header.pack(fill='x')
        
        tk.Label(prompt_header, text="AI Prompt:", font=('Arial', 10, 'bold')).pack(side='left')
        
        # AI Service buttons in header
        button_container = tk.Frame(prompt_header)
        button_container.pack(side='right')
        
        self.copy_claude_btn = tk.Button(
            button_container,
            text="🧠 Copy & Open Claude",
            command=self.copy_and_open_claude,
            bg='#ff6b35',
            fg='white',
            font=('Arial', 9, 'bold'),
            state='disabled'
        )
        self.copy_claude_btn.pack(side='right', padx=2)
        
        self.copy_gemini_btn = tk.Button(
            button_container,
            text="✨ Copy & Open Gemini",
            command=self.copy_and_open_gemini,
            bg='#4285f4',
            fg='white',
            font=('Arial', 9, 'bold'),
            state='disabled'
        )
        self.copy_gemini_btn.pack(side='right', padx=2)
        
        self.copy_chatgpt_btn = tk.Button(
            button_container,
            text="💬 Copy & Open ChatGPT",
            command=self.copy_and_open_chatgpt,
            bg='#10a37f',
            fg='white',
            font=('Arial', 9, 'bold'),
            state='disabled'
        )
        self.copy_chatgpt_btn.pack(side='right', padx=2)
        
        self.prompt_text = scrolledtext.ScrolledText(
            prompt_frame, 
            height=12, 
            width=80,
            font=('Courier', 9),
            state='disabled'
        )
        self.prompt_text.pack(fill='both', expand=True, pady=5)
    
    def create_step3(self):
        """Step 3: AI response input and PowerPoint generation"""
        
        step3_frame = ttk.Frame(self.notebook)
        self.notebook.add(step3_frame, text="3. Generate PowerPoint")
        
        # Instructions
        instructions = tk.Label(
            step3_frame,
            text="Paste the AI's JSON response below, then generate your filled PowerPoint:",
            font=('Arial', 12),
            wraplength=700,
            justify='left'
        )
        instructions.pack(pady=20)
        
        # AI response input
        response_label = tk.Label(step3_frame, text="AI JSON Response:", font=('Arial', 10, 'bold'))
        response_label.pack(anchor='w', padx=20)
        
        self.response_text = scrolledtext.ScrolledText(
            step3_frame, 
            height=10, 
            width=80,
            font=('Courier', 9)
        )
        self.response_text.pack(pady=10, padx=20, fill='x')
        
        # Generate PowerPoint button
        generate_ppt_btn = tk.Button(
            step3_frame,
            text="Generate Filled PowerPoint",
            command=self.generate_powerpoint,
            bg='#f39c12',
            fg='white',
            font=('Arial', 12, 'bold'),
            height=2
        )
        generate_ppt_btn.pack(pady=20)
        
        # Results display
        results_label = tk.Label(step3_frame, text="Results:", font=('Arial', 10, 'bold'))
        results_label.pack(anchor='w', padx=20)
        
        self.results_text = scrolledtext.ScrolledText(
            step3_frame, 
            height=10, 
            width=80,
            font=('Courier', 9),
            state='disabled'
        )
        self.results_text.pack(pady=10, padx=20, fill='both', expand=True)
    
    def create_ai_buttons(self):
        """Create AI service buttons at the bottom"""
        
        # AI buttons frame
        ai_frame = tk.Frame(self.root, bg='#f0f0f0')
        ai_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=20, pady=10)
        
        # Title for AI buttons
        ai_label = tk.Label(
            ai_frame, 
            text="🤖 Quick Access to AI Services:", 
            font=('Arial', 10, 'bold'),
            bg='#f0f0f0'
        )
        ai_label.pack(anchor='w')
        
        # Buttons container
        buttons_frame = tk.Frame(ai_frame, bg='#f0f0f0')
        buttons_frame.pack(fill='x', pady=5)
        
        # Claude button
        claude_btn = tk.Button(
            buttons_frame,
            text="🧠 Open Claude",
            command=lambda: webbrowser.open("https://claude.ai"),
            bg='#ff6b35',
            fg='white',
            font=('Arial', 10, 'bold'),
            width=15,
            height=2,
            cursor='hand2'
        )
        claude_btn.pack(side='left', padx=5)
        
        # ChatGPT button
        chatgpt_btn = tk.Button(
            buttons_frame,
            text="💬 Open ChatGPT",
            command=lambda: webbrowser.open("https://chat.openai.com"),
            bg='#10a37f',
            fg='white',
            font=('Arial', 10, 'bold'),
            width=15,
            height=2,
            cursor='hand2'
        )
        chatgpt_btn.pack(side='left', padx=5)
        
        # Gemini button
        gemini_btn = tk.Button(
            buttons_frame,
            text="✨ Open Gemini",
            command=lambda: webbrowser.open("https://gemini.google.com"),
            bg='#4285f4',
            fg='white',
            font=('Arial', 10, 'bold'),
            width=15,
            height=2,
            cursor='hand2'
        )
        gemini_btn.pack(side='left', padx=5)
        
        # Help text
        help_label = tk.Label(
            ai_frame,
            text="💡 Tip: Use the 'Copy & Open' buttons above to automatically copy your prompt and open your preferred AI service",
            font=('Arial', 9),
            fg='#7f8c8d',
            bg='#f0f0f0'
        )
        help_label.pack(anchor='w', pady=2)
    
    def copy_and_open_claude(self):
        """Copy prompt to clipboard and open Claude"""
        if self.ai_prompt:
            pyperclip.copy(self.ai_prompt)
            webbrowser.open("https://claude.ai")
            self.notebook.select(2)  # Move to step 3
            self.status_var.set("Prompt copied and Claude opened - Paste and get your response")
        else:
            messagebox.showerror("Error", "No prompt to copy")
    
    def copy_and_open_chatgpt(self):
        """Copy prompt to clipboard and open ChatGPT"""
        if self.ai_prompt:
            pyperclip.copy(self.ai_prompt)
            webbrowser.open("https://chat.openai.com")
            self.notebook.select(2)  # Move to step 3
            self.status_var.set("Prompt copied and ChatGPT opened - Paste and get your response")
        else:
            messagebox.showerror("Error", "No prompt to copy")
    
    def copy_and_open_gemini(self):
        """Copy prompt to clipboard and open Gemini"""
        if self.ai_prompt:
            pyperclip.copy(self.ai_prompt)
            webbrowser.open("https://gemini.google.com")
            self.notebook.select(2)  # Move to step 3
            self.status_var.set("Prompt copied and Gemini opened - Paste and get your response")
        else:
            messagebox.showerror("Error", "No prompt to copy")
    
    def generate_unique_filename(self, base_path):
        """Generate a unique filename by adding modifiers if file exists"""
        
        # Split the path into directory, name, and extension
        directory = os.path.dirname(base_path)
        filename = os.path.basename(base_path)
        name, ext = os.path.splitext(filename)
        
        # If file doesn't exist, return original path
        if not os.path.exists(base_path):
            return base_path
        
        # Generate unique filename with modifiers
        counter = 1
        while True:
            # Try with timestamp first
            if counter == 1:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                new_name = f"{name}_{timestamp}{ext}"
            else:
                # Then try with numbers
                new_name = f"{name}_{counter:02d}{ext}"
            
            new_path = os.path.join(directory, new_name)
            
            if not os.path.exists(new_path):
                return new_path
            
            counter += 1
            
            # Safety valve to prevent infinite loop
            if counter > 100:
                # Use timestamp with milliseconds as fallback
                timestamp_ms = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
                new_name = f"{name}_{timestamp_ms}{ext}"
                new_path = os.path.join(directory, new_name)
                return new_path
    
    def browse_image(self):
        """Browse for image file"""
        file_path = filedialog.askopenfilename(
            title="Select Product Image",
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.gif *.bmp *.tiff"),
                ("PNG files", "*.png"),
                ("JPEG files", "*.jpg *.jpeg"),
                ("All files", "*.*")
            ]
        )
        if file_path:
            self.uploaded_image = file_path
            self.image_path_var.set(os.path.basename(file_path))
            self.show_image_preview(file_path)
            self.status_var.set(f"Image selected: {os.path.basename(file_path)}")
    
    def show_image_preview(self, image_path):
        """Show preview of selected image"""
        try:
            # Open and resize image for preview
            pil_image = Image.open(image_path)
            
            # Calculate preview size (maintain aspect ratio, max 150x100)
            max_width, max_height = 150, 100
            img_width, img_height = pil_image.size
            
            ratio = min(max_width/img_width, max_height/img_height)
            new_width = int(img_width * ratio)
            new_height = int(img_height * ratio)
            
            pil_image = pil_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
            
            # Convert to PhotoImage
            photo = ImageTk.PhotoImage(pil_image)
            
            # Update preview label
            self.image_preview_label.config(
                image=photo,
                text="",
                compound='center'
            )
            self.image_preview_label.image = photo  # Keep a reference
            
        except Exception as e:
            self.image_preview_label.config(
                image="",
                text=f"Preview error: {str(e)[:50]}...",
                compound='none'
            )
    
    def browse_file(self):
        """Browse for PowerPoint file"""
        file_path = filedialog.askopenfilename(
            title="Select PowerPoint Template",
            filetypes=[("PowerPoint files", "*.pptx"), ("All files", "*.*")]
        )
        if file_path:
            self.pptx_path.set(file_path)
            self.status_var.set(f"Selected: {os.path.basename(file_path)}")
    
    def analyze_fields(self):
        """Analyze PowerPoint for fields"""
        pptx_path_str = self.pptx_path.get()
        if not pptx_path_str:
            messagebox.showerror("Error", "Please select a PowerPoint file first")
            return
        
        if not os.path.exists(pptx_path_str):
            messagebox.showerror("Error", f"PowerPoint file not found at: {pptx_path_str}")
            return
        
        try:
            self.status_var.set("Analyzing PowerPoint fields...")
            self.root.update()
            
            # Clear previous results
            self.fields_text.config(state='normal')
            self.fields_text.delete(1.0, tk.END)
            
            # Analyze fields
            self.found_fields = self.analyze_powerpoint_fields(pptx_path_str)
            
            if self.found_fields:
                self.fields_text.insert(tk.END, f"✅ Found {len(self.found_fields)} unique placeholders:\n\n")
                for field in sorted(self.found_fields):
                    self.fields_text.insert(tk.END, f"  • {{{{{field}}}}}\n")
                
                self.fields_text.insert(tk.END, f"\n🎯 Ready for next step!")
                self.notebook.select(1)  # Move to step 2
                self.status_var.set(f"Found {len(self.found_fields)} fields - Ready for project data")
            else:
                self.fields_text.insert(tk.END, "❌ No {{field_name}} placeholders found.\n\n")
                self.fields_text.insert(tk.END, "Make sure your PowerPoint contains text like:\n")
                self.fields_text.insert(tk.END, "{{project_title}}, {{commander_name}}, etc.")
                self.status_var.set("No fields found - check your placeholders")
            
            self.fields_text.config(state='disabled')
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to analyze PowerPoint:\n{str(e)}")
            self.status_var.set("Analysis failed")
    
    def generate_prompt(self):
        """Generate AI prompt"""
        if not self.found_fields:
            messagebox.showerror("Error", "Please analyze PowerPoint fields first")
            return
        
        project_data = self.project_text.get(1.0, tk.END).strip()
        if not project_data:
            messagebox.showerror("Error", "Please enter project data")
            return
        
        try:
            self.status_var.set("Generating AI prompt...")
            
            # Generate prompt
            self.ai_prompt = self.generate_ai_prompt(self.found_fields, project_data)
            self.project_data = project_data
            
            # Display prompt
            self.prompt_text.config(state='normal')
            self.prompt_text.delete(1.0, tk.END)
            self.prompt_text.insert(tk.END, self.ai_prompt)
            self.prompt_text.config(state='disabled')
            
            # Enable AI service buttons
            self.copy_claude_btn.config(state='normal')
            self.copy_chatgpt_btn.config(state='normal')
            self.copy_gemini_btn.config(state='normal')
            
            self.status_var.set("AI prompt generated - Choose an AI service to copy and open")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate prompt:\n{str(e)}")
    
    def generate_powerpoint(self):
        """Generate filled PowerPoint"""
        if not self.pptx_path.get():
            messagebox.showerror("Error", "Please select a PowerPoint file first")
            return
        
        ai_response = self.response_text.get(1.0, tk.END).strip()
        if not ai_response:
            messagebox.showerror("Error", "Please paste the AI's JSON response")
            return
        
        try:
            self.status_var.set("Generating PowerPoint...")
            self.root.update()
            
            # Clear results
            self.results_text.config(state='normal')
            self.results_text.delete(1.0, tk.END)
            
            # Generate output filename with unique name handling
            input_file = self.pptx_path.get()
            base_output_file = input_file.replace('.pptx', '_filled.pptx')
            output_file = self.generate_unique_filename(base_output_file)
            
            # Show the filename that will be used
            output_filename = os.path.basename(output_file)
            if output_file != base_output_file:
                self.results_text.insert(tk.END, f"📝 Original filename exists, using: {output_filename}\n\n")
            else:
                self.results_text.insert(tk.END, f"📝 Output filename: {output_filename}\n\n")
            
            self.results_text.insert(tk.END, "🚀 Starting PowerPoint generation...\n\n")
            self.root.update()
            
            # Fill PowerPoint
            success = self.fill_powerpoint_with_data(input_file, ai_response, output_file)
            
            if success:
                self.results_text.insert(tk.END, f"\n✅ SUCCESS!\n")
                self.results_text.insert(tk.END, f"📄 Filled PowerPoint saved as:\n{os.path.basename(output_file)}\n")
                self.results_text.insert(tk.END, f"📁 Location: {os.path.dirname(output_file)}\n\n")
                self.results_text.insert(tk.END, f"🎉 Your presentation is ready!")
                self.status_var.set("PowerPoint generated successfully!")
                
                # Ask if user wants to open the file
                if messagebox.askyesno("Success", f"PowerPoint generated successfully!\n\nSaved as: {os.path.basename(output_file)}\n\nOpen the filled presentation now?"):
                    try:
                        import sys
                        if sys.platform == "win32":
                            os.startfile(output_file)
                        elif sys.platform == "darwin":
                            import subprocess
                            subprocess.run(['open', output_file])
                        else:
                            import subprocess
                            subprocess.run(['xdg-open', output_file])
                    except Exception as open_error:
                        messagebox.showinfo("File Saved", f"PowerPoint saved successfully but couldn't open automatically.\n\nLocation: {output_file}")
            else:
                self.results_text.insert(tk.END, "\n❌ FAILED\n")
                self.results_text.insert(tk.END, "Check the AI response format and try again.")
                self.status_var.set("PowerPoint generation failed")
            
            self.results_text.config(state='disabled')
            
        except Exception as e:
            self.results_text.config(state='normal')
            self.results_text.insert(tk.END, f"\n❌ ERROR: {str(e)}\n")
            self.results_text.config(state='disabled')
            messagebox.showerror("Error", f"Failed to generate PowerPoint:\n{str(e)}")
            self.status_var.set("Generation failed")
    
    # Analysis and processing functions
    def analyze_powerpoint_fields(self, pptx_path):
        """Find all {{field_name}} placeholders in PowerPoint"""
        try:
            prs = Presentation(pptx_path)
            found_fields = set()
            
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text_frame") and shape.text_frame:
                        text_content = shape.text_frame.text
                        field_pattern = r'\{\{([^}]+)\}\}'
                        matches = re.findall(field_pattern, text_content)
                        found_fields.update(matches)
                    
                    elif shape.has_table:
                        table = shape.table
                        for row in table.rows:
                            for cell in row.cells:
                                text_content = cell.text
                                field_pattern = r'\{\{([^}]+)\}\}'
                                matches = re.findall(field_pattern, text_content)
                                found_fields.update(matches)
            
            return list(found_fields)
            
        except Exception as e:
            raise Exception(f"Error analyzing PowerPoint: {str(e)}")
    
    def generate_ai_prompt(self, fields, project_data):
        """Generate AI prompt to extract field values from project data"""
        field_descriptions = []
        for field in sorted(fields):
            field_descriptions.append(f"  - {field}")
        
        prompt = f"""I need you to analyze project data and extract information for specific PowerPoint fields. Return ONLY a valid JSON object with the field names as keys and extracted values as values.

**PowerPoint Fields to Fill:**
{chr(10).join(field_descriptions)}

**Instructions:**
1. Extract relevant information from the project data for each field
2. If a field name suggests specific content (e.g., "commander_name" should be a person's name), extract accordingly
3. Keep values concise but informative - suitable for presentation slides
4. Conduct market research with a focus on Department of Defense, Department of the Air Force, and with the goals of the 100th ARW and 352nd SOW mission goals in mind.
5. For fields with money, phone numbers, or other implied formatting, format the extracted values accordingly
6. For fields you can't determine from the data, use "TBD" or leave reasonable placeholder text
7. Return ONLY the JSON object - no explanations or additional text

**Project Data to Analyze:**
{project_data}

Please analyze the above data and return the JSON object with field values"""

        return prompt
    
    def copy_run_formatting(self, source_run, target_run):
        """Copy all formatting from source run to target run with enhanced color handling"""
        try:
            # Basic font properties - use safe attribute access
            try:
                if hasattr(source_run.font, 'name') and source_run.font.name:
                    target_run.font.name = source_run.font.name
            except:
                pass
                
            try:
                if hasattr(source_run.font, 'size') and source_run.font.size:
                    target_run.font.size = source_run.font.size
            except:
                pass
                
            try:
                if hasattr(source_run.font, 'bold') and source_run.font.bold is not None:
                    target_run.font.bold = source_run.font.bold
            except:
                pass
                
            try:
                if hasattr(source_run.font, 'italic') and source_run.font.italic is not None:
                    target_run.font.italic = source_run.font.italic
            except:
                pass
                
            try:
                if hasattr(source_run.font, 'underline') and source_run.font.underline is not None:
                    target_run.font.underline = source_run.font.underline
            except:
                pass
            
            # Enhanced color handling - very defensive approach
            try:
                source_color = source_run.font.color
                target_color = target_run.font.color
                
                # Check what type of color this is
                color_type = str(type(source_color)).lower()
                
                if 'rgb' in color_type or hasattr(source_color, 'rgb'):
                    try:
                        if source_color.rgb is not None:
                            target_color.rgb = source_color.rgb
                    except:
                        pass
                
                elif 'scheme' in color_type or 'theme' in color_type or hasattr(source_color, 'theme_color'):
                    try:
                        if hasattr(source_color, 'theme_color') and source_color.theme_color is not None:
                            target_color.theme_color = source_color.theme_color
                            
                            # Also copy brightness/tint if it exists
                            if hasattr(source_color, 'brightness') and source_color.brightness is not None:
                                target_color.brightness = source_color.brightness
                    except:
                        pass
                
                else:
                    # Unknown color type - try to preserve it by copying the whole object
                    try:
                        if hasattr(source_color, '_element'):
                            target_color._element = source_color._element
                    except:
                        pass
                        
            except:
                pass
                
        except:
            pass
    
    def replace_in_existing_runs(self, paragraph, placeholder, replacement_text):
        """Replace text in existing runs without recreating them"""
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
    
    def replace_text_preserve_formatting(self, paragraph, placeholder, replacement_text):
        """Replace placeholder text while preserving ALL formatting including colors"""
        if placeholder not in paragraph.text:
            return False

        # Find ALL runs that contain any part of the placeholder
        placeholder_start = paragraph.text.find(placeholder)
        placeholder_end = placeholder_start + len(placeholder)
        
        current_pos = 0
        reference_run = None
        
        for run in paragraph.runs:
            run_start = current_pos
            run_end = current_pos + len(run.text)
            
            # If this run intersects with the placeholder
            if run_start < placeholder_end and run_end > placeholder_start:
                # Use the first intersecting run as the formatting reference
                if reference_run is None:
                    reference_run = run
            
            current_pos = run_end

        if not reference_run:
            return False

        # Store paragraph-level formatting
        original_alignment = paragraph.alignment
        original_level = paragraph.level

        # Replace the text
        new_text = paragraph.text.replace(placeholder, replacement_text)
        
        # Clear and rebuild the paragraph
        paragraph.clear()
        
        # Restore paragraph formatting
        try:
            paragraph.alignment = original_alignment
            paragraph.level = original_level
        except:
            pass
        
        # Add new run with preserved formatting
        new_run = paragraph.add_run()
        new_run.text = new_text
        
        # Copy all formatting from reference run
        self.copy_run_formatting(reference_run, new_run)

        return True

    def fill_powerpoint_with_data(self, template_path, json_data, output_path):
        """Fill PowerPoint template with JSON data"""
        try:
            if isinstance(json_data, str):
                try:
                    # Clean the response to get only the JSON part
                    json_str_match = re.search(r'\{.*\}', json_data, re.DOTALL)
                    if json_str_match:
                        json_str = json_str_match.group(0)
                        data = json.loads(json_str)
                    else:
                        raise json.JSONDecodeError("No JSON object found in the response.", "", 0)
                except json.JSONDecodeError as e:
                    raise Exception(f"Invalid JSON format: {e}. Please ensure the AI response is a valid JSON object.") from e
            else:
                data = json_data

            prs = Presentation(template_path)
            replacements_made = 0

            # Handle image replacement first and separately
            if self.uploaded_image:
                self.results_text.insert(tk.END, f"\n🖼️ Looking for image placeholders...\n")
                self.root.update()
                self.replace_image_placeholders(prs, self.uploaded_image)
            
            self.results_text.insert(tk.END, f"\n📊 Processing {len(data)} text fields...\n")
            self.root.update()

            for slide_num, slide in enumerate(prs.slides, 1):
                self.results_text.insert(tk.END, f"\n📄 Processing text on Slide {slide_num}...\n")
                self.root.update()
                
                for shape in slide.shapes:
                    if hasattr(shape, "text_frame") and shape.text_frame:
                        
                        # METHOD 1: Try in-place replacement first (best for preserving colors)
                        for para in shape.text_frame.paragraphs:
                            for field, value in data.items():
                                placeholder = f"{{{{{field}}}}}"
                                if placeholder in para.text:
                                    # Try the in-place method first
                                    success = self.replace_in_existing_runs(para, placeholder, str(value))
                                    if success:
                                        replacements_made += 1
                                        self.results_text.insert(tk.END, f"  ✅ In-place: {placeholder} -> '{str(value)[:30]}...'\n")
                                        self.root.update()
                                        continue
                                    
                                    # If in-place failed, try the formatting-preserving method
                                    success = self.replace_text_preserve_formatting(para, placeholder, str(value))
                                    if success:
                                        replacements_made += 1
                                        self.results_text.insert(tk.END, f"  ✅ Format-preserved: {placeholder} -> '{str(value)[:30]}...'\n")
                                        self.root.update()
                                        continue
                        
                        # METHOD 2: Fallback for any remaining placeholders
                        current_text = shape.text_frame.text
                        needs_fallback = any(f"{{{{{field}}}}}" in current_text for field in data.keys())
                        
                        if needs_fallback:
                            self.results_text.insert(tk.END, f"  🔄 Using fallback method for complex text...\n")
                            self.root.update()
                            
                            # Store original formatting from first run
                            first_run_format = None
                            if shape.text_frame.paragraphs and shape.text_frame.paragraphs[0].runs:
                                first_run_format = shape.text_frame.paragraphs[0].runs[0]
                            
                            # Replace all placeholders in the text
                            modified_text = current_text
                            for field, value in data.items():
                                placeholder = f"{{{{{field}}}}}"
                                if placeholder in modified_text:
                                    modified_text = modified_text.replace(placeholder, str(value))
                                    replacements_made += 1
                                    self.results_text.insert(tk.END, f"  ✅ Fallback: {placeholder} -> '{str(value)[:30]}...'\n")
                                    self.root.update()
                            
                            # Only recreate if text actually changed
                            if modified_text != current_text:
                                shape.text_frame.clear()
                                p = shape.text_frame.paragraphs[0]
                                run = p.add_run()
                                run.text = modified_text
                                
                                # Apply saved formatting
                                if first_run_format:
                                    self.copy_run_formatting(first_run_format, run)
                    
                    elif shape.has_table:
                        for row in shape.table.rows:
                            for cell in row.cells:
                                for para in cell.text_frame.paragraphs:
                                    for field, value in data.items():
                                        placeholder = f"{{{{{field}}}}}"
                                        if placeholder in para.text:
                                            # Try in-place first, then formatting-preserving
                                            success = self.replace_in_existing_runs(para, placeholder, str(value))
                                            if not success:
                                                success = self.replace_text_preserve_formatting(para, placeholder, str(value))
                                            
                                            if success:
                                                replacements_made += 1
                                                self.results_text.insert(tk.END, f"  ✅ Table: {placeholder}\n")
                                                self.root.update()

            prs.save(output_path)
            self.results_text.insert(tk.END, f"\n📈 Total text replacements made: {replacements_made}\n")
            return True

        except Exception as e:
            raise Exception(f"Error filling PowerPoint: {e}")

    def replace_image_placeholders(self, prs, image_path):
        """
        Find images with placeholder alt text and replace them with the uploaded image.
        This works by finding images with alt text like {{project_image}}, {{product_image}}, etc.
        """
        if not image_path or not os.path.exists(image_path):
            return 0
        
        replacements_made = 0
        
        # Common image placeholder patterns for alt text
        image_patterns = [
            r'\{\{.*image.*\}\}',
            r'\{\{.*photo.*\}\}',
            r'\{\{.*picture.*\}\}',
            r'\{\{.*graphic.*\}\}',
            r'\{\{.*logo.*\}\}'
        ]
        
        for slide_num, slide in enumerate(prs.slides, 1):
            shapes_to_replace = []  # Store shapes to replace (can't modify during iteration)
            
            for shape in slide.shapes:
                # Check if this is a picture shape with alt text
                if hasattr(shape, '_element') and shape._element.tag.endswith('}pic'):
                    try:
                        # Get the alt text (description) of the image
                        alt_text = ""
                        if hasattr(shape, 'element'):
                            # Try to get alt text from the shape
                            nvPicPr = shape.element.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}nvPicPr')
                            if nvPicPr is not None:
                                cNvPr = nvPicPr.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}cNvPr')
                                if cNvPr is not None:
                                    alt_text = cNvPr.get('descr', '') or cNvPr.get('title', '')
                        
                        # Also try the more direct approach
                        if not alt_text and hasattr(shape, 'name'):
                            alt_text = getattr(shape, 'name', '')
                            
                        # Also check if there's a description property
                        if not alt_text:
                            try:
                                if hasattr(shape, '_element') and hasattr(shape._element, 'get'):
                                    alt_text = shape._element.get('descr', '') or shape._element.get('title', '')
                            except:
                                pass
                        
                        self.results_text.insert(tk.END, f"    → Checking image alt text: '{alt_text}'\n")
                        self.root.update()
                        
                        # Check if any image pattern matches the alt text
                        if alt_text:
                            for pattern in image_patterns:
                                if re.search(pattern, alt_text, re.IGNORECASE):
                                    shapes_to_replace.append({
                                        'shape': shape,
                                        'alt_text': alt_text,
                                        'slide_num': slide_num
                                    })
                                    break
                                    
                    except Exception as e:
                        self.results_text.insert(tk.END, f"    → Error checking image alt text: {str(e)}\n")
                        continue
            
            # Now replace the identified shapes
            for shape_info in shapes_to_replace:
                try:
                    shape = shape_info['shape']
                    alt_text = shape_info['alt_text']
                    
                    # Get the position and size of the existing image
                    left = shape.left
                    top = shape.top
                    width = shape.width
                    height = shape.height
                    
                    # Remove the old image
                    shape_element = shape._element
                    shape_element.getparent().remove(shape_element)
                    
                    # Add the new image in the same position
                    new_picture = slide.shapes.add_picture(
                        image_path, 
                        left, 
                        top, 
                        width, 
                        height
                    )
                    
                    # Optionally, set the alt text of the new image
                    try:
                        if hasattr(new_picture, '_element'):
                            nvPicPr = new_picture._element.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}nvPicPr')
                            if nvPicPr is not None:
                                cNvPr = nvPicPr.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}cNvPr')
                                if cNvPr is not None:
                                    cNvPr.set('descr', f"Replaced: {alt_text}")
                    except:
                        pass  # Alt text setting is optional
                    
                    replacements_made += 1
                    self.results_text.insert(tk.END, f"  🖼️ Replaced image with alt text '{alt_text}' on Slide {slide_num}\n")
                    self.root.update()
                    
                except Exception as e:
                    self.results_text.insert(tk.END, f"  ⚠️ Image replacement failed on Slide {slide_num}: {str(e)}\n")
                    self.root.update()
        
        # Also try the original picture placeholder method as fallback
        try:
            from pptx.enum.shapes import PP_PLACEHOLDER
            
            for slide_num, slide in enumerate(prs.slides, 1):
                for shape in slide.placeholders:
                    # Check if the placeholder is a Picture placeholder
                    if shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
                        try:
                            # Insert the picture into the placeholder
                            picture = shape.insert_picture(image_path)
                            
                            replacements_made += 1
                            self.results_text.insert(tk.END, f"  🖼️ Inserted image into picture placeholder on Slide {slide_num}\n")
                            self.root.update()
                            
                        except Exception as e:
                            self.results_text.insert(tk.END, f"  ⚠️ Picture placeholder insertion failed on Slide {slide_num}: {str(e)}\n")
                            self.root.update()
        except:
            pass  # Picture placeholder method is optional fallback

        if replacements_made == 0:
            self.results_text.insert(tk.END, "  ⚠️ No images with placeholder alt text found.\n")
            self.results_text.insert(tk.END, "  💡 Tip: Add alt text like '{{project_image}}' to images in PowerPoint\n")
            self.results_text.insert(tk.END, "      (Right-click image → Edit Alt Text → Description)\n")
            self.root.update()
        else:
            self.results_text.insert(tk.END, f"  ✅ Successfully replaced {replacements_made} image(s)\n")
            
        return replacements_made


def main():
    # Check for required libraries and provide guidance if missing.
    try:
        import pyperclip
        from pptx import Presentation
        from PIL import Image, ImageTk
    except ImportError as e:
        # Create a simple Tkinter window to show the error
        root = tk.Tk()
        root.withdraw() # Hide the main window
        messagebox.showerror(
            "Missing Libraries",
            f"A required library is missing: {e.name}.\n\n"
            "Please install the necessary packages by running this command in your terminal:\n\n"
            "pip install pyperclip python-pptx Pillow"
        )
        return
    
    root = tk.Tk()
    app = PowerPointFillerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()

