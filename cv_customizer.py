import json
import re
from pathlib import Path
from typing import Dict, List, Union
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import subprocess
import sys
from datetime import datetime
import os

class CVCustomizer:
    def __init__(self, template_path: str):
        """initialize with a word template containing placeholders"""
        self.template_path = template_path
        self.document = Document(template_path)
    
    def replace_placeholders(self, replacements: Dict[str, Union[str, List[str]]]):
        """replace placeholders throughout the document while preserving formatting"""
        # handle paragraphs
        for paragraph in self.document.paragraphs:
            self._replace_in_paragraph(paragraph, replacements)
        
        # handle tables
        for table in self.document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        self._replace_in_paragraph(paragraph, replacements)
        
        # handle headers and footers
        for section in self.document.sections:
            # header
            for paragraph in section.header.paragraphs:
                self._replace_in_paragraph(paragraph, replacements)
            # footer
            for paragraph in section.footer.paragraphs:
                self._replace_in_paragraph(paragraph, replacements)
    
    def _replace_in_paragraph(self, paragraph, replacements):
        """replace placeholders in a paragraph while preserving run formatting"""
        # get full paragraph text
        full_text = paragraph.text
        
        # Flatten nested dict structure for easier replacement
        flat_replacements = {}
        for key, value in replacements.items():
            if isinstance(value, dict):
                # Handle nested dict like "t": {"skills": "...", "bp1": "..."}
                for subkey, subvalue in value.items():
                    flat_replacements[f"{key}.{subkey}"] = subvalue
            else:
                flat_replacements[key] = value
        
        # check if any placeholders exist in this paragraph
        has_placeholder = any(f"{{{{{key}}}}}" in full_text for key in flat_replacements)
        if not has_placeholder:
            return
        
        # process each placeholder
        for key, value in flat_replacements.items():
            placeholder = f"{{{{{key}}}}}"
            
            if placeholder in full_text:
                if isinstance(value, list):
                    # for lists, create bullet points
                    value_text = '\n• '.join(value)
                    value_text = '• ' + value_text if value else ''
                else:
                    # Ensure single line for bullet points to maintain formatting
                    value_text = str(value).replace('\n', ' ').strip()
                
                # if it's a simple replacement (placeholder is complete in one run)
                for run in paragraph.runs:
                    if placeholder in run.text:
                        # preserve formatting
                        run.text = run.text.replace(placeholder, value_text)
                        return
                
                # complex case: placeholder spans multiple runs
                # rebuild the paragraph
                self._complex_replace(paragraph, placeholder, value_text)
    
    def _complex_replace(self, paragraph, placeholder, replacement):
        """handle placeholders that span multiple runs"""
        # store run properties
        run_props = []
        combined_text = ""
        
        for run in paragraph.runs:
            run_props.append({
                'bold': run.bold,
                'italic': run.italic,
                'underline': run.underline,
                'font_name': run.font.name,
                'font_size': run.font.size,
                'color': run.font.color.rgb if run.font.color.rgb else None
            })
            combined_text += run.text
        
        # replace in combined text
        new_text = combined_text.replace(placeholder, replacement)
        
        # clear runs and recreate with original formatting
        for run in paragraph.runs:
            run.text = ""
        
        # add the new text with the first run's formatting
        if paragraph.runs:
            paragraph.runs[0].text = new_text
            # apply the original formatting from the first run
            if run_props:
                props = run_props[0]
                paragraph.runs[0].bold = props['bold']
                paragraph.runs[0].italic = props['italic']
                paragraph.runs[0].underline = props['underline']
                if props['font_name']:
                    paragraph.runs[0].font.name = props['font_name']
                if props['font_size']:
                    paragraph.runs[0].font.size = props['font_size']
    
    def save_docx(self, output_path: str):
        """save the modified document as a word file"""
        self.document.save(output_path)
    
    def convert_to_pdf(self, docx_path: str, pdf_path: str):
        """convert word document to pdf using libreoffice (cross-platform)"""
        try:
            # using libreoffice in headless mode
            subprocess.run([
                'soffice',
                '--headless',
                '--convert-to',
                'pdf',
                '--outdir',
                str(Path(pdf_path).parent),
                docx_path
            ], check=True)
            
            # rename to desired output name if needed
            generated_pdf = Path(docx_path).with_suffix('.pdf')
            if generated_pdf.name != Path(pdf_path).name:
                generated_pdf.rename(pdf_path)
                
        except subprocess.CalledProcessError:
            print("libreoffice conversion failed. trying python-docx2pdf...")
            # fallback to python-docx2pdf (windows only)
            try:
                from docx2pdf import convert
                convert(docx_path, pdf_path)
            except ImportError:
                print("please install libreoffice or docx2pdf for pdf conversion")
                raise
    
    def customize_cv(self, job_data: Dict, output_name: str):
        """main method to customize cv with job data"""
        # Validate content length to ensure 1-page format
        self._validate_content_length(job_data)
        
        # Debug: Print the data structure being passed
        print("🔧 CV data structure:")
        for key, value in job_data.items():
            if isinstance(value, dict):
                print(f"  {key}:")
                for subkey, subvalue in value.items():
                    print(f"    {subkey}: {str(subvalue)[:60]}...")
            else:
                print(f"  {key}: {str(value)[:60]}...")
        
        # create execution folder with today's date and time
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        execution_folder = f"outputs/{timestamp}_{output_name}"
        
        # create the execution folder
        os.makedirs(execution_folder, exist_ok=True)
        print(f"📁 Created execution folder: {execution_folder}")
        
        # replace all placeholders
        self.replace_placeholders(job_data)
        
        # save as word document in execution folder
        docx_output = f"{execution_folder}/{output_name}.docx"
        self.save_docx(docx_output)
        print(f"✓ Word document saved: {docx_output}")
        
        # convert to pdf in execution folder
        pdf_output = f"{execution_folder}/{output_name}.pdf"
        try:
            self.convert_to_pdf(docx_output, pdf_output)
            print(f"✓ PDF generated: {pdf_output}")
        except Exception as e:
            print(f"⚠️  PDF conversion failed: {e}")
            print("   Word document is still available")
            pdf_output = None
        
        return docx_output, pdf_output
    
    def _validate_content_length(self, job_data: Dict):
        """Validate that content will fit on one page"""
        warnings = []
        
        # Check bio length
        bio = job_data.get('bio', '')
        if len(bio) > 400:
            warnings.append(f"Bio is too long ({len(bio)} chars, recommend <400)")
        
        # Check bullet points
        for section in ['t', 'a']:
            if section in job_data:
                for i in range(1, 5):  # bp1-bp4 for t, bp1-bp3 for a
                    bp_key = f'bp{i}'
                    if bp_key in job_data[section]:
                        bp_text = job_data[section][bp_key]
                        if len(bp_text) > 120:
                            warnings.append(f"{section}.{bp_key} is too long ({len(bp_text)} chars, recommend <120)")
        
        if warnings:
            print("⚠️  Content length warnings:")
            for warning in warnings:
                print(f"   - {warning}")
            print("   CV may exceed one page!")
        else:
            print("✓ Content length validation passed")