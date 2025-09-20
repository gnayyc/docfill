"""DOCX template processor for replacing placeholders with values."""

import re
from pathlib import Path
from typing import Dict, Union
from docx import Document


class DocxProcessor:
    """Process DOCX templates by replacing {{placeholder}} with actual values."""
    
    def __init__(self, template_path: Union[str, Path]):
        """Initialize with template file path."""
        self.template_path = Path(template_path)
        if not self.template_path.exists():
            raise FileNotFoundError(f"Template file not found: {template_path}")
    
    def fill_template(self, data: Dict[str, str], output_path: Union[str, Path], verbose: bool = False) -> None:
        """Fill template with data and save to output path."""
        try:
            if verbose:
                print(f"Opening template: {self.template_path}")
            doc = Document(self.template_path)
        except Exception as e:
            raise RuntimeError(f"Failed to open template '{self.template_path}': {e}")
        
        # Process paragraphs
        for paragraph in doc.paragraphs:
            self._replace_in_paragraph(paragraph, data, verbose)

        # Process tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        self._replace_in_paragraph(paragraph, data, verbose)

        # Process headers and footers
        for section in doc.sections:
            # Process header
            if section.header:
                for paragraph in section.header.paragraphs:
                    self._replace_in_paragraph(paragraph, data, verbose)

            # Process footer
            if section.footer:
                for paragraph in section.footer.paragraphs:
                    self._replace_in_paragraph(paragraph, data, verbose)
        
        # Save the document
        try:
            if verbose:
                print(f"Saving to: {output_path}")
            doc.save(output_path)
            if verbose:
                print(f"Successfully saved: {output_path}")
        except Exception as e:
            raise RuntimeError(f"Failed to save document to '{output_path}': {e}")
    
    def _replace_in_paragraph(self, paragraph, data: Dict[str, str], verbose: bool = False) -> None:
        """Replace placeholders in a paragraph while preserving formatting."""
        # First, try to handle the entire paragraph text in case placeholders span multiple runs
        full_text = paragraph.text
        placeholders = re.findall(r'\{\{([^}]+)\}\}', full_text)
        
        if not placeholders:
            return
        
        # Check if we can handle simple cases (when placeholder is within a single run)
        simple_replacement = True
        for run in paragraph.runs:
            run_placeholders = re.findall(r'\{\{([^}]+)\}\}', run.text)
            if run_placeholders:
                # This run contains complete placeholders, process it
                original_text = run.text
                new_text = original_text
                
                for placeholder in run_placeholders:
                    placeholder_with_braces = f"{{{{{placeholder}}}}}"
                    
                    # Look for exact match first
                    if placeholder in data:
                        replacement = data[placeholder]
                    # Look for case-insensitive match
                    elif placeholder.lower() in [k.lower() for k in data.keys()]:
                        # Find the actual key with case-insensitive match
                        actual_key = next(k for k in data.keys() if k.lower() == placeholder.lower())
                        replacement = data[actual_key]
                    else:
                        # Keep the placeholder if no replacement found
                        replacement = placeholder_with_braces
                        print(f"Warning: No replacement found for placeholder: {placeholder}")
                    
                    new_text = new_text.replace(placeholder_with_braces, replacement)
                
                # Update the run text (this preserves the run's formatting)
                if new_text != original_text:
                    run.text = new_text
        
        # Handle complex cases where placeholders span multiple runs
        # Check if we still have placeholders after simple replacement
        remaining_text = paragraph.text
        remaining_placeholders = re.findall(r'\{\{([^}]+)\}\}', remaining_text)
        
        if remaining_placeholders:
            # Advanced approach: reconstruct runs while preserving formatting
            self._replace_across_runs(paragraph, remaining_placeholders, data, verbose)
    
    def _replace_across_runs(self, paragraph, placeholders, data, verbose=False):
        """Replace placeholders that span across multiple runs while preserving formatting."""
        from docx.oxml.shared import qn
        from docx.shared import RGBColor
        
        # Build a character-level map of the paragraph
        char_map = []  # List of (char, run_index, position_in_run)
        run_formats = []  # Store formatting info for each run
        
        for run_idx, run in enumerate(paragraph.runs):
            # Store run formatting
            run_formats.append({
                'bold': run.bold,
                'italic': run.italic,
                'underline': run.underline,
                'font_name': run.font.name,
                'font_size': run.font.size,
                'font_color': run.font.color.rgb if run.font.color.rgb else None,
            })
            
            # Map each character to its run
            for char_idx, char in enumerate(run.text):
                char_map.append((char, run_idx, char_idx))
        
        # Build the full text and find placeholder positions
        full_text = ''.join([item[0] for item in char_map])
        
        # Replace placeholders in the full text
        new_text = full_text
        for placeholder in placeholders:
            placeholder_with_braces = f"{{{{{placeholder}}}}}"
            
            if placeholder in data:
                replacement = data[placeholder]
            elif placeholder.lower() in [k.lower() for k in data.keys()]:
                actual_key = next(k for k in data.keys() if k.lower() == placeholder.lower())
                replacement = data[actual_key]
            else:
                replacement = placeholder_with_braces
                print(f"Warning: No replacement found for placeholder: {placeholder}")
            
            new_text = new_text.replace(placeholder_with_braces, replacement)
        
        if new_text != full_text:
            # Find which runs are affected and rebuild only those
            # For now, use a simpler approach: preserve the first run's formatting
            first_run_format = run_formats[0] if run_formats else {}
            
            # Clear and rebuild with preserved formatting
            paragraph.clear()
            new_run = paragraph.add_run(new_text)
            
            # Apply the first run's formatting to the new run
            if first_run_format.get('bold') is not None:
                new_run.bold = first_run_format['bold']
            if first_run_format.get('italic') is not None:
                new_run.italic = first_run_format['italic']
            if first_run_format.get('underline') is not None:
                new_run.underline = first_run_format['underline']
            if first_run_format.get('font_name'):
                new_run.font.name = first_run_format['font_name']
            if first_run_format.get('font_size'):
                new_run.font.size = first_run_format['font_size']
            if first_run_format.get('font_color'):
                new_run.font.color.rgb = first_run_format['font_color']
            
            if verbose:
                print(f"Info: Replaced text across multiple runs, applied first run's formatting")
    
    def get_placeholders(self) -> set:
        """Extract all placeholders from the template."""
        doc = Document(self.template_path)
        placeholders = set()
        
        # Get placeholders from paragraphs
        for paragraph in doc.paragraphs:
            placeholders.update(re.findall(r'\{\{([^}]+)\}\}', paragraph.text))
        
        # Get placeholders from tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        placeholders.update(re.findall(r'\{\{([^}]+)\}\}', paragraph.text))
        
        # Get placeholders from headers and footers
        for section in doc.sections:
            if section.header:
                for paragraph in section.header.paragraphs:
                    placeholders.update(re.findall(r'\{\{([^}]+)\}\}', paragraph.text))
            
            if section.footer:
                for paragraph in section.footer.paragraphs:
                    placeholders.update(re.findall(r'\{\{([^}]+)\}\}', paragraph.text))
        
        return placeholders