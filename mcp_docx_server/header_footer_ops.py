"""
Header and footer operations for Word documents.
"""

from mcp_docx_server.utils import load_document, get_document_path, apply_paragraph_formatting

def add_header(doc_id: str, section_index: int, text: str = None, content: list = None) -> str:
    """Adds or modifies a header for a specific section.
    
    Args:
        doc_id (str): The document ID (filename without extension).
        section_index (int): The index of the section to add a header to (0-based).
        text (str, optional): Simple text to add to the header.
        content (list, optional): Complex content for the header, following the same format
                                 as document content in create_complete_document.
                                 
    Returns:
        str: A message indicating success or failure.
    """
    try:
        document = load_document(doc_id)
        
        if not document.sections or section_index >= len(document.sections):
            return f"Error: Section index {section_index} is out of range. Document has {len(document.sections) if document.sections else 0} sections."
        
        section = document.sections[section_index]
        header = section.header
        
        # Unlink from previous if it's currently linked
        if header.is_linked_to_previous:
            header.is_linked_to_previous = False
        
        # Clear existing content
        for paragraph in header.paragraphs[1:]:
            p = paragraph._element
            p.getparent().remove(p)
        
        # If first paragraph exists, use it, otherwise add one
        if header.paragraphs:
            first_paragraph = header.paragraphs[0]
            if text:
                first_paragraph.text = text
            else:
                first_paragraph.text = ""
        else:
            if text:
                header.paragraphs[0].text = text
        
        # If complex content is provided, add it
        if content:
            for item in content:
                content_type = item.get("type", "").lower()
                item_text = item.get("text", "")
                
                if content_type == "paragraph":
                    style = item.get("style", "Header")
                    
                    # Check if style exists
                    style_exists_in_doc = False
                    for doc_style in document.styles:
                        if doc_style.name == style:
                            style_exists_in_doc = True
                            break
                    
                    # If style doesn't exist, try to define it
                    if not style_exists_in_doc and style == "Header":
                        try:
                            # Add temporary paragraph to define the style
                            temp_para = document.add_paragraph("", style=style)
                            # Remove the temporary paragraph
                            p = temp_para._element
                            p.getparent().remove(p)
                        except KeyError:
                            style = None  # Style not found
                    
                    # Add the paragraph
                    para = header.add_paragraph(item_text)
                    if style:
                        try:
                            para.style = style
                        except:
                            pass  # Style not found, continue with default
                    
                    # Apply formatting if provided
                    formatting = item.get("formatting", {})
                    if formatting:
                        apply_paragraph_formatting(para, formatting)
                
                elif content_type == "table":
                    rows = item.get("rows", 1)
                    cols = item.get("cols", 1)
                    data = item.get("data", "")
                    
                    table = header.add_table(rows=rows, cols=cols)
                    
                    # Fill with data if provided
                    if data:
                        data_list = data.split(',')
                        
                        # Pad with empty strings if too few data elements
                        if len(data_list) < rows * cols:
                            data_list.extend([''] * (rows * cols - len(data_list)))
                            
                        # Fill table cells
                        for i in range(rows):
                            for j in range(cols):
                                cell_idx = i * cols + j
                                if cell_idx < len(data_list):
                                    table.cell(i, j).text = data_list[cell_idx].strip()
        
        doc_path = get_document_path(doc_id)
        document.save(doc_path)
        return f"Header added/modified for section {section_index}."
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error adding header: {str(e)}"

def add_footer(doc_id: str, section_index: int, text: str = None, content: list = None) -> str:
    """Adds or modifies a footer for a specific section.
    
    Args:
        doc_id (str): The document ID (filename without extension).
        section_index (int): The index of the section to add a footer to (0-based).
        text (str, optional): Simple text to add to the footer.
        content (list, optional): Complex content for the footer, following the same format
                                 as document content in create_complete_document.
                                 
    Returns:
        str: A message indicating success or failure.
    """
    try:
        document = load_document(doc_id)
        
        if not document.sections or section_index >= len(document.sections):
            return f"Error: Section index {section_index} is out of range. Document has {len(document.sections) if document.sections else 0} sections."
        
        section = document.sections[section_index]
        footer = section.footer
        
        # Unlink from previous if it's currently linked
        if footer.is_linked_to_previous:
            footer.is_linked_to_previous = False
        
        # Clear existing content
        for paragraph in footer.paragraphs[1:]:
            p = paragraph._element
            p.getparent().remove(p)
        
        # If first paragraph exists, use it, otherwise add one
        if footer.paragraphs:
            first_paragraph = footer.paragraphs[0]
            if text:
                first_paragraph.text = text
            else:
                first_paragraph.text = ""
        else:
            if text:
                footer.paragraphs[0].text = text
        
        # If complex content is provided, add it
        if content:
            for item in content:
                content_type = item.get("type", "").lower()
                item_text = item.get("text", "")
                
                if content_type == "paragraph":
                    style = item.get("style", "Footer")
                    
                    # Check if style exists
                    style_exists_in_doc = False
                    for doc_style in document.styles:
                        if doc_style.name == style:
                            style_exists_in_doc = True
                            break
                    
                    # If style doesn't exist, try to define it
                    if not style_exists_in_doc and style == "Footer":
                        try:
                            # Add temporary paragraph to define the style
                            temp_para = document.add_paragraph("", style=style)
                            # Remove the temporary paragraph
                            p = temp_para._element
                            p.getparent().remove(p)
                        except KeyError:
                            style = None  # Style not found
                    
                    # Add the paragraph
                    para = footer.add_paragraph(item_text)
                    if style:
                        try:
                            para.style = style
                        except:
                            pass  # Style not found, continue with default
                    
                    # Apply formatting if provided
                    formatting = item.get("formatting", {})
                    if formatting:
                        apply_paragraph_formatting(para, formatting)
                
                elif content_type == "table":
                    rows = item.get("rows", 1)
                    cols = item.get("cols", 1)
                    data = item.get("data", "")
                    
                    table = footer.add_table(rows=rows, cols=cols)
                    
                    # Fill with data if provided
                    if data:
                        data_list = data.split(',')
                        
                        # Pad with empty strings if too few data elements
                        if len(data_list) < rows * cols:
                            data_list.extend([''] * (rows * cols - len(data_list)))
                            
                        # Fill table cells
                        for i in range(rows):
                            for j in range(cols):
                                cell_idx = i * cols + j
                                if cell_idx < len(data_list):
                                    table.cell(i, j).text = data_list[cell_idx].strip()
        
        doc_path = get_document_path(doc_id)
        document.save(doc_path)
        return f"Footer added/modified for section {section_index}."
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error adding footer: {str(e)}"

def add_zoned_header(doc_id: str, section_index: int, left_text: str = "", center_text: str = "", right_text: str = "") -> str:
    """Adds a three-zone header with left, center, and right aligned text.
    
    This is a convenience function that creates a properly formatted header
    with text aligned in the left, center, and right zones using tab stops.
    
    Args:
        doc_id (str): The document ID (filename without extension).
        section_index (int): The index of the section to add a header to (0-based).
        left_text (str): Text for the left zone.
        center_text (str): Text for the center zone.
        right_text (str): Text for the right zone.
        
    Returns:
        str: A message indicating success or failure.
    """
    try:
        document = load_document(doc_id)
        
        if not document.sections or section_index >= len(document.sections):
            return f"Error: Section index {section_index} is out of range. Document has {len(document.sections) if document.sections else 0} sections."
        
        section = document.sections[section_index]
        header = section.header
        
        # Unlink from previous if it's currently linked
        if header.is_linked_to_previous:
            header.is_linked_to_previous = False
        
        # Create the zoned header with tab-separated text
        header_text = f"{left_text}"
        if center_text:
            header_text += f"\t{center_text}"
            if right_text:
                header_text += f"\t{right_text}"
        elif right_text:
            header_text += f"\t\t{right_text}"
        
        # Clear existing content
        for paragraph in header.paragraphs[1:]:
            p = paragraph._element
            p.getparent().remove(p)
        
        # Check if "Header" style exists and define it if needed
        header_style_exists = False
        for style in document.styles:
            if style.name == "Header":
                header_style_exists = True
                break
        
        if not header_style_exists:
            try:
                # Add temporary paragraph to define the style
                temp_para = document.add_paragraph("", style="Header")
                # Remove the temporary paragraph
                p = temp_para._element
                p.getparent().remove(p)
            except KeyError:
                pass  # Style not found, continue with default style
        
        # Apply the text to the first paragraph
        if header.paragraphs:
            paragraph = header.paragraphs[0]
            paragraph.text = header_text
            try:
                paragraph.style = document.styles["Header"]
            except:
                pass  # Style not found, continue with default style
        else:
            paragraph = header.add_paragraph(header_text)
            try:
                paragraph.style = document.styles["Header"]
            except:
                pass
        
        doc_path = get_document_path(doc_id)
        document.save(doc_path)
        return f"Zoned header added for section {section_index}."
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error adding zoned header: {str(e)}"

def add_zoned_footer(doc_id: str, section_index: int, left_text: str = "", center_text: str = "", right_text: str = "") -> str:
    """Adds a three-zone footer with left, center, and right aligned text.
    
    This is a convenience function that creates a properly formatted footer
    with text aligned in the left, center, and right zones using tab stops.
    
    Args:
        doc_id (str): The document ID (filename without extension).
        section_index (int): The index of the section to add a footer to (0-based).
        left_text (str): Text for the left zone.
        center_text (str): Text for the center zone.
        right_text (str): Text for the right zone.
        
    Returns:
        str: A message indicating success or failure.
    """
    try:
        document = load_document(doc_id)
        
        if not document.sections or section_index >= len(document.sections):
            return f"Error: Section index {section_index} is out of range. Document has {len(document.sections) if document.sections else 0} sections."
        
        section = document.sections[section_index]
        footer = section.footer
        
        # Unlink from previous if it's currently linked
        if footer.is_linked_to_previous:
            footer.is_linked_to_previous = False
        
        # Create the zoned footer with tab-separated text
        footer_text = f"{left_text}"
        if center_text:
            footer_text += f"\t{center_text}"
            if right_text:
                footer_text += f"\t{right_text}"
        elif right_text:
            footer_text += f"\t\t{right_text}"
        
        # Clear existing content
        for paragraph in footer.paragraphs[1:]:
            p = paragraph._element
            p.getparent().remove(p)
        
        # Check if "Footer" style exists and define it if needed
        footer_style_exists = False
        for style in document.styles:
            if style.name == "Footer":
                footer_style_exists = True
                break
        
        if not footer_style_exists:
            try:
                # Add temporary paragraph to define the style
                temp_para = document.add_paragraph("", style="Footer")
                # Remove the temporary paragraph
                p = temp_para._element
                p.getparent().remove(p)
            except KeyError:
                pass  # Style not found, continue with default style
        
        # Apply the text to the first paragraph
        if footer.paragraphs:
            paragraph = footer.paragraphs[0]
            paragraph.text = footer_text
            try:
                paragraph.style = document.styles["Footer"]
            except:
                pass  # Style not found, continue with default style
        else:
            paragraph = footer.add_paragraph(footer_text)
            try:
                paragraph.style = document.styles["Footer"]
            except:
                pass
        
        doc_path = get_document_path(doc_id)
        document.save(doc_path)
        return f"Zoned footer added for section {section_index}."
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error adding zoned footer: {str(e)}"

def remove_header(doc_id: str, section_index: int) -> str:
    """Removes the header from a specific section, linking it to the previous section.
    
    Args:
        doc_id (str): The document ID (filename without extension).
        section_index (int): The index of the section to remove the header from (0-based).
        
    Returns:
        str: A message indicating success or failure.
    """
    try:
        document = load_document(doc_id)
        
        if not document.sections or section_index >= len(document.sections):
            return f"Error: Section index {section_index} is out of range. Document has {len(document.sections) if document.sections else 0} sections."
        
        section = document.sections[section_index]
        header = section.header
        
        # Link to previous, which removes this header definition
        header.is_linked_to_previous = True
        
        doc_path = get_document_path(doc_id)
        document.save(doc_path)
        return f"Header removed from section {section_index}."
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error removing header: {str(e)}"

def remove_footer(doc_id: str, section_index: int) -> str:
    """Removes the footer from a specific section, linking it to the previous section.
    
    Args:
        doc_id (str): The document ID (filename without extension).
        section_index (int): The index of the section to remove the footer from (0-based).
        
    Returns:
        str: A message indicating success or failure.
    """
    try:
        document = load_document(doc_id)
        
        if not document.sections or section_index >= len(document.sections):
            return f"Error: Section index {section_index} is out of range. Document has {len(document.sections) if document.sections else 0} sections."
        
        section = document.sections[section_index]
        footer = section.footer
        
        # Link to previous, which removes this footer definition
        footer.is_linked_to_previous = True
        
        doc_path = get_document_path(doc_id)
        document.save(doc_path)
        return f"Footer removed from section {section_index}."
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error removing footer: {str(e)}"

def get_header_text(doc_id: str, section_index: int) -> str:
    """Gets the text content of a header for a specific section.
    
    Args:
        doc_id (str): The document ID (filename without extension).
        section_index (int): The index of the section (0-based).
        
    Returns:
        str: The text content of the header or status message.
    """
    try:
        document = load_document(doc_id)
        
        if not document.sections or section_index >= len(document.sections):
            return f"Error: Section index {section_index} is out of range. Document has {len(document.sections) if document.sections else 0} sections."
        
        section = document.sections[section_index]
        header = section.header
        
        if header.is_linked_to_previous:
            # Find the first previous section with a header definition
            linked_section_index = section_index
            while linked_section_index > 0:
                linked_section_index -= 1
                prev_header = document.sections[linked_section_index].header
                if not prev_header.is_linked_to_previous:
                    return f"Header is linked to section {linked_section_index}. Content: {get_header_text(doc_id, linked_section_index)}"
            
            return "No header defined for this section (linked to previous, but no previous header found)."
        
        # Header has its own definition, extract the text
        header_text = []
        for paragraph in header.paragraphs:
            header_text.append(paragraph.text)
        
        if not header_text:
            return "Header is defined but contains no text."
        
        return "\n".join(header_text)
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error getting header text: {str(e)}"

def get_footer_text(doc_id: str, section_index: int) -> str:
    """Gets the text content of a footer for a specific section.
    
    Args:
        doc_id (str): The document ID (filename without extension).
        section_index (int): The index of the section (0-based).
        
    Returns:
        str: The text content of the footer or status message.
    """
    try:
        document = load_document(doc_id)
        
        if not document.sections or section_index >= len(document.sections):
            return f"Error: Section index {section_index} is out of range. Document has {len(document.sections) if document.sections else 0} sections."
        
        section = document.sections[section_index]
        footer = section.footer
        
        if footer.is_linked_to_previous:
            # Find the first previous section with a footer definition
            linked_section_index = section_index
            while linked_section_index > 0:
                linked_section_index -= 1
                prev_footer = document.sections[linked_section_index].footer
                if not prev_footer.is_linked_to_previous:
                    return f"Footer is linked to section {linked_section_index}. Content: {get_footer_text(doc_id, linked_section_index)}"
            
            return "No footer defined for this section (linked to previous, but no previous footer found)."
        
        # Footer has its own definition, extract the text
        footer_text = []
        for paragraph in footer.paragraphs:
            footer_text.append(paragraph.text)
        
        if not footer_text:
            return "Footer is defined but contains no text."
        
        return "\n".join(footer_text)
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error getting footer text: {str(e)}"
