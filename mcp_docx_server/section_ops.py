"""
Section operations for Word documents.
"""

from docx.enum.section import WD_SECTION, WD_ORIENT
from mcp_docx_server.utils import load_document, get_document_path

def add_section(doc_id: str, start_type: str = "NEW_PAGE") -> str:
    """Adds a new section to the end of a document.
    
    Args:
        doc_id (str): The document ID (filename without extension).
        start_type (str): The type of section break, one of:
            - "NEW_PAGE" (default) - Start the section on a new page
            - "EVEN_PAGE" - Start the section on the next even-numbered page
            - "ODD_PAGE" - Start the section on the next odd-numbered page
            - "CONTINUOUS" - No page break, continue on same page
    
    Returns:
        str: A message indicating success or failure.
    """
    try:
        document = load_document(doc_id)
        
        # Map string values to WD_SECTION enum
        section_types = {
            "NEW_PAGE": WD_SECTION.NEW_PAGE,
            "EVEN_PAGE": WD_SECTION.EVEN_PAGE, 
            "ODD_PAGE": WD_SECTION.ODD_PAGE,
            "CONTINUOUS": WD_SECTION.CONTINUOUS
        }
        
        # Get the section type value
        section_type = section_types.get(start_type.upper())
        if not section_type:
            return f"Error: Invalid section start type '{start_type}'. Valid values are: {', '.join(section_types.keys())}"
        
        # Add the new section
        document.add_section(section_type)
        
        doc_path = get_document_path(doc_id)
        document.save(doc_path)
        return f"Section with start type '{start_type}' added successfully."
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error adding section: {str(e)}"

def list_sections(doc_id: str) -> str:
    """Lists all sections in a document with their properties.
    
    Args:
        doc_id (str): The document ID (filename without extension).
    
    Returns:
        str: Information about each section in the document.
    """
    try:
        document = load_document(doc_id)
        
        if not document.sections:
            return f"No sections found in document '{doc_id}.docx'."
        
        sections_info = []
        for i, section in enumerate(document.sections):
            # Map orientation value to readable string
            orientation = "PORTRAIT" if section.orientation == WD_ORIENT.PORTRAIT else "LANDSCAPE"
            
            # Convert dimensions to inches for readability
            page_width_inches = section.page_width / 914400  # 914400 = 1 inch in EMUs
            page_height_inches = section.page_height / 914400
            
            # Create info string
            section_info = [
                f"Section {i}:",
                f"  Start Type: {section.start_type}",
                f"  Orientation: {orientation}",
                f"  Page Size: {page_width_inches:.2f}\" x {page_height_inches:.2f}\"",
                f"  Margins (inches):",
                f"    Left: {section.left_margin/914400:.2f}\"",
                f"    Right: {section.right_margin/914400:.2f}\"",
                f"    Top: {section.top_margin/914400:.2f}\"",
                f"    Bottom: {section.bottom_margin/914400:.2f}\"",
                f"    Gutter: {section.gutter/914400:.2f}\"",
                f"    Header Distance: {section.header_distance/914400:.2f}\"",
                f"    Footer Distance: {section.footer_distance/914400:.2f}\""
            ]
            sections_info.append("\n".join(section_info))
        
        return "\n\n".join(sections_info)
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error listing sections: {str(e)}"

def set_section_properties(doc_id: str, section_index: int, properties: dict) -> str:
    """Sets properties for a specific section in the document.
    
    Args:
        doc_id (str): The document ID (filename without extension).
        section_index (int): The index of the section to modify (0-based).
        properties (dict): Dictionary with section properties:
            - start_type: Section break type ("NEW_PAGE", "EVEN_PAGE", "ODD_PAGE", "CONTINUOUS")
            - orientation: "PORTRAIT" or "LANDSCAPE"
            - page_width: Page width in inches
            - page_height: Page height in inches
            - left_margin, right_margin, top_margin, bottom_margin: Margins in inches
            - gutter: Gutter margin in inches
            - header_distance, footer_distance: Header/footer distance in inches
    
    Returns:
        str: A message indicating success or failure.
    """
    try:
        document = load_document(doc_id)
        
        if not document.sections or section_index >= len(document.sections):
            return f"Error: Section index {section_index} is out of range. Document has {len(document.sections) if document.sections else 0} sections."
        
        section = document.sections[section_index]
        
        # Handle start_type
        if "start_type" in properties:
            start_type = properties["start_type"].upper()
            section_types = {
                "NEW_PAGE": WD_SECTION.NEW_PAGE,
                "EVEN_PAGE": WD_SECTION.EVEN_PAGE, 
                "ODD_PAGE": WD_SECTION.ODD_PAGE,
                "CONTINUOUS": WD_SECTION.CONTINUOUS
            }
            if start_type in section_types:
                section.start_type = section_types[start_type]
            else:
                return f"Error: Invalid section start type '{start_type}'. Valid values are: {', '.join(section_types.keys())}"
        
        # Handle orientation
        if "orientation" in properties:
            orientation = properties["orientation"].upper()
            if orientation == "LANDSCAPE":
                # If changing to landscape, may need to swap width and height
                if section.orientation == WD_ORIENT.PORTRAIT:
                    # Store current dimensions before changing orientation
                    old_width, old_height = section.page_width, section.page_height
                    # Set orientation first
                    section.orientation = WD_ORIENT.LANDSCAPE
                    # If page dimensions are not explicitly set in properties, swap them
                    if "page_width" not in properties and "page_height" not in properties:
                        section.page_width, section.page_height = old_height, old_width
                else:
                    # Already landscape, just ensure orientation is set
                    section.orientation = WD_ORIENT.LANDSCAPE
            elif orientation == "PORTRAIT":
                # If changing to portrait, may need to swap width and height
                if section.orientation == WD_ORIENT.LANDSCAPE:
                    # Store current dimensions before changing orientation
                    old_width, old_height = section.page_width, section.page_height
                    # Set orientation first
                    section.orientation = WD_ORIENT.PORTRAIT
                    # If page dimensions are not explicitly set in properties, swap them
                    if "page_width" not in properties and "page_height" not in properties:
                        section.page_width, section.page_height = old_height, old_width
                else:
                    # Already portrait, just ensure orientation is set
                    section.orientation = WD_ORIENT.PORTRAIT
            else:
                return f"Error: Invalid orientation '{orientation}'. Valid values are: PORTRAIT, LANDSCAPE"
        
        # Handle page dimensions (after orientation changes, if any)
        if "page_width" in properties:
            section.page_width = int(float(properties["page_width"]) * 914400)  # Convert inches to EMUs
        
        if "page_height" in properties:
            section.page_height = int(float(properties["page_height"]) * 914400)  # Convert inches to EMUs
        
        # Handle margins
        for margin_prop in ["left_margin", "right_margin", "top_margin", "bottom_margin", 
                           "gutter", "header_distance", "footer_distance"]:
            if margin_prop in properties:
                setattr(section, margin_prop, int(float(properties[margin_prop]) * 914400))  # Convert inches to EMUs
        
        doc_path = get_document_path(doc_id)
        document.save(doc_path)
        return f"Properties for section {section_index} updated successfully."
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error setting section properties: {str(e)}"

def change_page_orientation(doc_id: str, section_index: int, orientation: str) -> str:
    """Changes the page orientation for a specific section.
    
    This is a convenience function that wraps set_section_properties for
    the common task of changing page orientation.
    
    Args:
        doc_id (str): The document ID (filename without extension).
        section_index (int): The index of the section to modify (0-based).
        orientation (str): "PORTRAIT" or "LANDSCAPE"
    
    Returns:
        str: A message indicating success or failure.
    """
    return set_section_properties(doc_id, section_index, {"orientation": orientation})
