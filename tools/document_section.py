from docx.enum.section import WD_SECTION
from docx.enum.section import WD_ORIENT
from docx.shared import Inches
from utils.helpers import load_document, save_document

def register_document_section_tools(mcp):
    """Register all section-related tools with the MCP server."""
    
    # Check if tools are already registered to avoid duplicates
    registered_tools = getattr(mcp, 'registered_tools', set())
    
    @mcp.tool()
    def list_sections(document_name: str) -> list:
        """
        Lists all sections in the document with basic information.
        
        Args:
            document_name: The name of the document (without .docx extension)
            
        Returns:
            List of dictionaries with section information
        """
        doc = load_document(document_name)
        sections_info = []
        
        for i, section in enumerate(doc.sections):
            sections_info.append({
                "index": i,
                "orientation": "PORTRAIT" if section.orientation == WD_ORIENT.PORTRAIT else "LANDSCAPE",
                "page_width": round(section.page_width.inches, 2),
                "page_height": round(section.page_height.inches, 2),
                "start_type": str(section.start_type).split(".")[1] if section.start_type else "NEW_PAGE"
            })
            
        return sections_info
    
    if 'add_section' not in registered_tools:
        @mcp.tool()
        def add_section(document_name: str, start_type: str = "NEW_PAGE") -> int:
            """
            Adds a new section to the document.
            
            Args:
                document_name: The name of the document (without .docx extension)
                start_type: Section break type (NEW_PAGE, CONTINUOUS, EVEN_PAGE, ODD_PAGE)
                
            Returns:
                The index of the newly added section or -1 if failed
            """
            doc = load_document(document_name)
            
            try:
                # Map string to WD_SECTION enum
                section_types = {
                    "CONTINUOUS": WD_SECTION.CONTINUOUS,
                    "NEW_PAGE": WD_SECTION.NEW_PAGE,
                    "EVEN_PAGE": WD_SECTION.EVEN_PAGE,
                    "ODD_PAGE": WD_SECTION.ODD_PAGE
                }
                
                section_type = section_types.get(start_type.upper(), WD_SECTION.NEW_PAGE)
                
                # Add a new section break
                doc.add_section(section_type)
                
                section_idx = len(doc.sections) - 1
                save_document(doc, document_name)
                
                return section_idx
            except Exception as e:
                print(f"Error adding section: {e}")
                return -1
    
    if 'change_page_orientation' not in registered_tools:
        @mcp.tool()
        def change_page_orientation(document_name: str, section_index: int, 
                                  orientation: str) -> bool:
            """
            Changes the orientation of a section's pages.
            
            Args:
                document_name: The name of the document (without .docx extension)
                section_index: The index of the section to modify
                orientation: New orientation (PORTRAIT or LANDSCAPE)
                
            Returns:
                True if successful, False otherwise
            """
            doc = load_document(document_name)
            
            try:
                if section_index < 0 or section_index >= len(doc.sections):
                    return False
                    
                section = doc.sections[section_index]
                
                # Set orientation
                if orientation.upper() == "LANDSCAPE":
                    section.orientation = WD_ORIENT.LANDSCAPE
                    # Swap width and height
                    current_width = section.page_width
                    section.page_width = section.page_height
                    section.page_height = current_width
                else:
                    section.orientation = WD_ORIENT.PORTRAIT
                    # Swap width and height if currently in landscape
                    if section.page_width.inches > section.page_height.inches:
                        current_width = section.page_width
                        section.page_width = section.page_height
                        section.page_height = current_width
                
                save_document(doc, document_name)
                return True
            except Exception as e:
                print(f"Error changing orientation: {e}")
                return False
    
    @mcp.tool()
    def set_section_properties(document_name: str, section_index: int, 
                              properties: dict) -> bool:
        """
        Sets multiple properties for a section.
        
        Args:
            document_name: The name of the document (without .docx extension)
            section_index: The index of the section to modify
            properties: Dictionary of properties to set:
                {
                    "orientation": "PORTRAIT" or "LANDSCAPE",
                    "page_width": 8.5,  # in inches
                    "page_height": 11.0,  # in inches
                    "left_margin": 1.0,  # in inches
                    "right_margin": 1.0,  # in inches
                    "top_margin": 1.0,  # in inches
                    "bottom_margin": 1.0,  # in inches
                    "header_distance": 0.5,  # in inches
                    "footer_distance": 0.5,  # in inches
                }
            
        Returns:
            True if successful, False otherwise
        """
        doc = load_document(document_name)
        
        try:
            if section_index < 0 or section_index >= len(doc.sections):
                return False
                
            section = doc.sections[section_index]
            
            # Set orientation
            if "orientation" in properties:
                if properties["orientation"].upper() == "LANDSCAPE":
                    section.orientation = WD_ORIENT.LANDSCAPE
                else:
                    section.orientation = WD_ORIENT.PORTRAIT
            
            # Set page dimensions
            if "page_width" in properties:
                section.page_width = Inches(properties["page_width"])
            if "page_height" in properties:
                section.page_height = Inches(properties["page_height"])
            
            # Set margins
            if "left_margin" in properties:
                section.left_margin = Inches(properties["left_margin"])
            if "right_margin" in properties:
                section.right_margin = Inches(properties["right_margin"])
            if "top_margin" in properties:
                section.top_margin = Inches(properties["top_margin"])
            if "bottom_margin" in properties:
                section.bottom_margin = Inches(properties["bottom_margin"])
            
            # Set header and footer distances
            if "header_distance" in properties:
                section.header_distance = Inches(properties["header_distance"])
            if "footer_distance" in properties:
                section.footer_distance = Inches(properties["footer_distance"])
            
            save_document(doc, document_name)
            return True
        except Exception as e:
            print(f"Error setting section properties: {e}")
            return False
    
    @mcp.tool()
    def copy_section_properties(document_name: str, source_section: int, 
                               target_section: int) -> bool:
        """
        Copies properties from one section to another.
        
        Args:
            document_name: The name of the document (without .docx extension)
            source_section: The index of the section to copy from
            target_section: The index of the section to copy to
            
        Returns:
            True if successful, False otherwise
        """
        doc = load_document(document_name)
        
        try:
            if (source_section < 0 or source_section >= len(doc.sections) or
                target_section < 0 or target_section >= len(doc.sections)):
                return False
                
            source = doc.sections[source_section]
            target = doc.sections[target_section]
            
            # Copy section properties
            target.orientation = source.orientation
            target.page_width = source.page_width
            target.page_height = source.page_height
            target.left_margin = source.left_margin
            target.right_margin = source.right_margin
            target.top_margin = source.top_margin
            target.bottom_margin = source.bottom_margin
            target.header_distance = source.header_distance
            target.footer_distance = source.footer_distance
            
            save_document(doc, document_name)
            return True
        except:
            return False
    
    # Update the list of registered tools
    if not hasattr(mcp, 'registered_tools'):
        mcp.registered_tools = set()
    mcp.registered_tools.update(['list_sections', 'add_section', 'change_page_orientation', 
                               'set_section_properties', 'copy_section_properties'])
    
    print("Document Section tools registered.")
    return mcp
