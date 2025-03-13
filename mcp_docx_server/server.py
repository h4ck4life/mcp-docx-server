"""
MCP server for Word document operations.
"""

from mcp.server.fastmcp import FastMCP, Context

# Import all operation modules using absolute imports
from mcp_docx_server.document_ops import (
    create_document, create_complete_document, update_document,
    append_to_document, replace_document, read_document,
    check_document_exists, list_available_documents,
    convert_to_pdf, analyze_document_structure
)

from mcp_docx_server.style_ops import (
    ensure_style_exists, create_custom_style, modify_style,
    get_styles_detail, check_style_usage, list_styles
)

from mcp_docx_server.content_ops import (
    add_paragraph, add_formatted_text, add_image, add_heading,
    add_table, merge_table_cells, get_table_data, list_tables,
    set_paragraph_properties, set_text_properties
)

from mcp_docx_server.section_ops import (
    add_section, list_sections, set_section_properties, change_page_orientation
)

from mcp_docx_server.header_footer_ops import (
    add_header, add_footer, add_zoned_header, add_zoned_footer,
    remove_header, remove_footer, get_header_text, get_footer_text
)

# Create an MCP server specifically for Word document operations
mcp = FastMCP("WordDocServer", 
              description="An MCP server that allows reading and manipulating Microsoft Word (.docx) files. "
                          "This server can create, read, and modify Word documents stored in the same directory as the script.")

# Resource for reading document content
@mcp.resource("word://{doc_id}/content")
def get_document_content(doc_id: str) -> str:
    """Reads the content of a Microsoft Word (.docx) document and returns it as text."""
    return read_document(doc_id)

# Register all the document operations
mcp.tool()(read_document)
mcp.tool()(check_document_exists)
mcp.tool()(list_available_documents)
mcp.tool()(create_document)
mcp.tool()(create_complete_document)
mcp.tool()(update_document)
mcp.tool()(append_to_document)
mcp.tool()(replace_document)
mcp.tool()(convert_to_pdf)
mcp.tool()(analyze_document_structure)

# Register all the style operations
mcp.tool()(ensure_style_exists)
mcp.tool()(create_custom_style)
mcp.tool()(modify_style)
mcp.tool()(get_styles_detail)
mcp.tool()(check_style_usage)
mcp.tool()(list_styles)

# Register all the content operations
mcp.tool()(add_paragraph)
mcp.tool()(add_formatted_text)
mcp.tool()(add_image)
mcp.tool()(add_heading)
mcp.tool()(add_table)
mcp.tool()(merge_table_cells)
mcp.tool()(get_table_data)
mcp.tool()(list_tables)
mcp.tool()(set_paragraph_properties)
mcp.tool()(set_text_properties)

# Register all the section operations
mcp.tool()(add_section)
mcp.tool()(list_sections)
mcp.tool()(set_section_properties)
mcp.tool()(change_page_orientation)

# Register all the header/footer operations
mcp.tool()(add_header)
mcp.tool()(add_footer)
mcp.tool()(add_zoned_header)
mcp.tool()(add_zoned_footer)
mcp.tool()(remove_header)
mcp.tool()(remove_footer)
mcp.tool()(get_header_text)
mcp.tool()(get_footer_text)

@mcp.prompt()
def word_document_usage() -> str:
    """Provides guidance on how to use this MCP server with Word documents."""
    return """
# Word Document Server Usage Guide

This server allows you to create, read, and manipulate Microsoft Word (.docx) documents. Here's how to use it effectively:

## Reading Documents
To read an existing document:
- Use the resource: `word://document_name/content` (replace "document_name" with the filename without .docx extension)
- Or call the tool: `read_document("document_name")`

Example: To read a file named "bitcoin_overview.docx":
- Request the resource: `word://bitcoin_overview/content`
- Or call: `read_document("bitcoin_overview")`

## Working with Styles
Word documents heavily rely on styles for consistent formatting. This server provides several tools for working with styles:

### Style Management
- `ensure_style_exists("document_name", "Heading 1", "paragraph")` - Ensures a built-in style exists in the document
- `create_custom_style("document_name", "MyCustomStyle", "paragraph", "Normal")` - Creates a new custom style
- `modify_style("document_name", "MyStyle", {"font": {"size": 12, "bold": True}, "paragraph": {"alignment": "CENTER"}})` - Modifies a style
- `get_styles_detail("document_name", "paragraph")` - Gets detailed information about styles
- `check_style_usage("document_name", "Heading 1")` - Checks where a style is used in a document
- `list_styles("document_name")` - Lists all styles in a document

### Style Tips
- Always ensure a style exists before using it (Word ignores undefined styles)
- Use `ensure_style_exists()` to define built-in styles before using them
- Apply styles by name when adding content: `add_paragraph("document_name", "Text", style="Heading 1")`
- Custom styles can be based on existing styles using the `base_style` parameter
- Style changes affect all content using that style

## Creating Documents

### Step-by-Step Method

First create a document: `create_document("my_doc", "My Document Title")`

Add content using any of these tools:
- `add_paragraph("my_doc", "This is a paragraph of text", style="Normal", formatting={"alignment": "CENTER"})`
- `add_heading("my_doc", "Section Heading", 1)` (levels 0-4, where 0 is title)
- `add_table("my_doc", 3, 3, "Cell 1,Cell 2,Cell 3,Cell 4,Cell 5,Cell 6,Cell 7,Cell 8,Cell 9", "Table Grid")`
- `add_image("my_doc", base64_image_data, "image.png", 4.0)`

### Updating Documents

Append content to existing documents:
`append_to_document("my_doc", [{"type": "paragraph", "text": "New content", "style": "Normal"}])`

Replace existing document:
`replace_document("my_doc", "New Title", [{"type": "paragraph", "text": "Replacement content", "style": "Normal"}])`

## Text Formatting

### Paragraph Formatting
Set paragraph properties:
```
set_paragraph_properties("my_doc", 1, {
    "alignment": "CENTER",
    "left_indent": 0.5,        # in inches
    "right_indent": 0.5,       # in inches
    "first_line_indent": 0.25, # in inches
    "space_before": 12,        # in points
    "space_after": 12,         # in points
    "line_spacing": 1.5,       # multiple or points
    "keep_together": True,     # pagination options"
    "keep_with_next": True,
    "page_break_before": False,
    "widow_control": True
})
```

### Text/Run Formatting
Add formatted text to an existing paragraph:
```
add_formatted_text("my_doc", 1, "This text will be formatted", {
    "name": "Arial",          # font name
    "size": 12,               # point size
    "bold": True,             # Boolean
    "italic": False,          # Boolean
    "underline": True,        # Boolean
    "color": "#FF0000"        # hex color or rgb(r,g,b)
})
```

Or set properties of an existing text run:
```
set_text_properties("my_doc", 1, 0, {  # paragraph 1, run 0
    "bold": True,
    "color": "rgb(0,0,255)"
})
```

## Table Operations

- Create tables: `add_table("my_doc", 3, 3, "A,B,C,D,E,F,G,H,I", "Table Grid")`
- Merge cells: `merge_table_cells("my_doc", 0, 0, 0, 0, 1)` (table 0, from cell (0,0) to (0,1))
- Get table data: `get_table_data("my_doc", 0)`
- List tables: `list_tables("my_doc")`

## Section Operations
Sections define page layout settings like margins and orientation:

### Working with Sections

- Add a section: `add_section("my_doc", "NEW_PAGE")`  # Other options: "EVEN_PAGE", "ODD_PAGE", "CONTINUOUS"
- List sections: `list_sections("my_doc")`
- Change page orientation: `change_page_orientation("my_doc", 0, "LANDSCAPE")`  # Change section 0 to landscape

### Set Section Properties
```
set_section_properties("my_doc", 0, {  # section 0
    "orientation": "LANDSCAPE",
    "page_width": 11,         # in inches
    "page_height": 8.5,       # in inches
    "left_margin": 1,         # in inches
    "right_margin": 1,        # in inches
    "top_margin": 0.75,       # in inches
    "bottom_margin": 0.75,    # in inches
    "header_distance": 0.5,   # in inches
    "footer_distance": 0.5    # in inches
})
```

## Headers and Footers
Headers and footers are linked to sections and provide content that appears at the top/bottom of each page:

### Basic Headers and Footers

- Add a simple header: `add_header("my_doc", 0, "My Document Title")`  # Add to section 0
- Add a simple footer: `add_footer("my_doc", 0, "Page X of Y")`  # Add to section 0
- Remove a header: `remove_header("my_doc", 0)`  # Remove from section 0
- Remove a footer: `remove_footer("my_doc", 0)`  # Remove from section 0

### Three-Zone Headers and Footers
Create headers/footers with left, center, and right aligned content:
```
add_zoned_header("my_doc", 0, 
                "Left Text",      # Left-aligned
                "Center Text",    # Center-aligned
                "Right Text")     # Right-aligned

add_zoned_footer("my_doc", 0, 
                "Author: John Doe",  # Left-aligned
                "Confidential",      # Center-aligned
                "Page 1")            # Right-aligned
```

### Getting Header/Footer Content

- Get header text: `get_header_text("my_doc", 0)`  # From section 0
- Get footer text: `get_footer_text("my_doc", 0)`  # From section 0

## Analyzing Documents
To understand a document's structure before modifying it:

`analyze_document_structure("my_doc")`

## Converting to PDF (only if requested)
To convert a Word document to PDF format:

`convert_to_pdf("my_document")` - This will create a PDF with the same name in the server directory

## Utility Functions

- Check if a document exists: `check_document_exists("my_document")`
- List all available documents: `list_available_documents()`
- List available styles in a document: `list_styles("my_document")`

## Tips for Working with Word Documents

### Style Management:

- Check if a style exists before using it with `ensure_style_exists()`
- Word ignores style applications if the style isn't defined in the document
- Built-in styles need to be defined (used at least once) before they appear in the document
- Modifying a style affects all content using that style

### Document Structure:

- Use `analyze_document_structure()` to understand document contents
- Check paragraph indexes carefully (they start at 0)
- Remember that changes are saved immediately

### Working with Sections:

- When changing orientation, you may need to explicitly set page dimensions
- Headers and footers are linked to sections, so create sections first if needed
- Most documents start with one section by default

Remember to check the document path returned by `create_document()` to know where your files are stored.
"""

if __name__ == "__main__":
    mcp.run()
