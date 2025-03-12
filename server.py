from mcp.server.fastmcp import FastMCP, Context
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from io import BytesIO
import base64
import os
from docx2pdf import convert
import tempfile

# Create an MCP server specifically for Word document operations
mcp = FastMCP("WordDocServer", 
              description="An MCP server that allows reading and manipulating Microsoft Word (.docx) files. "
                          "This server can create, read, and modify Word documents stored in the same directory as the script.")

# Helper function to get the path to a document in the same directory as the script
def get_document_path(doc_id: str) -> str:
    """Returns the full path to a document in the same directory as this script."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(script_dir, f"{doc_id}.docx")

# Helper function to load a document safely
def load_document(doc_id: str) -> Document:
    """Loads a Word document, handling potential FileNotFoundError."""
    doc_path = get_document_path(doc_id)
    try:
        return Document(doc_path)
    except FileNotFoundError:
        raise ValueError(f"Document '{doc_id}.docx' not found.")
    except Exception as e:
        raise ValueError(f"Error loading document '{doc_id}.docx': {str(e)}")


@mcp.resource("word://{doc_id}/content")
def get_document_content(doc_id: str) -> str:
    """Reads the content of a Microsoft Word (.docx) document and returns it as text.
    
    This resource allows directly reading Word documents by document ID.
    
    Example usage:
    To read a file named "bitcoin_overview.docx", request: word://bitcoin_overview/content
    
    Args:
        doc_id (str): The document ID (filename without extension).
        
    Returns:
        str: The full text content of the document with paragraphs separated by newlines.
    """
    try:
        document = load_document(doc_id)
        full_text = []
        for paragraph in document.paragraphs:
            full_text.append(paragraph.text)
        return '\n'.join(full_text)
    except ValueError as e:  # Catch document loading errors
        return str(e)  # Return error message to the client
    except Exception as e:
        return f"Unexpected error: {str(e)}"

@mcp.tool()
def read_document(doc_id: str) -> str:
    """Reads the entire content of a Word document.
    
    This tool provides an easy way to read the contents of any .docx file
    that exists in the server's directory.
    
    Example: To read "bitcoin_overview.docx", call this function with doc_id="bitcoin_overview"
    
    Args:
        doc_id (str): The document ID (filename without extension).
        
    Returns:
        str: The full text content of the document.
    """
    try:
        document = load_document(doc_id)
        full_text = []
        for paragraph in document.paragraphs:
            full_text.append(paragraph.text)
        return '\n'.join(full_text)
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error reading document: {str(e)}"


@mcp.tool()
def check_document_exists(doc_id: str) -> str:
    """Checks if a Word document exists and can be read.
    
    Args:
        doc_id (str): The document ID (filename without extension).
        
    Returns:
        str: A message indicating whether the document exists and is readable.
    """
    doc_path = get_document_path(doc_id)
    try:
        if os.path.exists(doc_path):
            # Try to open and read the document to verify it's readable
            document = Document(doc_path)
            paragraph_count = len(document.paragraphs)
            return f"Document '{doc_id}.docx' exists and is readable at path: {os.path.abspath(doc_path)}. Contains {paragraph_count} paragraphs."
        else:
            return f"Document '{doc_id}.docx' does not exist at path: {os.path.abspath(doc_path)}"
    except Exception as e:
        return f"Document '{doc_id}.docx' exists but cannot be read: {str(e)}"


@mcp.tool()
def create_document(doc_id: str, title: str = "New Document") -> str:
    """Creates a new Word document with a title."""
    try:
        document = Document()
        document.add_heading(title, 0)
        doc_path = get_document_path(doc_id)
        document.save(doc_path)
        return f"Document '{doc_id}.docx' created successfully at path: {os.path.abspath(doc_path)}"
    except Exception as e:
        return f"Error creating document: {str(e)}"

@mcp.tool()
def create_complete_document(doc_id: str, title: str = "New Document", content: list = None) -> str:
    """Creates a new Word document with title and content in a single operation.
    
    This tool simplifies document creation by allowing you to provide a title and
    multiple content elements in a single call. Each content element is a dictionary
    that specifies the type and text, along with any additional formatting options.
    
    Content types:
    - "heading": Adds a heading with specified level (0-4)
    - "paragraph": Adds a paragraph with optional style
    - "table": Adds a table with specified rows, columns, and comma-separated data
    
    Args:
        doc_id (str): The document ID (filename without extension).
        title (str): The document title.
        content (list): A list of content dictionaries. Each dictionary must have:
                      - "type": The content type ("heading", "paragraph", "table")
                      - "text": The content text
                      - Other optional keys depending on the type
    
    Example content list:
    [
        {"type": "heading", "text": "Introduction", "level": 1},
        {"type": "paragraph", "text": "This is a sample paragraph."},
        {"type": "heading", "text": "Data", "level": 2},
        {"type": "table", "rows": 2, "cols": 2, "data": "A,B,C,D"}
    ]
    
    Returns:
        str: A message indicating success or failure of the document creation.
    """
    try:
        # Create a new document
        document = Document()
        
        # Add the title
        document.add_heading(title, 0)
        
        # Add content if provided
        if content:
            for item in content:
                content_type = item.get("type", "").lower()
                text = item.get("text", "")
                
                if content_type == "heading":
                    level = item.get("level", 1)
                    document.add_heading(text, level)
                    
                elif content_type == "paragraph":
                    style = item.get("style")
                    if style:
                        try:
                            document.add_paragraph(text, style=style)
                        except KeyError:
                            document.add_paragraph(text)
                    else:
                        document.add_paragraph(text)
                        
                elif content_type == "table":
                    rows = item.get("rows", 1)
                    cols = item.get("cols", 1)
                    data = item.get("data", "")
                    
                    table = document.add_table(rows=rows, cols=cols)
                    data_list = data.split(',')
                    
                    if len(data_list) != rows * cols:
                        return f"Error in table data: Number of data elements ({len(data_list)}) does not match table dimensions ({rows}x{cols})."
                    
                    for i in range(rows):
                        for j in range(cols):
                            table.cell(i, j).text = data_list[i * cols + j].strip()
        
        # Save the document
        doc_path = get_document_path(doc_id)
        document.save(doc_path)
        
        return f"Document '{doc_id}.docx' created successfully with title and {len(content) if content else 0} content items at path: {os.path.abspath(doc_path)}"
    except Exception as e:
        return f"Error creating document: {str(e)}"

@mcp.tool()
def list_available_documents() -> str:
    """Lists all Word documents (.docx files) available in the server directory."""
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        docx_files = [f.replace('.docx', '') for f in os.listdir(script_dir) if f.endswith('.docx')]
        
        if not docx_files:
            return "No Word documents (.docx files) found in the server directory."
        
        doc_list = "\n".join([f"- {doc}" for doc in docx_files])
        return f"Available Word documents (without .docx extension):\n{doc_list}"
    except Exception as e:
        return f"Error listing documents: {str(e)}"

@mcp.tool()
def add_paragraph(doc_id: str, text: str, style: str = None) -> str:
    """Adds a paragraph to an existing Word document, optionally with a style."""
    try:
        document = load_document(doc_id)
        if style:
            try:
                document.add_paragraph(text, style=style)
            except KeyError:
                return f"Error: Style '{style}' not found. Added without Style"
        else:
            document.add_paragraph(text)
        doc_path = get_document_path(doc_id)
        document.save(doc_path)
        return "Paragraph added successfully."
    except ValueError as e:
        return str(e) # Propagate the error message from load_document
    except Exception as e:
        return f"Error adding paragraph: {str(e)}"


@mcp.tool()
def add_image(doc_id: str, image_data: str, image_name: str, width_inches: float = 6.0) -> str:
    """Adds an image to an existing Word document.

    Args:
        doc_id (str): The document ID (filename without extension).
        image_data (str): The base64 encoded image data.
        image_name (str): The name of the image (e.g., "image.png").
        width_inches (float): The width of the image in inches.
    """
    try:
        document = load_document(doc_id)
        # Decode the base64 image data
        image_bytes = base64.b64decode(image_data)
        # Create an in-memory file-like object
        image_stream = BytesIO(image_bytes)
        # Add the image to the document
        document.add_picture(image_stream, width=Inches(width_inches))  # Adjust width as needed
        doc_path = get_document_path(doc_id)
        document.save(doc_path)
        return f"Image '{image_name}' added to document '{doc_id}.docx' successfully."
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error adding image: {str(e)}"


@mcp.tool()
def add_heading(doc_id: str, text: str, level: int) -> str:
    """Adds a heading to an existing Word document."""
    try:
        document = load_document(doc_id)
        document.add_heading(text, level)
        doc_path = get_document_path(doc_id)
        document.save(doc_path)
        return "Heading added successfully."
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error adding heading: {str(e)}"


@mcp.tool()
def add_table(doc_id: str, rows: int, cols: int, data: str) -> str:
    """Adds a table to an existing Word document.

    Args:
        doc_id (str): The document ID (filename without extension).
        rows (int): The number of rows in the table.
        cols (int): The number of columns in the table.
        data (str): Comma-separated data for the table cells (row-wise).
    """
    try:
        document = load_document(doc_id)
        table = document.add_table(rows=rows, cols=cols)
        data_list = data.split(',')  # Split comma-separated data
        if len(data_list) != rows * cols:
            return "Error: Number of data elements does not match table dimensions."

        for i in range(rows):
            for j in range(cols):
                table.cell(i, j).text = data_list[i * cols + j].strip()  # Fill cells

        doc_path = get_document_path(doc_id)
        document.save(doc_path)
        return "Table added successfully."
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error adding table: {str(e)}"


@mcp.tool()
def set_paragraph_alignment(doc_id: str, paragraph_index: int, alignment: str) -> str:
    """Sets the alignment of a specific paragraph in the document.

    Args:
        doc_id (str): The document ID (filename without extension).
        paragraph_index (int): The index of the paragraph (0-based).
        alignment (str): The desired alignment ('LEFT', 'CENTER', 'RIGHT', 'JUSTIFY').
    """
    try:
        document = load_document(doc_id)
        if paragraph_index < 0 or paragraph_index >= len(document.paragraphs):
            return "Error: Paragraph index out of range."

        paragraph = document.paragraphs[paragraph_index]

        if alignment == 'LEFT':
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        elif alignment == 'CENTER':
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif alignment == 'RIGHT':
            paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif alignment == 'JUSTIFY':
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        else:
            return "Error: Invalid alignment value."

        doc_path = get_document_path(doc_id)
        document.save(doc_path)
        return f"Paragraph {paragraph_index} alignment set to {alignment}."
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error setting paragraph alignment: {str(e)}"


@mcp.tool()
def set_paragraph_font(doc_id: str, paragraph_index: int, font_name: str = None, font_size: int = None, bold: bool = None, italic: bool = None, underline: bool = None) -> str:
    """Sets the font properties of a specific paragraph in the document.

    Args:
        doc_id (str): The document ID (filename without extension).
        paragraph_index (int): The index of the paragraph (0-based).
        font_name (str, optional): The desired font name (e.g., "Arial").
        font_size (int, optional): The desired font size in points.
        bold (bool, optional): Whether to set the font to bold.
        italic (bool, optional): Whether to set the font to italic.
        underline (bool, optional): Whether to set the font to underlined.
    """
    try:
        document = load_document(doc_id)
        if paragraph_index < 0 or paragraph_index >= len(document.paragraphs):
            return "Error: Paragraph index out of range."

        paragraph = document.paragraphs[paragraph_index]
        for run in paragraph.runs:  # Apply formatting to all runs in the paragraph
            font = run.font
            if font_name:
                font.name = font_name
            if font_size:
                font.size = Pt(font_size)
            if bold is not None:
                font.bold = bold
            if italic is not None:
                font.italic = italic
            if underline is not None:
                font.underline = underline

        doc_path = get_document_path(doc_id)
        document.save(doc_path)
        return f"Font properties set for paragraph {paragraph_index}."
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error setting paragraph font: {str(e)}"

@mcp.tool()
def list_styles(doc_id: str) -> str:
    """Lists available styles in the document."""
    try:
        document = load_document(doc_id)
        styles_list = []
        for style in document.styles:
            if style.type == WD_STYLE_TYPE.PARAGRAPH:
                styles_list.append(style.name)
        return ", ".join(styles_list)
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error listing styles: {str(e)}"

@mcp.tool()
def convert_to_pdf(doc_id: str) -> str:
    """Converts a Word document to PDF format.
    
    This tool converts a .docx file to PDF and saves it in the same directory
    with the same name but .pdf extension.
    
    Args:
        doc_id (str): The document ID (filename without extension).
        
    Returns:
        str: A message indicating success or failure of the conversion.
    """
    try:
        doc_path = get_document_path(doc_id)
        if not os.path.exists(doc_path):
            return f"Error: Document '{doc_id}.docx' not found."
        
        script_dir = os.path.dirname(os.path.abspath(__file__))
        pdf_path = os.path.join(script_dir, f"{doc_id}.pdf")
        
        # Convert docx to PDF
        convert(doc_path, pdf_path)
        
        return f"Document successfully converted to PDF at: {os.path.abspath(pdf_path)}"
    except Exception as e:
        return f"Error converting document to PDF: {str(e)}"

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

## Creating Documents
### Quick Method (All-in-One)
Use the comprehensive creation function that handles multiple content elements in one call:

```python
create_complete_document(
    "my_document", 
    "Document Title",
    [
        {"type": "heading", "text": "Introduction", "level": 1},
        {"type": "paragraph", "text": "This is the first paragraph."},
        {"type": "heading", "text": "Data Section", "level": 2},
        {"type": "table", "rows": 2, "cols": 2, "data": "A,B,C,D"}
    ]
)
```

### Step-by-Step Method
1. First create a document: `create_document("my_doc", "My Document Title")`
2. Add content using any of these tools:
   - `add_paragraph("my_doc", "This is a paragraph of text")`
   - `add_heading("my_doc", "Section Heading", 1)` (levels 0-4, where 0 is title)
   - `add_table("my_doc", 3, 3, "Cell 1,Cell 2,Cell 3,Cell 4,Cell 5,Cell 6,Cell 7,Cell 8,Cell 9")`
   - `add_image("my_doc", base64_image_data, "image.png", 4.0)`

## Formatting Content
- Set alignment: `set_paragraph_alignment("my_doc", 1, "CENTER")` (options: LEFT, CENTER, RIGHT, JUSTIFY)
- Set font properties: `set_paragraph_font("my_doc", 1, font_name="Arial", font_size=12, bold=True)`

## Converting to PDF
To convert a Word document to PDF format:
- `convert_to_pdf("my_document")` - This will create a PDF with the same name in the server directory

## Utility Functions
- Check if a document exists: `check_document_exists("my_document")`
- List all available documents: `list_available_documents()`
- List available styles in a document: `list_styles("my_document")`

## Tips for Working with Word Documents
- Check if a document exists before trying to modify it
- Use paragraph indexes carefully (they start at 0)
- Remember that changes are saved immediately
- Word styles can be used for consistent formatting

Remember to check the document path returned by create_document() to know where your files are stored.
"""

if __name__ == "__main__":
    mcp.run()