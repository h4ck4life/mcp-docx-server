# server.py
from mcp.server.fastmcp import FastMCP, Context
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from io import BytesIO
import base64
import os

mcp = FastMCP("WordDocServer")

# Helper function to get the path to a document in the home directory
def get_document_path(doc_id: str) -> str:
    """Returns the full path to a document in the user's home directory."""
    home_dir = os.path.expanduser("~")
    return os.path.join(home_dir, f"{doc_id}.docx")

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


@mcp.resource("doc://{doc_id}/content")
def get_document_content(doc_id: str) -> str:
    """Reads the content of a Word document and returns it as text."""
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
def create_document(doc_id: str, title: str = "New Document") -> str:
    """Creates a new Word document with a title."""
    try:
        document = Document()
        document.add_heading(title, 0)
        doc_path = get_document_path(doc_id)
        document.save(doc_path)
        return f"Document '{doc_id}.docx' created successfully in home directory."
    except Exception as e:
        return f"Error creating document: {str(e)}"


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

if __name__ == "__main__":
    mcp.run()