from mcp.server.fastmcp import FastMCP, Context
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.dml import MSO_THEME_COLOR
from io import BytesIO
import base64
import os
from docx2pdf import convert

# Create an MCP server specifically for Word document operations
mcp = FastMCP("WordDocServer", 
              description="An MCP server that allows reading and manipulating Microsoft Word (.docx) files. "
                          "This server can create, read, and modify Word documents stored in the same directory as the script.")

# Helper functions
def get_document_path(doc_id: str) -> str:
    """Returns the full path to a document in the same directory as this script."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(script_dir, f"{doc_id}.docx")

def load_document(doc_id: str) -> Document:
    """Loads a Word document, handling potential FileNotFoundError."""
    doc_path = get_document_path(doc_id)
    try:
        return Document(doc_path)
    except FileNotFoundError:
        raise ValueError(f"Document '{doc_id}.docx' not found.")
    except Exception as e:
        raise ValueError(f"Error loading document '{doc_id}.docx': {str(e)}")

# Document reading operations
@mcp.resource("word://{doc_id}/content")
def get_document_content(doc_id: str) -> str:
    """Reads the content of a Microsoft Word (.docx) document and returns it as text."""
    try:
        document = load_document(doc_id)
        full_text = [paragraph.text for paragraph in document.paragraphs]
        return '\n'.join(full_text)
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Unexpected error: {str(e)}"

@mcp.tool()
def read_document(doc_id: str) -> str:
    """Reads the entire content of a Word document."""
    return get_document_content(doc_id)

@mcp.tool()
def check_document_exists(doc_id: str) -> str:
    """Checks if a Word document exists and can be read."""
    doc_path = get_document_path(doc_id)
    try:
        if os.path.exists(doc_path):
            document = Document(doc_path)
            paragraph_count = len(document.paragraphs)
            return f"Document '{doc_id}.docx' exists and is readable at path: {os.path.abspath(doc_path)}. Contains {paragraph_count} paragraphs."
        else:
            return f"Document '{doc_id}.docx' does not exist at path: {os.path.abspath(doc_path)}"
    except Exception as e:
        return f"Document '{doc_id}.docx' exists but cannot be read: {str(e)}"

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

# Document creation and update operations
def _add_content_to_document(document, content):
    """Helper function to add content to a document object."""
    if not content:
        return True
        
    for item in content:
        content_type = item.get("type", "").lower()
        text = item.get("text", "")
        
        if content_type == "heading":
            level = item.get("level", 1)
            heading = document.add_heading(text, level)
            
            # Apply formatting if provided
            formatting = item.get("formatting", {})
            if formatting:
                _apply_paragraph_formatting(heading, formatting)
            
        elif content_type == "paragraph":
            style = item.get("style")
            try:
                paragraph = document.add_paragraph(text, style=style if style else None)
                
                # Apply formatting if provided
                formatting = item.get("formatting", {})
                if formatting:
                    _apply_paragraph_formatting(paragraph, formatting)
                
                # Apply run formatting if provided
                run_formatting = item.get("run_formatting", {})
                if run_formatting and len(paragraph.runs) > 0:
                    _apply_run_formatting(paragraph.runs[0], run_formatting)
                
            except KeyError:
                document.add_paragraph(text)
                
        elif content_type == "table":
            rows = item.get("rows", 1)
            cols = item.get("cols", 1)
            data = item.get("data", "")
            style = item.get("style")
            
            table = document.add_table(rows=rows, cols=cols)
            
            # Apply style if specified
            if style:
                try:
                    table.style = style
                except KeyError:
                    pass  # Continue without style if not found
            
            # Fill with data if provided
            if data:
                data_list = data.split(',')
                
                # Pad with empty strings if too few data elements
                if len(data_list) < rows * cols:
                    data_list.extend([''] * (rows * cols - len(data_list)))
                    
                # Check if data matches table dimensions
                if len(data_list) != rows * cols:
                    return False
                
                # Fill table cells
                for i in range(rows):
                    for j in range(cols):
                        cell_idx = i * cols + j
                        if cell_idx < len(data_list):
                            table.cell(i, j).text = data_list[cell_idx].strip()
            
            # Process cell_formatting if provided
            cell_formatting = item.get("cell_formatting", [])
            for cell_format in cell_formatting:
                row = cell_format.get("row", 0)
                col = cell_format.get("col", 0)
                formatting = cell_format.get("formatting", {})
                
                if row < rows and col < cols:
                    cell = table.cell(row, col)
                    if cell and len(cell.paragraphs) > 0:
                        _apply_paragraph_formatting(cell.paragraphs[0], formatting)
    
    return True

def _apply_paragraph_formatting(paragraph, formatting):
    """Apply formatting to a paragraph."""
    if not formatting:
        return
    
    para_format = paragraph.paragraph_format
    
    # Alignment
    alignment = formatting.get("alignment")
    if alignment:
        alignment_map = {
            "LEFT": WD_ALIGN_PARAGRAPH.LEFT,
            "CENTER": WD_ALIGN_PARAGRAPH.CENTER,
            "RIGHT": WD_ALIGN_PARAGRAPH.RIGHT,
            "JUSTIFY": WD_ALIGN_PARAGRAPH.JUSTIFY
        }
        if alignment.upper() in alignment_map:
            para_format.alignment = alignment_map[alignment.upper()]
    
    # Indentation
    left_indent = formatting.get("left_indent")
    if left_indent is not None:
        para_format.left_indent = Inches(float(left_indent))
    
    right_indent = formatting.get("right_indent")
    if right_indent is not None:
        para_format.right_indent = Inches(float(right_indent))
    
    first_line_indent = formatting.get("first_line_indent")
    if first_line_indent is not None:
        para_format.first_line_indent = Inches(float(first_line_indent))
    
    # Spacing
    space_before = formatting.get("space_before")
    if space_before is not None:
        para_format.space_before = Pt(float(space_before))
    
    space_after = formatting.get("space_after")
    if space_after is not None:
        para_format.space_after = Pt(float(space_after))
    
    # Line spacing
    line_spacing = formatting.get("line_spacing")
    if line_spacing is not None:
        try:
            # Try as a float (multiple)
            spacing_float = float(line_spacing)
            para_format.line_spacing = spacing_float
        except ValueError:
            # Try as a point value
            para_format.line_spacing = Pt(float(line_spacing))
    
    # Pagination
    keep_together = formatting.get("keep_together")
    if keep_together is not None:
        para_format.keep_together = bool(keep_together)
    
    keep_with_next = formatting.get("keep_with_next")
    if keep_with_next is not None:
        para_format.keep_with_next = bool(keep_with_next)
    
    page_break_before = formatting.get("page_break_before")
    if page_break_before is not None:
        para_format.page_break_before = bool(page_break_before)
    
    widow_control = formatting.get("widow_control")
    if widow_control is not None:
        para_format.widow_control = bool(widow_control)

def _apply_run_formatting(run, formatting):
    """Apply formatting to a run of text."""
    if not formatting:
        return
    
    font = run.font
    
    # Font name and size
    font_name = formatting.get("name")
    if font_name:
        font.name = font_name
    
    font_size = formatting.get("size")
    if font_size is not None:
        font.size = Pt(float(font_size))
    
    # Font styles
    bold = formatting.get("bold")
    if bold is not None:
        font.bold = bool(bold)
    
    italic = formatting.get("italic")
    if italic is not None:
        font.italic = bool(italic)
    
    underline = formatting.get("underline")
    if underline is not None:
        font.underline = bool(underline)
    
    # Font color
    color = formatting.get("color")
    if color:
        if color.startswith('#'):
            # Convert hex color to RGB
            r = int(color[1:3], 16)
            g = int(color[3:5], 16)
            b = int(color[5:7], 16)
            font.color.rgb = RGBColor(r, g, b)
        elif color.startswith('rgb('):
            # Parse rgb() format
            rgb = color.strip('rgb()').split(',')
            if len(rgb) == 3:
                r = int(rgb[0].strip())
                g = int(rgb[1].strip())
                b = int(rgb[2].strip())
                font.color.rgb = RGBColor(r, g, b)

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
    """Creates a new Word document with title and content in a single operation."""
    try:
        document = Document()
        document.add_heading(title, 0)
        
        if not _add_content_to_document(document, content):
            return "Error in table data: Number of data elements does not match table dimensions."
        
        doc_path = get_document_path(doc_id)
        document.save(doc_path)
        
        return f"Document '{doc_id}.docx' created successfully with title and {len(content) if content else 0} content items at path: {os.path.abspath(doc_path)}"
    except Exception as e:
        return f"Error creating document: {str(e)}"

@mcp.tool()
def update_document(doc_id: str, title: str = None, content: list = None, append: bool = True) -> str:
    """Updates an existing Word document by appending or replacing content."""
    try:
        doc_path = get_document_path(doc_id)
        
        if not os.path.exists(doc_path) and append:
            return f"Document '{doc_id}.docx' does not exist and cannot be updated. Create it first."
        
        if not append or not os.path.exists(doc_path):
            document = Document()
            if title:
                document.add_heading(title, 0)
        else:
            document = Document(doc_path)
            if title:
                document.add_heading(title, 1)
        
        if not _add_content_to_document(document, content):
            return "Error in table data: Number of data elements does not match table dimensions."
        
        document.save(doc_path)
        
        action = "updated by appending" if append else "replaced"
        title_msg = f" with new title" if title else ""
        content_msg = f" and {len(content) if content else 0} content items" if content else ""
        
        return f"Document '{doc_id}.docx' {action}{title_msg}{content_msg} successfully."
    except Exception as e:
        return f"Error updating document: {str(e)}"

@mcp.tool()
def append_to_document(doc_id: str, content: list) -> str:
    """Appends content to an existing Word document."""
    return update_document(doc_id, title=None, content=content, append=True)

@mcp.tool()
def replace_document(doc_id: str, title: str = None, content: list = None) -> str:
    """Replaces an existing Word document with new content."""
    return update_document(doc_id, title=title, content=content, append=False)

# Individual content addition operations
@mcp.tool()
def add_paragraph(doc_id: str, text: str, style: str = None, formatting: dict = None) -> str:
    """Adds a paragraph to an existing Word document, optionally with style and formatting.
    
    Args:
        doc_id (str): The document ID (filename without extension).
        text (str): The paragraph text.
        style (str, optional): The paragraph style.
        formatting (dict, optional): Dictionary with paragraph formatting options like:
            - alignment: 'LEFT', 'CENTER', 'RIGHT', 'JUSTIFY'
            - left_indent, right_indent, first_line_indent: Indentation in inches
            - space_before, space_after: Spacing in points
            - line_spacing: Line spacing as multiple or points
            - keep_together, keep_with_next, page_break_before, widow_control: Boolean pagination options
    """
    try:
        document = load_document(doc_id)
        if style:
            try:
                paragraph = document.add_paragraph(text, style=style)
            except KeyError:
                paragraph = document.add_paragraph(text)
                result_message = f"Warning: Style '{style}' not found. Added without style."
            else:
                result_message = "Paragraph added successfully."
        else:
            paragraph = document.add_paragraph(text)
            result_message = "Paragraph added successfully."
        
        # Apply formatting if provided
        if formatting:
            _apply_paragraph_formatting(paragraph, formatting)
        
        doc_path = get_document_path(doc_id)
        document.save(doc_path)
        return result_message
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error adding paragraph: {str(e)}"

@mcp.tool()
def add_formatted_text(doc_id: str, paragraph_index: int, text: str, formatting: dict = None) -> str:
    """Adds formatted text to an existing paragraph.
    
    Args:
        doc_id (str): The document ID (filename without extension).
        paragraph_index (int): The index of the paragraph to add text to (0-based).
        text (str): The text to add.
        formatting (dict, optional): Dictionary with text formatting options like:
            - name: Font name
            - size: Font size in points
            - bold, italic, underline: Boolean style options
            - color: Color as hex (#RRGGBB) or rgb(r,g,b)
    """
    try:
        document = load_document(doc_id)
        
        if paragraph_index < 0 or paragraph_index >= len(document.paragraphs):
            return "Error: Paragraph index out of range."
        
        paragraph = document.paragraphs[paragraph_index]
        run = paragraph.add_run(text)
        
        # Apply formatting if provided
        if formatting:
            _apply_run_formatting(run, formatting)
        
        doc_path = get_document_path(doc_id)
        document.save(doc_path)
        return "Formatted text added successfully."
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error adding formatted text: {str(e)}"

@mcp.tool()
def add_image(doc_id: str, image_data: str, image_name: str, width_inches: float = 6.0) -> str:
    """Adds an image to an existing Word document."""
    try:
        document = load_document(doc_id)
        image_bytes = base64.b64decode(image_data)
        image_stream = BytesIO(image_bytes)
        document.add_picture(image_stream, width=Inches(width_inches))
        doc_path = get_document_path(doc_id)
        document.save(doc_path)
        return f"Image '{image_name}' added to document '{doc_id}.docx' successfully."
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error adding image: {str(e)}"

@mcp.tool()
def add_heading(doc_id: str, text: str, level: int, formatting: dict = None) -> str:
    """Adds a heading to an existing Word document with optional formatting.
    
    Args:
        doc_id (str): The document ID (filename without extension).
        text (str): The heading text.
        level (int): The heading level (0-9, where 0 is Title).
        formatting (dict, optional): Dictionary with paragraph formatting options.
    """
    try:
        document = load_document(doc_id)
        heading = document.add_heading(text, level)
        
        # Apply formatting if provided
        if formatting:
            _apply_paragraph_formatting(heading, formatting)
        
        doc_path = get_document_path(doc_id)
        document.save(doc_path)
        return "Heading added successfully."
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error adding heading: {str(e)}"

# Table operations
@mcp.tool()
def add_table(doc_id: str, rows: int, cols: int, data: str = None, style: str = None) -> str:
    """Adds a table to an existing Word document.
    
    Args:
        doc_id (str): The document ID (filename without extension).
        rows (int): The number of rows in the table.
        cols (int): The number of columns in the table.
        data (str, optional): Comma-separated data for the table cells (row-wise).
        style (str, optional): Table style (e.g., "Table Grid", "Light Shading").
    """
    try:
        document = load_document(doc_id)
        
        # Create table with specified dimensions
        table = document.add_table(rows=rows, cols=cols)
        
        # Apply style if specified
        if style:
            try:
                table.style = style
            except KeyError:
                return f"Warning: Style '{style}' not found. Table added with default style."
        
        # Fill with data if provided
        if data:
            data_list = data.split(',')
            
            # Check if data matches table dimensions
            if len(data_list) > rows * cols:
                return f"Error: Number of data elements ({len(data_list)}) exceeds table dimensions ({rows}x{cols})."
                
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
        return f"Table with {rows} rows and {cols} columns added successfully."
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error adding table: {str(e)}"

@mcp.tool()
def merge_table_cells(doc_id: str, table_index: int, start_row: int, start_col: int, end_row: int, end_col: int) -> str:
    """Merges cells in a table.
    
    Args:
        doc_id (str): The document ID (filename without extension).
        table_index (int): The index of the table in the document (0-based).
        start_row (int): The starting row index (0-based).
        start_col (int): The starting column index (0-based).
        end_row (int): The ending row index (0-based).
        end_col (int): The ending column index (0-based).
    """
    try:
        document = load_document(doc_id)
        
        # Check if document has tables
        if not document.tables or len(document.tables) <= table_index:
            return f"Error: Table index {table_index} is out of range. Document has {len(document.tables) if document.tables else 0} tables."
        
        table = document.tables[table_index]
        
        # Validate row and column indices
        if start_row < 0 or end_row >= len(table.rows) or start_col < 0:
            return "Error: Row or column index out of range."
        
        # Get first cell in the range
        first_cell = table.cell(start_row, start_col)
        
        # Get last cell in the range
        last_cell = table.cell(end_row, end_col)
        
        # Merge the cells
        first_cell.merge(last_cell)
        
        doc_path = get_document_path(doc_id)
        document.save(doc_path)
        return f"Cells merged from ({start_row},{start_col}) to ({end_row},{end_col}) in table {table_index}."
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error merging table cells: {str(e)}"

@mcp.tool()
def get_table_data(doc_id: str, table_index: int, include_empty_cells: bool = True) -> str:
    """Retrieves a table's data as formatted text.
    
    Args:
        doc_id (str): The document ID (filename without extension).
        table_index (int): The index of the table in the document (0-based).
        include_empty_cells (bool): Whether to include empty cells in the output.
    """
    try:
        document = load_document(doc_id)
        
        # Check if document has tables
        if not document.tables or len(document.tables) <= table_index:
            return f"Error: Table index {table_index} is out of range. Document has {len(document.tables) if document.tables else 0} tables."
        
        table = document.tables[table_index]
        
        # Get table data
        result = []
        for i, row in enumerate(table.rows):
            row_data = []
            
            # Handle cells before the first actual cell
            for _ in range(row.grid_cols_before):
                if include_empty_cells:
                    row_data.append("")
            
            # Handle actual cells
            for cell in row.cells:
                # Get all text from the cell, including from nested tables
                cell_text = []
                for paragraph in cell.paragraphs:
                    if paragraph.text.strip() or include_empty_cells:
                        cell_text.append(paragraph.text)
                
                # Add nested tables as a note
                if cell.tables:
                    cell_text.append("[Contains nested table]")
                
                row_data.append("\n".join(cell_text))
            
            # Handle cells after the last actual cell
            for _ in range(row.grid_cols_after):
                if include_empty_cells:
                    row_data.append("")
            
            result.append(" | ".join(row_data))
        
        return "\n".join(result)
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error retrieving table data: {str(e)}"

@mcp.tool()
def list_tables(doc_id: str) -> str:
    """Lists all tables in a document with their dimensions.
    
    Args:
        doc_id (str): The document ID (filename without extension).
    """
    try:
        document = load_document(doc_id)
        
        if not document.tables:
            return f"No tables found in document '{doc_id}.docx'."
        
        tables_info = []
        for i, table in enumerate(document.tables):
            row_count = len(table.rows)
            # Get the maximum number of columns across all rows
            col_count = max([row.grid_cols_before + len(row.cells) + row.grid_cols_after 
                            for row in table.rows]) if row_count > 0 else 0
            
            first_cell_text = ""
            if row_count > 0 and len(table.rows[0].cells) > 0:
                first_cell_text = table.rows[0].cells[0].text[:30]
                if len(table.rows[0].cells[0].text) > 30:
                    first_cell_text += "..."
            
            tables_info.append(f"Table {i}: {row_count} rows x {col_count} columns. First cell: '{first_cell_text}'")
        
        return "\n".join(tables_info)
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error listing tables: {str(e)}"

# Text formatting operations
@mcp.tool()
def set_paragraph_properties(doc_id: str, paragraph_index: int, formatting: dict) -> str:
    """Sets various properties of a paragraph.
    
    Args:
        doc_id (str): The document ID (filename without extension).
        paragraph_index (int): Index of the paragraph to format (0-based).
        formatting (dict): Dictionary with paragraph formatting options like:
            - alignment: 'LEFT', 'CENTER', 'RIGHT', 'JUSTIFY'
            - left_indent, right_indent, first_line_indent: Indentation in inches
            - space_before, space_after: Spacing in points
            - line_spacing: Line spacing as multiple or points
            - keep_together, keep_with_next, page_break_before, widow_control: Boolean pagination options
    """
    try:
        document = load_document(doc_id)
        
        if paragraph_index < 0 or paragraph_index >= len(document.paragraphs):
            return "Error: Paragraph index out of range."
        
        paragraph = document.paragraphs[paragraph_index]
        _apply_paragraph_formatting(paragraph, formatting)
        
        doc_path = get_document_path(doc_id)
        document.save(doc_path)
        return f"Paragraph {paragraph_index} properties set successfully."
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error setting paragraph properties: {str(e)}"

@mcp.tool()
def set_text_properties(doc_id: str, paragraph_index: int, run_index: int, formatting: dict) -> str:
    """Sets various properties of a run of text.
    
    Args:
        doc_id (str): The document ID (filename without extension).
        paragraph_index (int): Index of the paragraph containing the run (0-based).
        run_index (int): Index of the run within the paragraph (0-based).
        formatting (dict): Dictionary with text formatting options like:
            - name: Font name
            - size: Font size in points
            - bold, italic, underline: Boolean style options
            - color: Color as hex (#RRGGBB) or rgb(r,g,b)
    """
    try:
        document = load_document(doc_id)
        
        if paragraph_index < 0 or paragraph_index >= len(document.paragraphs):
            return "Error: Paragraph index out of range."
        
        paragraph = document.paragraphs[paragraph_index]
        
        if run_index < 0 or run_index >= len(paragraph.runs):
            return f"Error: Run index {run_index} is out of range. Paragraph has {len(paragraph.runs)} runs."
        
        run = paragraph.runs[run_index]
        _apply_run_formatting(run, formatting)
        
        doc_path = get_document_path(doc_id)
        document.save(doc_path)
        return f"Text properties set for run {run_index} in paragraph {paragraph_index}."
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error setting text properties: {str(e)}"

# Utility functions
@mcp.tool()
def list_styles(doc_id: str) -> str:
    """Lists available paragraph and character styles in the document."""
    try:
        document = load_document(doc_id)
        para_styles = []
        char_styles = []
        table_styles = []
        
        for style in document.styles:
            if style.type == WD_STYLE_TYPE.PARAGRAPH:
                para_styles.append(style.name)
            elif style.type == WD_STYLE_TYPE.CHARACTER:
                char_styles.append(style.name)
            elif style.type == WD_STYLE_TYPE.TABLE:
                table_styles.append(style.name)
        
        result = []
        if para_styles:
            result.append("Paragraph styles:\n" + ", ".join(para_styles))
        if char_styles:
            result.append("\nCharacter styles:\n" + ", ".join(char_styles))
        if table_styles:
            result.append("\nTable styles:\n" + ", ".join(table_styles))
            
        return "\n".join(result)
    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Error listing styles: {str(e)}"

@mcp.tool()
def convert_to_pdf(doc_id: str) -> str:
   """Converts a Word document to PDF format."""
   try:
       doc_path = get_document_path(doc_id)
       if not os.path.exists(doc_path):
           return f"Error: Document '{doc_id}.docx' not found."
       
       script_dir = os.path.dirname(os.path.abspath(__file__))
       pdf_path = os.path.join(script_dir, f"{doc_id}.pdf")
       
       convert(doc_path, pdf_path)
       
       return f"Document successfully converted to PDF at: {os.path.abspath(pdf_path)}"
   except Exception as e:
       return f"Error converting document to PDF: {str(e)}"

@mcp.tool()
def analyze_document_structure(doc_id: str) -> str:
   """Analyzes the structure of a document, showing paragraphs, runs, and tables.
   
   This tool provides a detailed breakdown of the document structure,
   which can be helpful for understanding how to target specific elements
   for modification.
   
   Args:
       doc_id (str): The document ID (filename without extension).
   """
   try:
       document = load_document(doc_id)
       
       structure = []
       structure.append(f"Document Structure Analysis for '{doc_id}.docx':")
       structure.append(f"Total paragraphs: {len(document.paragraphs)}")
       structure.append(f"Total tables: {len(document.tables)}")
       structure.append("\nParagraph Details:")
       
       for i, para in enumerate(document.paragraphs):
           if not para.text.strip():
               structure.append(f"  Paragraph {i}: [Empty paragraph]")
               continue
               
           style = para.style.name if para.style else "Default"
           run_count = len(para.runs)
           structure.append(f"  Paragraph {i}: Style='{style}', Runs={run_count}")
           structure.append(f"    Text: \"{para.text[:50]}{'...' if len(para.text) > 50 else ''}\"")
           
           if run_count > 0:
               structure.append(f"    Run details:")
               for j, run in enumerate(para.runs):
                   bold = "Bold" if run.bold else "Normal"
                   italic = "Italic" if run.italic else "Normal"
                   structure.append(f"      Run {j}: {bold}, {italic}, Text=\"{run.text[:30]}{'...' if len(run.text) > 30 else ''}\"")
       
       if document.tables:
           structure.append("\nTable Details:")
           for i, table in enumerate(document.tables):
               row_count = len(table.rows)
               col_count = max([row.grid_cols_before + len(row.cells) + row.grid_cols_after 
                               for row in table.rows]) if row_count > 0 else 0
               
               structure.append(f"  Table {i}: {row_count} rows x {col_count} columns")
               if table.style:
                   structure.append(f"    Style: {table.style.name}")
               
               # Show a preview of the first few cells
               if row_count > 0 and len(table.rows[0].cells) > 0:
                   structure.append(f"    Preview:")
                   max_preview_rows = min(3, row_count)
                   for r in range(max_preview_rows):
                       row = table.rows[r]
                       cell_texts = []
                       for cell in row.cells[:min(3, len(row.cells))]:
                           cell_text = cell.text[:20]
                           if len(cell.text) > 20:
                               cell_text += "..."
                           cell_texts.append(f"\"{cell_text}\"")
                       
                       additional = "..." if len(row.cells) > 3 else ""
                       structure.append(f"      Row {r}: {', '.join(cell_texts)}{additional}")
       
       return "\n".join(structure)
   except ValueError as e:
       return str(e)
   except Exception as e:
       return f"Error analyzing document structure: {str(e)}"

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
        {"type": "paragraph", "text": "This is the first paragraph.", 
         "formatting": {"alignment": "JUSTIFY", "space_after": 12}},
        {"type": "heading", "text": "Data Section", "level": 2},
        {"type": "table", "rows": 2, "cols": 2, "data": "A,B,C,D", "style": "Table Grid"}
    ]
)
```

### Step-by-Step Method
1. First create a document: `create_document("my_doc", "My Document Title")`
2. Add content using any of these tools:
   - `add_paragraph("my_doc", "This is a paragraph of text", formatting={"alignment": "CENTER"})`
   - `add_heading("my_doc", "Section Heading", 1)` (levels 0-4, where 0 is title)
   - `add_table("my_doc", 3, 3, "Cell 1,Cell 2,Cell 3,Cell 4,Cell 5,Cell 6,Cell 7,Cell 8,Cell 9", "Table Grid")`
   - `add_image("my_doc", base64_image_data, "image.png", 4.0)`

## Updating Documents
### Append content to existing documents:
- `append_to_document("my_doc", [{"type": "paragraph", "text": "New content"}])`

### Replace existing document:
- `replace_document("my_doc", "New Title", [{"type": "paragraph", "text": "Replacement content"}])`

## Text Formatting
### Paragraph Formatting
Set paragraph properties:
```python
set_paragraph_properties("my_doc", 1, {
    "alignment": "CENTER",
    "left_indent": 0.5,        # in inches
    "right_indent": 0.5,       # in inches
    "first_line_indent": 0.25, # in inches
    "space_before": 12,        # in points
    "space_after": 12,         # in points
    "line_spacing": 1.5,       # multiple or points
    "keep_together": True,     # pagination options
    "keep_with_next": True,
    "page_break_before": False,
    "widow_control": True
})
```

### Text/Run Formatting
Add formatted text to an existing paragraph:
```python
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
```python
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

## Analyzing Documents
To understand a document's structure before modifying it:
- `analyze_document_structure("my_doc")`

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
- For complex documents, use analyze_document_structure() first

Remember to check the document path returned by create_document() to know where your files are stored.
"""

if __name__ == "__main__":
    mcp.run()