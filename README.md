# MCP Word Document Server

A FastMCP server that allows reading and manipulating Microsoft Word (.docx) files.

## Features

- Create, read, and modify Word documents
- Apply formatting and styles
- Add tables, images, headers, and footers
- Manage document sections
- Convert documents to PDF

## Installation

1. Clone this repository:
```bash
git clone https://github.com/yourusername/mcp-docx-server.git
cd mcp-docx-server
```

2. Install the package:
```bash
pip install -e .
```

## Usage

### Starting the Server

You can run the server using either of these methods:

```bash
# Option 1: Using the console script
mcp-docx-server

# Option 2: Using the runner script
python server_runner.py
```

### Working with Documents

#### Reading Documents

To read an existing document:
- Use the resource: `word://document_name/content` (replace "document_name" with the filename without .docx extension)
- Or call the tool: `read_document("document_name")`

#### Creating Documents

```python
# Create a new document
create_document("my_doc", "My Document Title")

# Add content
add_paragraph("my_doc", "This is a paragraph of text")
add_heading("my_doc", "Section Heading", 1)
```

#### Formatting and Styles

```python
# Apply styles
ensure_style_exists("my_doc", "Heading 1", "paragraph")
add_paragraph("my_doc", "Styled text", style="Heading 1")

# Create custom styles
create_custom_style("my_doc", "MyStyle", "paragraph", "Normal")
```

#### Tables

```python
# Create a table
add_table("my_doc", 3, 3, "Cell 1,Cell 2,Cell 3,Cell 4,Cell 5,Cell 6,Cell 7,Cell 8,Cell 9", "Table Grid")

# Merge cells
merge_table_cells("my_doc", 0, 0, 0, 0, 1)
```

#### Headers and Footers

```python
# Add headers and footers
add_zoned_header("my_doc", 0, "Left", "Center", "Right")
add_footer("my_doc", 0, "Page X of Y")
```

## Project Structure

- `mcp_docx_server/`: Main package
  - `server.py`: FastMCP server definition
  - `utils.py`: Utility functions
  - `document_ops.py`: Document creation and management
  - `style_ops.py`: Style operations
  - `content_ops.py`: Content operations (paragraphs, tables, etc.)
  - `section_ops.py`: Section operations
  - `header_footer_ops.py`: Header and footer operations
- `server_runner.py`: Script to run the server

## Dependencies

- [mcp](https://github.com/microsoft/mcp): Microsoft Conversational Protocol
- [python-docx](https://python-docx.readthedocs.io/): Python library for Word documents
- [docx2pdf](https://github.com/AlJohri/docx2pdf): Convert Word to PDF

## License

See the [LICENSE](LICENSE) file for details.
