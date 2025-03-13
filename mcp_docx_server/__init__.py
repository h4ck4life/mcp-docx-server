"""
MCP Word Document Server

A FastMCP server that allows reading and manipulating Microsoft Word (.docx) files.
"""

__version__ = "0.1.0"

# These imports make the package directly importable
# You can access the mcp object with: from mcp_docx_server import mcp
from .server import mcp
