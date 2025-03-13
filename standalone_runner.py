#!/usr/bin/env python
"""
Standalone script to start the MCP Word Document Server.
This script avoids package import issues by loading the server module directly.
"""

import os
import sys
import importlib

# Get the absolute path of the server directory
current_dir = os.path.dirname(os.path.abspath(__file__))

# Add the current directory to sys.path
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

# Create an empty mcp_docx_server package in memory
if "mcp_docx_server" not in sys.modules:
    sys.modules["mcp_docx_server"] = importlib.util.module_from_spec(
        importlib.util.find_spec("types")
    )

# Define paths
utils_path = os.path.join(current_dir, "mcp_docx_server", "utils.py")
document_ops_path = os.path.join(current_dir, "mcp_docx_server", "document_ops.py")
style_ops_path = os.path.join(current_dir, "mcp_docx_server", "style_ops.py")
content_ops_path = os.path.join(current_dir, "mcp_docx_server", "content_ops.py")
section_ops_path = os.path.join(current_dir, "mcp_docx_server", "section_ops.py")
header_footer_ops_path = os.path.join(current_dir, "mcp_docx_server", "header_footer_ops.py")
server_path = os.path.join(current_dir, "mcp_docx_server", "server.py")

# First, load utils module
utils_spec = importlib.util.spec_from_file_location("mcp_docx_server.utils", utils_path)
utils_module = importlib.util.module_from_spec(utils_spec)
sys.modules["mcp_docx_server.utils"] = utils_module
utils_spec.loader.exec_module(utils_module)

# Load all other modules in the right order
modules = [
    ("mcp_docx_server.document_ops", document_ops_path),
    ("mcp_docx_server.style_ops", style_ops_path), 
    ("mcp_docx_server.content_ops", content_ops_path),
    ("mcp_docx_server.section_ops", section_ops_path),
    ("mcp_docx_server.header_footer_ops", header_footer_ops_path)
]

for module_name, module_path in modules:
    module_spec = importlib.util.spec_from_file_location(module_name, module_path)
    module_obj = importlib.util.module_from_spec(module_spec)
    sys.modules[module_name] = module_obj
    module_spec.loader.exec_module(module_obj)

# Finally, load the server module
server_spec = importlib.util.spec_from_file_location("mcp_docx_server.server", server_path)
server_module = importlib.util.module_from_spec(server_spec)
sys.modules["mcp_docx_server.server"] = server_module
server_spec.loader.exec_module(server_module)

if __name__ == "__main__":
    print("Starting MCP Word Document Server...")
    print(f"Server directory: {current_dir}")
    # Run the server
    server_module.mcp.run()
