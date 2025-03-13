#!/usr/bin/env python
"""
Launcher script for the MCP Word Document Server.
Run this file to start the server.
"""

import os
import sys
import importlib.util

# Get the absolute path to the current directory
current_dir = os.path.dirname(os.path.abspath(__file__))

# Add the current directory to sys.path if not already there
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

# Define the path to the server module
server_module_path = os.path.join(current_dir, "mcp_docx_server", "server.py")

# Check if the server module exists
if not os.path.exists(server_module_path):
    print(f"Error: Server module not found at {server_module_path}")
    sys.exit(1)

# Load the server module
spec = importlib.util.spec_from_file_location("server_module", server_module_path)
server_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(server_module)

# Get the mcp object from the server module
mcp = server_module.mcp

if __name__ == "__main__":
    print("Starting MCP Word Document Server...")
    print(f"Server directory: {current_dir}")
    mcp.run()
