#!/usr/bin/env python
"""
Simple script to start the MCP Word Document Server directly.
This script uses absolute imports to avoid package import issues.
"""

import os
import sys

# Add the project root directory to the Python path
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

# Import the server's mcp object directly
try:
    # Importing this way bypasses the relative import issues
    from mcp_docx_server import server
    mcp = server.mcp
    
    print("Starting MCP Word Document Server...")
    print(f"Server directory: {current_dir}")
    mcp.run()
except ImportError as e:
    print(f"Error importing server module: {e}")
    print("\nTrying to install the package in development mode...")
    
    try:
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-e", "."])
        print("Installation successful. Please run this script again.")
    except Exception as install_error:
        print(f"Installation failed: {install_error}")
        print("\nAlternative solution: manually install the package with:")
        print("pip install -e .")
