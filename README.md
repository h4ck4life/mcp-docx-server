# MCP DOCX Server

This is the MCP DOCX Server. It is used to manipulate DOCX files, including creating and editing them. Below is an example configuration to run the server in Claude Desktop.

## Example Configuration

```json
{
  "mcpServers": {
    "WordDocServer": {
      "command": "uv",
      "args": [
        "run",
        "--with",
        "mcp[cli],python-docx",
        "mcp",
        "run",
        "<your local path>/mcp-docx-server/server.py"
      ]
    }
  }
}
```
