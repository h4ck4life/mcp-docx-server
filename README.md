# MCP DOCX Server

This is the MCP DOCX Server. Below is an example configuration to run the server.

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
