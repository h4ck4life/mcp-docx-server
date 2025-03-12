# MCP docx server

This is the MCP docx server. It is used to manipulate DOCX files, including creating and editing them. Below is an example configuration to run the server in Claude Desktop.

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

## License

This project is licensed under the MIT License. See the [LICENSE](./LICENSE) file for details.
