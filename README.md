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
        "mcp[cli],python-docx,docx2pdf",
        "mcp",
        "run",
        "<your local path>/mcp-docx-server/server_runner.py"
      ]
    }
  }
}
```

## License

See the [LICENSE](LICENSE) file for details.
