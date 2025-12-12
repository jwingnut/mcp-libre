# LibreOffice MCP Server

A comprehensive Model Context Protocol (MCP) server that provides tools and resources for interacting with LibreOffice documents. This server enables AI assistants and other MCP clients to create, read, convert, and manipulate LibreOffice documents programmatically.

[![Python 3.12+](https://img.shields.io/badge/python-3.12+-blue.svg)](https://www.python.org/downloads/)
[![LibreOffice](https://img.shields.io/badge/LibreOffice-24.2+-green.svg)](https://www.libreoffice.org/)
[![MCP Protocol](https://img.shields.io/badge/MCP-2024--11--05-orange.svg)](https://spec.modelcontextprotocol.io/)

## 🚀 Features

### LibreOffice Extension (Plugin) - Recommended!
- **Native Integration**: Embedded MCP server directly in LibreOffice
- **Real-time Editing**: Live document manipulation with instant visual feedback
- **Performance**: 10x faster than external server (direct UNO API access)
- **9 Consolidated Tools**: Reduced from 32 individual tools for better UX
- **Track Changes Support**: Full revision tracking awareness
- **Multi-document**: Work with all open LibreOffice documents
- **HTTP API**: External AI assistant access via localhost:8765

### Document Operations
- **Create Documents**: New Writer, Calc, Impress, and Draw documents
- **Read Content**: Extract text with visible_content for Track Changes awareness
- **Navigate**: Paragraph-level navigation, cursor positioning, document outline
- **Edit**: Insert, format, select, and replace text
- **Search**: Find/replace with Track Changes awareness (skips tracked deletions)
- **Comments**: Add and retrieve document annotations
- **Track Changes**: Enable, disable, list, accept/reject revisions

## 🔧 9 Consolidated MCP Tools

The MCP interface provides 9 logical tool groups (consolidated from 32 individual tools):

| Tool | Actions | Description |
|------|---------|-------------|
| `document` | create, info, list, content, status | Document management |
| `structure` | outline, paragraph, range, count | Document navigation |
| `cursor` | goto_paragraph, goto_position, position, context | Cursor control |
| `selection` | paragraph, range, delete, replace | Text selection |
| `search` | find, replace, replace_all | Search/replace (Track Changes aware) |
| `track_changes` | status, enable, disable, list, accept, reject, accept_all, reject_all | Revision tracking |
| `comments` | list, add | Comment management |
| `save` | save, export | Save/export documents |
| `text` | insert, format | Text insertion and formatting |

See [docs/TOOL_REFERENCE.md](docs/TOOL_REFERENCE.md) for complete documentation.

## 📋 Requirements

- **LibreOffice**: 24.2+ (must be accessible via command line)
- **Python**: 3.12+
- **UV Package Manager**: For dependency management

## 🛠 Installation

### LibreOffice Extension (Recommended)

```bash
# Clone the repository
git clone https://github.com/jwingnut/mcp-libre.git
cd mcp-libre

# Build and install the LibreOffice extension
cd plugin/
./build.sh
unopkg add ../build/libreoffice-mcp-extension-1.0.0.oxt

# Restart LibreOffice
```

After installation:
1. Open LibreOffice Writer
2. Go to **Tools > MCP Server > Start MCP Server**
3. The HTTP API is now available at `http://localhost:8765`

### FastMCP Bridge (for Claude Code)

```bash
# Install FastMCP
pip install fastmcp httpx

# Configure Claude Code
claude mcp add libreoffice -- fastmcp run /path/to/libreoffice_mcp_server.py
```

## 🎯 Quick Start

### Using with Claude Code

Once configured, you can use natural language commands:

```
"Get document info"
→ document(action="info")

"Find all occurrences of 'hello'"
→ search(action="find", query="hello")

"Enable track changes"
→ track_changes(action="enable")

"Go to paragraph 5 and get context"
→ cursor(action="goto_paragraph", n=5)
→ cursor(action="context")
```

### HTTP API Examples

```bash
# Check server status
curl http://localhost:8765/health

# Get document info
curl -X POST http://localhost:8765/tools/get_document_info_live -d '{}'

# Find text
curl -X POST http://localhost:8765/tools/find_text_live \
  -H "Content-Type: application/json" \
  -d '{"query": "hello"}'

# Enable Track Changes
curl -X POST http://localhost:8765/tools/set_track_changes_live \
  -H "Content-Type: application/json" \
  -d '{"enabled": true}'
```

## 📂 Repository Structure

```
mcp-libre/
├── plugin/                    # LibreOffice extension
│   ├── pythonpath/
│   │   ├── uno_bridge.py     # UNO API wrapper
│   │   └── mcp_server.py     # HTTP API server
│   ├── build.sh              # Build script
│   └── README.md             # Plugin documentation
├── libreoffice_mcp_server.py # FastMCP bridge (9 consolidated tools)
├── docs/
│   ├── TOOL_REFERENCE.md     # Complete tool documentation
│   └── KNOWN_ISSUES_AND_ROADMAP.md
├── src/                      # Legacy external server
├── tests/                    # Test suite
└── README.md
```

## 📚 Documentation

- **[Tool Reference](docs/TOOL_REFERENCE.md)**: Complete documentation of all 9 tools
- **[Plugin README](plugin/README.md)**: LibreOffice extension details
- **[Known Issues & Roadmap](docs/KNOWN_ISSUES_AND_ROADMAP.md)**: Future plans

## 🆕 Recent Changes

### v0.3.0 - Tool Consolidation
- Consolidated 32 tools into 9 logical groups
- Reduced permission prompts for better UX
- Each tool uses `action` parameter for routing

### v0.2.0 - Track Changes Awareness
- Added 7 Track Changes management tools
- Search/replace now skips tracked deletions
- `get_paragraph` returns `visible_content` field
- `find_text` returns `track_changes_active` field

### v0.1.0 - Initial Release
- 25 individual MCP tools
- HTTP API on localhost:8765
- LibreOffice extension with UNO API integration

## 🛡 Security

- **Local Execution**: All operations run locally
- **File Permissions**: Limited to user's file access
- **No Network**: No external network dependencies
- **Email Privacy**: Uses GitHub noreply email for commits

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🔗 Links

- **Repository**: https://github.com/jwingnut/mcp-libre
- **MCP Specification**: https://spec.modelcontextprotocol.io/
- **LibreOffice**: https://www.libreoffice.org/
- **FastMCP**: https://github.com/modelcontextprotocol/python-sdk

---

*LibreOffice MCP Server v0.3.0 - AI-Powered Document Editing*
